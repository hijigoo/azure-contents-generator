"""skills/pptx-local 규칙에 맞춰 기존 슬라이드를 in-place 갱신한다."""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.slide import Slide
from pptx.text.text import TextFrame

from update_ppt import dedupe_pptx_zip

AUTO_DATE = "2026-06-04"
AUTO_MARKER = f"[auto-update:{AUTO_DATE}]"
MAX_INSERT_SUMMARY_LINES = 3
DEFAULT_ANCHOR_SLIDE_IDX = 9


@dataclass
class Decision:
    action: str
    reason: str
    keywords: list[str]


def validate_skill_context(skill_dir: Path) -> None:
    for name in ("SKILL.md", "editing.md"):
        path = skill_dir / name
        if not path.exists():
            sys.exit(f"필수 skill 파일을 찾을 수 없습니다: {path}")
        path.read_text(encoding="utf-8")


def normalize_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "")
    return text.strip().lower()


def extract_shape_text(shape) -> str:
    chunks: list[str] = []
    if getattr(shape, "has_text_frame", False):
        text = shape.text_frame.text.strip()
        if text:
            chunks.append(text)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    chunks.append(text)
    return "\n".join(chunks)


def collect_slide_text(slide) -> str:
    chunks = [extract_shape_text(shape) for shape in slide.shapes]
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
        chunks.append(slide.notes_slide.notes_text_frame.text)
    return "\n".join(part for part in chunks if part)


def iter_text_frames(slide):
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            yield shape.text_frame
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame


def replace_in_text_frame(tf: TextFrame, old: str, new: str) -> int:
    changed = 0
    for para in tf.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
                changed += 1
    return changed


def add_bullet(slide, text: str) -> bool:
    candidates = []
    title_shape = slide.shapes.title
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        if title_shape is not None and shape.shape_id == title_shape.shape_id:
            continue
        tf = shape.text_frame
        para_count = len(tf.paragraphs)
        chars = sum(len(p.text or "") for p in tf.paragraphs)
        candidates.append((para_count, chars, tf))
    if not candidates:
        return False

    _, _, target = max(candidates, key=lambda x: (x[0], x[1]))
    p = target.add_paragraph()
    p.text = text
    p.level = 0
    return True


def pick_layout(prs: Presentation, anchor_idx: int):
    for layout in prs.slide_layouts:
        name = (layout.name or "").lower()
        if "title and content" in name or "content" in name:
            return layout
    anchor_layout = prs.slides[anchor_idx].slide_layout
    if anchor_layout:
        return anchor_layout
    return prs.slide_layouts[1]


def insert_slide_after(prs: Presentation, anchor_idx: int, layout):
    slide = prs.slides.add_slide(layout)
    sld_id_lst = prs.slides._sldIdLst  # noqa: SLF001
    new_id = sld_id_lst[-1]
    sld_id_lst.remove(new_id)
    sld_id_lst.insert(anchor_idx + 1, new_id)
    return slide


def clean_title(title: str) -> str:
    cleaned = re.sub(r"^\[[^\]]+\]\s*", "", title or "")
    return re.sub(r"^(Generally Available|Public Preview|Preview):\s*", "", cleaned).strip()


def split_summary(summary: str) -> list[str]:
    raw = re.sub(r"\s+", " ", summary or "").strip()
    parts = [s.strip(" .") for s in re.split(r"[.!?]\s+", raw) if s.strip()]
    return parts[:MAX_INSERT_SUMMARY_LINES] if parts else ["관련 서비스 업데이트가 발표되었습니다."]


def fill_insert_slide(slide, item: dict) -> None:
    title_text = clean_title(item.get("title", "Azure 업데이트"))
    published = item.get("published", "").strip()
    source = item.get("source", "").strip()
    link = item.get("link", "").strip()
    summary_lines = split_summary(item.get("summary", ""))

    title_shape = slide.shapes.title
    if title_shape and title_shape.has_text_frame:
        title_shape.text_frame.clear()
        title_shape.text_frame.paragraphs[0].text = title_text
    else:
        box = slide.shapes.add_textbox(457200, 365760, 8229600, 685800)
        box.text_frame.text = title_text

    body_frame = None
    for shape in slide.shapes:
        if not getattr(shape, "is_placeholder", False):
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        phf = shape.placeholder_format
        if phf.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.CONTENT):
            body_frame = shape.text_frame
            break
    if body_frame is None:
        box = slide.shapes.add_textbox(457200, 1280160, 8229600, 4069080)
        body_frame = box.text_frame
    body_frame.clear()
    body_frame.paragraphs[0].text = f"📅 발행일: {published[:16]}"
    p1 = body_frame.add_paragraph()
    p1.text = f"🏷 출처: {source}"
    for line in summary_lines:
        p = body_frame.add_paragraph()
        p.text = f"• {line}"
    p_link = body_frame.add_paragraph()
    p_link.text = f"🔗 {link}"


def find_anchor_slide(slide_texts: list[str], keywords: list[str], default_idx: int) -> int:
    best_idx = default_idx
    best_score = -1
    lowered_keywords = [normalize_text(k) for k in keywords]
    for idx, text in enumerate(slide_texts):
        norm = normalize_text(text)
        score = sum(1 for k in lowered_keywords if k and k in norm)
        if score > best_score:
            best_idx = idx
            best_score = score
    return best_idx


def decide(item: dict) -> Decision:
    title = item.get("title", "")
    if "Global PTU Reservations Are Now Region-Agnostic" in title:
        return Decision("REPLACE", "PTU 설명의 최신 정책 반영", ["ptu", "reserved capacity"])
    if "Azure Policy Coverage for Model Router in Foundry Models" in title:
        return Decision("AUGMENT", "모델 라우터 거버넌스 기능 보강", ["모델 라우터", "router"])
    if "Voice Live integration with Microsoft Foundry Agent Service" in title:
        return Decision("AUGMENT", "Foundry Agent Service GA 반영", ["agent service", "one-click publishing"])
    if "Code-first observability for Foundry Agents in VS Code" in title:
        return Decision("AUGMENT", "코드 중심 관찰성 기능 반영", ["운영 가시성", "agent 평가", "evaluators"])
    if "Agent kit for Azure Cosmos DB" in title:
        return Decision("AUGMENT", "에이전트 개발 키트 GA 반영", ["agent frameworks", "deploy custom-code"])
    if "Unified Model API for multi-model AI applications" in title:
        return Decision("INSERT", "기존 흐름에 직접 대응되는 본문 부족", ["models", "platform integrations"])
    return Decision("SKIP", "덱 주제와 직접 연관 낮음", [])


def apply_replace(slide, item: dict) -> bool:
    title = item.get("title", "")
    if "Global PTU Reservations Are Now Region-Agnostic" not in title:
        return False
    replaced = 0
    for tf in iter_text_frames(slide):
        replaced += replace_in_text_frame(tf, "비용 최적화된 PTU", "Global PTU 예약(리전 무관) 기반 비용 최적화")
        replaced += replace_in_text_frame(tf, "Reserved capacity", "Reserved capacity (Global PTU)")
    return replaced > 0


def apply_augment(slide, item: dict) -> bool:
    title = item.get("title", "")
    if "Azure Policy Coverage for Model Router in Foundry Models" in title:
        return add_bullet(slide, "Azure Policy로 Model Router 라우팅 기준을 중앙 통제 (Public Preview)")
    if "Voice Live integration with Microsoft Foundry Agent Service" in title:
        return add_bullet(slide, "Voice Live 연동이 Foundry Agent Service에서 GA로 제공")
    if "Code-first observability for Foundry Agents in VS Code" in title:
        return add_bullet(slide, "VS Code에서 에이전트 관찰·평가 루프를 코드 중심으로 실행 (Public Preview)")
    if "Agent kit for Azure Cosmos DB" in title:
        return add_bullet(slide, "Azure Cosmos DB Agent Kit GA로 데이터 모델/쿼리 모범 사례 내장")
    return False


def remove_previous_insert_slides(prs: Presentation) -> int:
    removed = 0
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    for idx in range(len(prs.slides) - 1, -1, -1):
        slide = prs.slides[idx]
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame:
            continue
        note_text = slide.notes_slide.notes_text_frame.text or ""
        lines = note_text.splitlines()
        if not lines:
            continue
        header = lines[0].strip()
        if header.startswith("[auto-update:") and "ACTION=INSERT" in header:
            xml_slides.remove(list(xml_slides)[idx])
            removed += 1
    return removed


def update_cover_if_slot_exists(prs: Presentation) -> bool:
    cover = prs.slides[0]
    changed = False
    date_pattern = re.compile(r"(updated\s+)(\d{4}-\d{2}-\d{2})", re.IGNORECASE)
    for tf in iter_text_frames(cover):
        for para in tf.paragraphs:
            for run in para.runs:
                if date_pattern.search(run.text):
                    run.text = date_pattern.sub(rf"\g<1>{AUTO_DATE}", run.text)
                    changed = True
    return changed


def write_notes(slide, action: str, reason: str, items: list[dict]) -> None:
    tf = slide.notes_slide.notes_text_frame
    tf.clear()
    tf.paragraphs[0].text = f"{AUTO_MARKER} ACTION={action}; reason={reason}; executed={AUTO_DATE}"
    for item in items:
        title = item.get("title", "").strip()
        published = item.get("published", "").strip()
        link = item.get("link", "").strip()
        p = tf.add_paragraph()
        p.text = f"- {title} | {published} | {link}"


def main() -> None:
    parser = argparse.ArgumentParser(description="기존 슬라이드 in-place 자동 갱신")
    parser.add_argument("--input", required=True)
    parser.add_argument("--updates", required=True)
    parser.add_argument("--skill", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--insert-limit", type=int, default=2)
    args = parser.parse_args()

    input_path = Path(args.input)
    updates_path = Path(args.updates)
    skill_dir = Path(args.skill)
    output_path = Path(args.output)

    if not input_path.exists():
        sys.exit(f"입력 PPT를 찾을 수 없습니다: {input_path}")
    if not updates_path.exists():
        sys.exit(f"업데이트 JSON을 찾을 수 없습니다: {updates_path}")
    if input_path.resolve() == output_path.resolve():
        sys.exit("samples 원본 덮어쓰기는 허용되지 않습니다. --output 에 새 경로를 지정하세요.")

    validate_skill_context(skill_dir)
    data = json.loads(updates_path.read_text(encoding="utf-8"))
    items = data.get("items", [])

    prs = Presentation(input_path)
    if len(prs.slides) == 0:
        sys.exit("프레젠테이션에 슬라이드가 없어 업데이트할 수 없습니다.")
    removed_inserts = remove_previous_insert_slides(prs)
    update_cover_if_slot_exists(prs)

    slide_texts = [collect_slide_text(slide) for slide in prs.slides]
    changed_items_by_slide: dict[int, list[dict]] = defaultdict(list)
    action_meta: dict[int, tuple[Slide, str, str]] = {}
    inserted = 0
    counts = defaultdict(int)

    for item in items:
        decision = decide(item)
        if decision.action == "SKIP":
            counts["SKIP"] += 1
            continue

        anchor_idx = find_anchor_slide(
            slide_texts,
            decision.keywords,
            default_idx=min(DEFAULT_ANCHOR_SLIDE_IDX, len(prs.slides) - 1),
        )

        if decision.action == "REPLACE":
            target_slide = prs.slides[anchor_idx]
            if apply_replace(target_slide, item):
                counts["REPLACE"] += 1
                key = id(target_slide)
                changed_items_by_slide[key].append(item)
                action_meta[key] = (target_slide, "REPLACE", decision.reason)
            else:
                counts["SKIP"] += 1
            continue

        if decision.action == "AUGMENT":
            target_slide = prs.slides[anchor_idx]
            if apply_augment(target_slide, item):
                counts["AUGMENT"] += 1
                key = id(target_slide)
                changed_items_by_slide[key].append(item)
                action_meta[key] = (target_slide, "AUGMENT", decision.reason)
            else:
                counts["SKIP"] += 1
            continue

        if decision.action == "INSERT":
            if inserted >= args.insert_limit:
                counts["SKIP"] += 1
                continue
            layout = pick_layout(prs, anchor_idx)
            inserted_slide = insert_slide_after(prs, anchor_idx, layout)
            fill_insert_slide(inserted_slide, item)
            key = id(inserted_slide)
            changed_items_by_slide[key].append(item)
            action_meta[key] = (inserted_slide, "INSERT", decision.reason)
            inserted += 1
            counts["INSERT"] += 1
            new_idx = anchor_idx + 1
            slide_texts.insert(new_idx, collect_slide_text(inserted_slide))
            continue

    for key, changed_items in changed_items_by_slide.items():
        slide, action, reason = action_meta[key]
        write_notes(slide, action, reason, changed_items)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    deduped = dedupe_pptx_zip(output_path)
    print(
        "[agent_apply_updates] "
        f"REPLACE={counts['REPLACE']} AUGMENT={counts['AUGMENT']} INSERT={counts['INSERT']} SKIP={counts['SKIP']} "
        f"removed_insert_slides={removed_inserts} deduped={deduped} output={output_path}"
    )


if __name__ == "__main__":
    main()
