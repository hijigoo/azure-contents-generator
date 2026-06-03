"""skills/pptx-local 가이드를 직접 읽고 PPT에 최신 업데이트 슬라이드를 추가한다."""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import zipfile
from datetime import UTC, datetime
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

SLIDE_TITLE = "Latest Azure Updates"
PREFERRED_ITEMS = [
    "Code-first observability for Foundry Agents in VS Code",
    "Agent kit for Azure Cosmos DB",
    "Unified Model API for multi-model AI applications",
    "Azure Policy Coverage for Model Router in Foundry Models",
    "Voice Live integration with Microsoft Foundry Agent Service",
    "Global PTU Reservations Are Now Region-Agnostic",
]
FALLBACK_BULLETS = {
    "Code-first observability for Foundry Agents in VS Code": (
        "Foundry Agents in VS Code",
        "코드 중심 관찰성으로 평가·개선 루프를 에디터에서 실행 (Azure Updates)",
    ),
    "Agent kit for Azure Cosmos DB": (
        "Azure Cosmos DB Agent Kit",
        "AI 코딩 에이전트에 데이터 모델·쿼리 모범 사례를 내장 (Azure Updates)",
    ),
    "Unified Model API for multi-model AI applications": (
        "Azure API Management Unified Model API",
        "여러 모델 API 형식을 하나로 묶어 교체·거버넌스를 단순화 (Azure Updates)",
    ),
    "Azure Policy Coverage for Model Router in Foundry Models": (
        "Foundry Models Model Router Policy",
        "모델 라우팅 기준을 Azure Policy로 중앙 통제 (Azure Updates)",
    ),
    "Voice Live integration with Microsoft Foundry Agent Service": (
        "Voice Live + Foundry Agent Service",
        "음성 입출력을 별도 오디오 파이프라인 없이 바로 연결 (Azure Updates)",
    ),
    "Global PTU Reservations Are Now Region-Agnostic": (
        "Global PTU Reservations",
        "단일 예약으로 여러 리전의 Global PTU 사용량을 함께 최적화 (Azure Updates)",
    ),
}


def validate_skill_context(skill_dir: Path) -> None:
    for name in ("SKILL.md", "editing.md"):
        path = skill_dir / name
        if not path.exists():
            sys.exit(f"필수 skill 파일을 찾을 수 없습니다: {path}")
        path.read_text(encoding="utf-8")


def slide_has_title_text(slide, title: str) -> bool:
    title_shape = slide.shapes.title
    if title_shape and title_shape.has_text_frame:
        if title_shape.text_frame.text.strip() == title:
            return True
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        if shape.text_frame.text.strip() == title:
            return True
    return False


def remove_existing(prs: Presentation, title: str) -> int:
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    removed = 0
    for idx, slide in enumerate(list(prs.slides)):
        if slide_has_title_text(slide, title):
            xml_slides.remove(list(xml_slides)[idx - removed])
            removed += 1
    return removed


def select_items(items: list[dict], limit: int) -> list[dict]:
    selected: list[dict] = []
    seen: set[str] = set()
    for needle in PREFERRED_ITEMS:
        for item in items:
            title = item.get("title", "")
            if needle in title and title not in seen:
                selected.append(item)
                seen.add(title)
                break
        if len(selected) >= limit:
            return selected[:limit]
    for item in items:
        title = item.get("title", "")
        if title and title not in seen:
            selected.append(item)
            seen.add(title)
        if len(selected) >= limit:
            break
    return selected[:limit]


def parse_summary(summary_path: Path | None) -> list[tuple[str, str]]:
    if not summary_path or not summary_path.exists():
        return []
    bullets: list[tuple[str, str]] = []
    for raw_line in summary_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line[:1] in "-*•":
            line = re.sub(r"^[-*•]\s*", "", line)
            match = re.match(r"\*\*(.+?)\*\*\s*(.*)", line)
            if match:
                bullets.append((match.group(1).strip(), match.group(2).strip()))
            else:
                name, _, desc = line.partition(" ")
                bullets.append((name.strip(), desc.strip()))
    return bullets


def build_bullets(items: list[dict], summary_path: Path | None) -> list[tuple[str, str]]:
    parsed = parse_summary(summary_path)
    if parsed:
        return parsed

    bullets: list[tuple[str, str]] = []
    for item in items:
        title = item.get("title", "")
        matched = False
        for needle, bullet in FALLBACK_BULLETS.items():
            if needle in title:
                bullets.append(bullet)
                matched = True
                break
        if not matched:
            cleaned = re.sub(r"^\[[^\]]+\]\s*", "", title)
            cleaned = re.sub(r"^(Generally Available|Public Preview|Preview):\s*", "", cleaned)
            bullets.append((cleaned, "최신 Azure 공지에 반영된 기능 업데이트입니다. (Azure Updates)"))
    return bullets


def add_updates_slide(prs: Presentation, bullets: list[tuple[str, str]], items: list[dict]) -> None:
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)

    title_box = slide.shapes.add_textbox(Inches(0.64), Inches(0.42), Inches(11.4), Inches(0.7))
    title_tf = title_box.text_frame
    title_tf.clear()
    title_run = title_tf.paragraphs[0].add_run()
    title_run.text = SLIDE_TITLE
    title_run.font.bold = True
    title_run.font.size = Pt(30)

    body_box = slide.shapes.add_textbox(Inches(0.78), Inches(1.35), Inches(11.1), Inches(4.85))
    body_tf = body_box.text_frame
    body_tf.clear()
    body_tf.word_wrap = True

    for idx, (name, desc) in enumerate(bullets):
        paragraph = body_tf.paragraphs[0] if idx == 0 else body_tf.add_paragraph()
        paragraph.space_after = Pt(10)
        bullet_run = paragraph.add_run()
        bullet_run.text = "• "
        bullet_run.font.size = Pt(17)
        name_run = paragraph.add_run()
        name_run.text = name
        name_run.font.bold = True
        name_run.font.size = Pt(17)
        desc_run = paragraph.add_run()
        desc_run.text = f" {desc.strip()}" if desc else ""
        desc_run.font.size = Pt(17)

    notes_tf = slide.notes_slide.notes_text_frame
    notes_tf.text = f"Updated: {datetime.now(UTC).isoformat()}"
    for item in items:
        title = item.get("title", "").strip()
        published = item.get("published", "").strip()
        link = item.get("link", "").strip()
        paragraph = notes_tf.add_paragraph()
        paragraph.text = f"- {title} ({published}) {link}".strip()


def dedupe_pptx_zip(pptx_path: Path) -> int:
    with zipfile.ZipFile(pptx_path, "r") as zin:
        infos = zin.infolist()
        kept: dict[str, tuple[zipfile.ZipInfo, bytes]] = {}
        for info in infos:
            with zin.open(info) as handle:
                kept[info.filename] = (info, handle.read())
        removed = len(infos) - len(kept)

    if removed == 0:
        return 0

    temp_path = pptx_path.with_suffix(pptx_path.suffix + ".dedup")
    with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, (info, data) in kept.items():
            new_info = zipfile.ZipInfo(filename=name, date_time=info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, data)
    shutil.move(str(temp_path), str(pptx_path))
    return removed


def main() -> None:
    parser = argparse.ArgumentParser(description="직접 작성한 python-pptx 로 Azure 업데이트 슬라이드 추가")
    parser.add_argument("--input", required=True)
    parser.add_argument("--updates", required=True)
    parser.add_argument("--skill", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--summary")
    parser.add_argument("--limit", type=int, default=5)
    args = parser.parse_args()

    input_path = Path(args.input)
    updates_path = Path(args.updates)
    skill_dir = Path(args.skill)
    output_path = Path(args.output)
    summary_path = Path(args.summary) if args.summary else None

    if not input_path.exists():
        sys.exit(f"입력 PPT를 찾을 수 없습니다: {input_path}")
    if not updates_path.exists():
        sys.exit(f"업데이트 JSON을 찾을 수 없습니다: {updates_path}")
    if input_path.resolve() == output_path.resolve():
        sys.exit("samples 원본 덮어쓰기는 허용되지 않습니다. --output 에 새 경로를 지정하세요.")

    validate_skill_context(skill_dir)
    data = json.loads(updates_path.read_text(encoding="utf-8"))
    selected_items = select_items(data.get("items", []), args.limit)
    bullets = build_bullets(selected_items, summary_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation(input_path)
    before_count = len(prs.slides)
    removed = remove_existing(prs, SLIDE_TITLE)
    add_updates_slide(prs, bullets, selected_items)
    prs.save(output_path)
    deduped = dedupe_pptx_zip(output_path)
    print(
        f"[agent_apply_updates] {before_count}→{len(prs.slides)} slides "
        f"(removed={removed}, deduped={deduped}) → {output_path}"
    )


if __name__ == "__main__":
    main()
