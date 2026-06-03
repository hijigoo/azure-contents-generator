from __future__ import annotations

import argparse
import json
import re
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation
from pptx.slide import SlideLayout
from pptx.util import Pt

SLIDE_TITLE = "Latest Azure Updates"


def load_skill_context(skill_dir: Path) -> tuple[str, str]:
    skill_md = skill_dir / "SKILL.md"
    editing_md = skill_dir / "editing.md"
    return skill_md.read_text(encoding="utf-8"), editing_md.read_text(encoding="utf-8")


def remove_existing(prs: Presentation, title: str) -> int:
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    removed = 0
    for idx, slide in enumerate(list(prs.slides)):
        slide_title = slide.shapes.title
        if slide_title and slide_title.has_text_frame and slide_title.text_frame.text.strip() == title:
            xml_slides.remove(list(xml_slides)[idx])
            removed += 1
    return removed


def choose_layout(prs: Presentation) -> SlideLayout:
    for layout in prs.slide_layouts:
        name = (layout.name or "").strip().lower()
        if name in {"title and content", "제목 및 내용"}:
            return layout
    return prs.slide_layouts[1]


def split_title(title: str) -> tuple[str, str]:
    cleaned = re.sub(r"^\[[^\]]+\]\s*", "", title).strip()
    if ":" in cleaned:
        head, tail = cleaned.split(":", 1)
        return head.strip(), tail.strip()
    return cleaned, ""


def summarize_kr(text: str) -> str:
    s = re.sub(r"\s+", " ", text or "").strip()
    if not s:
        return "업데이트가 공개되었습니다."
    if len(s) > 58:
        s = s[:57].rstrip() + "…"
    return s


def populate_body(slide, items: list[dict], limit: int) -> None:
    body = None
    for shape in slide.placeholders:
        if shape.placeholder_format.idx != 0 and shape.has_text_frame:
            body = shape
            break
    if body is None:
        body = slide.shapes.add_textbox(Pt(36), Pt(108), Pt(620), Pt(320))

    tf = body.text_frame
    tf.clear()
    tf.word_wrap = True

    selected = items[:limit]
    if not selected:
        p = tf.paragraphs[0]
        p.text = "업데이트 항목이 없습니다."
        p.level = 0
        for run in p.runs:
            run.font.size = Pt(17)
        return

    for idx, item in enumerate(selected):
        title = item.get("title", "").strip()
        summary = item.get("summary", "").strip()
        product, desc = split_title(title)
        line_desc = summarize_kr(desc or summary)

        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.level = 0
        p.text = ""

        run_product = p.add_run()
        run_product.text = product
        run_product.font.bold = True
        run_product.font.size = Pt(17)

        run_desc = p.add_run()
        run_desc.text = f" — {line_desc} (Azure Updates)"
        run_desc.font.size = Pt(17)


def populate_notes(slide, items: list[dict], limit: int) -> None:
    notes = slide.notes_slide.notes_text_frame
    notes.clear()

    p0 = notes.paragraphs[0]
    p0.text = f"Generated: {datetime.now(timezone.utc).isoformat()}"

    for item in items[:limit]:
        title = item.get("title", "").strip()
        published = item.get("published", "").strip()
        link = item.get("link", "").strip()
        p = notes.add_paragraph()
        p.text = f"- {title} | published: {published} | link: {link}"


def dedupe_pptx_zip(pptx_path: Path) -> int:
    with zipfile.ZipFile(pptx_path, "r") as zin:
        infos = zin.infolist()
        kept: dict[str, tuple[zipfile.ZipInfo, bytes]] = {}
        for info in infos:
            with zin.open(info) as fh:
                kept[info.filename] = (info, fh.read())
        removed = len(infos) - len(kept)

    if removed == 0:
        return 0

    tmp = pptx_path.with_suffix(pptx_path.suffix + ".dedup")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, (info, data) in kept.items():
            new_info = zipfile.ZipInfo(filename=name, date_time=info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, data)
    shutil.move(str(tmp), str(pptx_path))
    return removed


def build_slide(input_pptx: Path, output_pptx: Path, updates_path: Path, limit: int, skill_dir: Path) -> tuple[int, int]:
    load_skill_context(skill_dir)

    data = json.loads(updates_path.read_text(encoding="utf-8"))
    items = data.get("items", [])

    prs = Presentation(input_pptx)
    removed = remove_existing(prs, SLIDE_TITLE)

    layout = choose_layout(prs)
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = SLIDE_TITLE

    if slide.shapes.title and slide.shapes.title.has_text_frame:
        for run in slide.shapes.title.text_frame.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(30)

    populate_body(slide, items, limit)
    populate_notes(slide, items, limit)

    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_pptx)
    deduped = dedupe_pptx_zip(output_pptx)
    return removed, deduped


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply latest Azure updates slide with python-pptx")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--updates", required=True)
    parser.add_argument("--skill-dir", default="skills/pptx-local")
    parser.add_argument("--limit", type=int, default=5)
    args = parser.parse_args()

    input_pptx = Path(args.input)
    output_pptx = Path(args.output)
    updates_path = Path(args.updates)
    skill_dir = Path(args.skill_dir)

    if not input_pptx.exists():
        raise SystemExit(f"Input not found: {input_pptx}")
    if not updates_path.exists():
        raise SystemExit(f"Updates not found: {updates_path}")
    if not skill_dir.exists():
        raise SystemExit(f"Skill dir not found: {skill_dir}")

    removed, deduped = build_slide(input_pptx, output_pptx, updates_path, args.limit, skill_dir)
    print(f"saved={output_pptx}")
    print(f"removed_existing={removed}")
    print(f"deduped_entries={deduped}")


if __name__ == "__main__":
    main()
