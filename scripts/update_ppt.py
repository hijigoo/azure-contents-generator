"""Anthropic pptx skill의 오픈소스 스크립트들을 활용해
PPT에 'Latest Azure Updates' 슬라이드를 기계적으로 추가/갱신한다.

이 스크립트는 LLM/외부 API를 호출하지 않는다. 인증 불필요.

흐름:
1. .cache/updates.json (fetch_azure_updates.py 결과) 로드
2. 입력 PPT의 마지막에 동일 제목 슬라이드가 있으면 제거 (멱등성)
3. 새로운 'Latest Azure Updates' 슬라이드 1장을 추가
   - 항목 제목 + 발행일을 bullet 으로 표시
   - 노트(notes)에 링크/요약 첨부
4. (선택) anthropic/skills 의 clean.py 를 호출해 최종 정리
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

SLIDE_TITLE = "Latest Azure Updates"


def remove_existing(prs: Presentation, title: str) -> None:
    """제목이 일치하는 기존 슬라이드를 제거 (멱등 업데이트)."""
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    slides = list(prs.slides)
    for idx, slide in enumerate(slides):
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            if slide.shapes.title.text_frame.text.strip() == title:
                slide_id = list(xml_slides)[idx]
                xml_slides.remove(slide_id)


def add_updates_slide(prs: Presentation, items: list[dict], limit: int = 8, summary_text: str | None = None) -> None:
    layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = SLIDE_TITLE

    body = None
    for shape in slide.placeholders:
        if shape.placeholder_format.idx != 0 and shape.has_text_frame:
            body = shape
            break

    if body is None:
        # 레이아웃에 content placeholder가 없으면 새 textbox 사용
        body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))

    tf = body.text_frame
    tf.clear()
    tf.word_wrap = True

    body_lines: list[str] = []
    if summary_text:
        # LLM(코파일럿)이 작성한 한국어 요약을 그대로 본문에 사용
        body_lines.extend([ln for ln in summary_text.splitlines() if ln.strip()])
    else:
        for item in items[:limit]:
            title = item.get("title", "").strip()
            published = item.get("published", "").strip()
            line = f"• {title}"
            if published:
                line += f"  ({published[:10]})"
            body_lines.append(line)

    if not body_lines:
        body_lines = ["(no updates)"]

    tf.text = body_lines[0]
    for run in tf.paragraphs[0].runs:
        run.font.size = Pt(14)
    for line in body_lines[1:]:
        p = tf.add_paragraph()
        p.text = line
        for run in p.runs:
            run.font.size = Pt(14)

    # 노트에는 항상 원문 RSS 메타데이터를 기록 (출처 추적용)
    notes_lines = [f"Updated: {datetime.utcnow().isoformat()}Z", ""]
    for item in items[:limit]:
        title = item.get("title", "").strip()
        published = item.get("published", "").strip()
        link = item.get("link", "").strip()
        summary = item.get("summary", "").strip()
        notes_lines.append(f"- {title}")
        if published:
            notes_lines.append(f"  published: {published}")
        if link:
            notes_lines.append(f"  link: {link}")
        if summary:
            short = summary if len(summary) < 240 else summary[:237] + "..."
            notes_lines.append(f"  summary: {short}")
        notes_lines.append("")
    slide.notes_slide.notes_text_frame.text = "\n".join(notes_lines)


def maybe_run_skill_clean(pptx: Path, skill_dir: Path) -> None:
    """skill 의 clean.py 가 있으면 호출해 임시 자산을 정리."""
    clean_script = skill_dir / "scripts" / "clean.py"
    if not clean_script.exists():
        return
    try:
        subprocess.run(
            [sys.executable, str(clean_script), str(pptx)],
            check=True,
            capture_output=True,
        )
        print(f"[update_ppt] skill clean.py 실행 완료")
    except subprocess.CalledProcessError as exc:
        # clean 단계는 실패해도 본 업데이트는 유지
        print(f"[update_ppt] skill clean.py 실패(무시): {exc.stderr.decode(errors='ignore')[:200]}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Azure 업데이트 슬라이드 기계 삽입")
    parser.add_argument("--input", required=True, help="원본 PPT 경로 (덮어씀)")
    parser.add_argument("--updates", required=True, help="updates.json 경로")
    parser.add_argument("--skill", required=True, help="anthropics/skills/skills/pptx 디렉터리")
    parser.add_argument("--limit", type=int, default=8, help="삽입할 항목 수")
    parser.add_argument("--summary", help="LLM(코파일럿) 생성 요약 마크다운 경로 (없으면 RSS bullet 사용)")
    parser.add_argument("--output", help="출력 경로 (기본: 원본 덮어쓰기)")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    pptx_path = Path(args.input)
    updates_path = Path(args.updates)
    skill_dir = Path(args.skill)
    output_path = Path(args.output) if args.output else pptx_path

    if not pptx_path.exists():
        sys.exit(f"입력 PPT를 찾을 수 없습니다: {pptx_path}")
    if not updates_path.exists():
        sys.exit(f"업데이트 JSON을 찾을 수 없습니다: {updates_path}")

    data = json.loads(updates_path.read_text(encoding="utf-8"))
    items = data.get("items", [])
    print(f"[update_ppt] 항목 {len(items)}건 중 상위 {args.limit}건 반영")

    if args.dry_run:
        for it in items[: args.limit]:
            print(f"  - {it.get('title', '')} ({it.get('published', '')[:10]})")
        return

    prs = Presentation(pptx_path)
    remove_existing(prs, SLIDE_TITLE)
    summary_text = None
    if args.summary:
        sp = Path(args.summary)
        if sp.exists():
            summary_text = sp.read_text(encoding="utf-8")
            print(f"[update_ppt] LLM 요약 사용: {sp} ({len(summary_text)}자)")
    add_updates_slide(prs, items, limit=args.limit, summary_text=summary_text)
    prs.save(output_path)
    print(f"[update_ppt] 저장 완료: {output_path}")

    maybe_run_skill_clean(output_path, skill_dir)


if __name__ == "__main__":
    main()
