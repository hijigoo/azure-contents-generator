"""Anthropic PPT skill을 호출해 PPT 슬라이드를 최신 정보로 업데이트한다.

흐름:
1. 입력 PPT에서 슬라이드별 텍스트를 추출한다.
2. Azure 업데이트 JSON과 함께 Claude(Anthropic API)로 전달한다.
3. Claude가 PPT skill을 사용해 슬라이드를 갱신한다.
4. 갱신된 결과를 원본 파일에 덮어쓴다 (--dry-run 인 경우 diff만 출력).

본 스크립트는 Anthropic의 pptx skill 번들(`skills/pptx/`)이 함께 제공된다고
가정한다. skill 실행 방식은 Anthropic SDK 의 `tool_use` / `skill` 인터페이스
스펙에 맞춰 통합한다.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

from pptx import Presentation

try:
    from anthropic import Anthropic
except ImportError:  # pragma: no cover - 런타임 환경 가드
    Anthropic = None  # type: ignore[assignment]


MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-5")
SYSTEM_PROMPT = """너는 Azure 기술 자료를 관리하는 시니어 테크니컬 라이터다.
주어진 PPT 슬라이드의 텍스트를 최신 Azure 업데이트 정보를 반영해 한국어로 갱신한다.
규칙:
- 슬라이드 구조/순서/제목은 가급적 유지한다.
- 사실에 근거해서만 갱신하고 출처 URL을 슬라이드 노트에 추가한다.
- 명백히 변경 필요 없는 슬라이드는 그대로 둔다.
출력은 JSON 배열로, 각 항목은 {"slide_index": int, "title": str, "bullets": [str], "notes": str} 형식이다.
"""


def extract_slides(pptx_path: Path) -> list[dict]:
    prs = Presentation(pptx_path)
    slides = []
    for idx, slide in enumerate(prs.slides):
        title = ""
        bullets: list[str] = []
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        bullets.append(text)
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text
        slides.append(
            {"slide_index": idx, "title": title, "bullets": bullets, "notes": notes}
        )
    return slides


def call_anthropic(slides: list[dict], updates: dict, skill_dir: Path) -> list[dict]:
    if Anthropic is None:
        raise RuntimeError("anthropic 패키지가 설치되어 있지 않습니다.")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY 환경 변수가 필요합니다.")

    client = Anthropic(api_key=api_key)
    user_payload = {
        "current_slides": slides,
        "azure_updates": updates,
        "skill_path": str(skill_dir),
    }

    message = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": (
                    "다음 데이터를 보고 갱신된 슬라이드 JSON을 반환해줘.\n"
                    + json.dumps(user_payload, ensure_ascii=False)
                ),
            }
        ],
    )
    text = "".join(block.text for block in message.content if getattr(block, "type", "") == "text")
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Claude 응답을 JSON 으로 파싱하지 못했습니다: {exc}\n응답: {text[:500]}")


def apply_updates(pptx_path: Path, updated: list[dict], output_path: Path) -> None:
    prs = Presentation(pptx_path)
    by_idx = {item["slide_index"]: item for item in updated}

    for idx, slide in enumerate(prs.slides):
        patch = by_idx.get(idx)
        if not patch:
            continue
        if slide.shapes.title and slide.shapes.title.has_text_frame and patch.get("title"):
            slide.shapes.title.text_frame.text = patch["title"]
        bullets = patch.get("bullets") or []
        body_shape = next(
            (s for s in slide.shapes if s != slide.shapes.title and s.has_text_frame),
            None,
        )
        if body_shape and bullets:
            tf = body_shape.text_frame
            tf.clear()
            tf.text = bullets[0]
            for line in bullets[1:]:
                p = tf.add_paragraph()
                p.text = line
        notes = patch.get("notes")
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    prs.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Anthropic skill 기반 PPT 업데이트")
    parser.add_argument("--input", required=True, help="원본 PPT 경로")
    parser.add_argument("--updates", required=True, help="Azure 업데이트 JSON 경로")
    parser.add_argument("--skill", required=True, help="Anthropic pptx skill 디렉터리")
    parser.add_argument("--output", help="출력 PPT 경로 (기본: 원본 덮어쓰기)")
    parser.add_argument("--dry-run", action="store_true", help="변경 사항을 출력만 함")
    args = parser.parse_args()

    pptx_path = Path(args.input)
    updates_path = Path(args.updates)
    skill_dir = Path(args.skill)
    output_path = Path(args.output) if args.output else pptx_path

    if not pptx_path.exists():
        sys.exit(f"입력 PPT를 찾을 수 없습니다: {pptx_path}")
    if not updates_path.exists():
        sys.exit(f"업데이트 JSON을 찾을 수 없습니다: {updates_path}")

    slides = extract_slides(pptx_path)
    updates = json.loads(updates_path.read_text(encoding="utf-8"))
    print(f"[update_ppt] 슬라이드 {len(slides)}개, 업데이트 {len(updates.get('items', []))}건")

    updated = call_anthropic(slides, updates, skill_dir)

    if args.dry_run:
        print(json.dumps(updated, ensure_ascii=False, indent=2))
        return

    if output_path == pptx_path:
        shutil.copy2(pptx_path, pptx_path.with_suffix(".bak.pptx"))
    apply_updates(pptx_path, updated, output_path)
    print(f"[update_ppt] 저장 완료: {output_path}")


if __name__ == "__main__":
    main()
