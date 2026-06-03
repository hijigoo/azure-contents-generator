"""PPTX에서 임베드된 동영상을 모두 제거한다.

- ppt/media/media*.* (mp4, mov, wmv, avi, m4v 등) 파일을 zip에서 빼낸다.
- 관련 [Content_Types].xml / _rels 항목은 그대로 두어도 PPT는 열린다
  (참조된 미디어가 없으면 PowerPoint가 빈 자리로 표시).
- 더 안전한 정리를 원하면 .rels 에서 video relationship 도 제거 가능.

용법:
    python scripts/strip_video.py --input INPUT.pptx [--output OUT.pptx]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

VIDEO_RE = re.compile(
    r"^ppt/media/(media\d+\.(?:mp4|mov|wmv|avi|m4v|mkv|webm))$",
    re.IGNORECASE,
)


def strip(src: Path, dst: Path) -> tuple[int, int]:
    removed = 0
    kept = 0
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if VIDEO_RE.match(item.filename):
                removed += 1
                print(f"  [strip] {item.filename}  ({item.file_size/1024/1024:.1f} MB)")
                continue
            zout.writestr(item, zin.read(item.filename))
            kept += 1
    return removed, kept


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--output", help="기본: 입력 파일 덮어쓰기")
    args = p.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"입력 파일을 찾을 수 없습니다: {src}")
    dst = Path(args.output) if args.output else src

    tmp = src.with_suffix(src.suffix + ".tmp")
    removed, kept = strip(src, tmp)

    if dst == src:
        # 원자적 교체
        shutil.move(str(tmp), str(src))
    else:
        shutil.move(str(tmp), str(dst))

    size_mb = dst.stat().st_size / 1024 / 1024
    print(f"[strip_video] 제거 {removed}개 / 유지 {kept}개 → {dst} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
