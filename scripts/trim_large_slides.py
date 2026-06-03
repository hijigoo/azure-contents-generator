"""슬라이드별 미디어 사용량을 추정해 큰 슬라이드를 제거한다."""
from __future__ import annotations
import argparse, os, re, sys, zipfile
from collections import defaultdict
from pathlib import Path
from pptx import Presentation


def measure(path: Path) -> list[tuple[int, int]]:
    media_sizes: dict[str, int] = {}
    slide_media: dict[int, set[str]] = defaultdict(set)
    rels_re = re.compile(r'Target="\.\./media/([^"]+)"')
    slide_rels_re = re.compile(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$")
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            if info.filename.startswith("ppt/media/"):
                media_sizes[os.path.basename(info.filename)] = info.file_size
        for info in z.infolist():
            m = slide_rels_re.match(info.filename)
            if not m:
                continue
            idx = int(m.group(1))
            data = z.read(info.filename).decode("utf-8", errors="ignore")
            for media in rels_re.findall(data):
                slide_media[idx].add(media)
    rows = [(idx, sum(media_sizes.get(m, 0) for m in medias)) for idx, medias in slide_media.items()]
    rows.sort(key=lambda r: -r[1])
    return rows


def remove_slides(pptx_path: Path, slide_numbers_1based: list[int]) -> None:
    prs = Presentation(pptx_path)
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    slide_ids = list(xml_slides)
    # slide N (1-based) maps to index N-1
    to_remove = sorted({n - 1 for n in slide_numbers_1based}, reverse=True)
    for i in to_remove:
        if 0 <= i < len(slide_ids):
            xml_slides.remove(slide_ids[i])
    prs.save(pptx_path)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--top", type=int, default=8, help="크기 상위 N개 슬라이드 제거")
    ap.add_argument("--min-mb", type=float, default=1.0, help="이 미만이면 제거 안 함")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    path = Path(args.input)
    if not path.exists():
        sys.exit(f"파일 없음: {path}")

    rows = measure(path)
    print("[측정] 상위 미디어 사용 슬라이드 (MB):")
    for idx, size in rows[: max(args.top, 10)]:
        print(f"  slide {idx:>3}: {size/1024/1024:6.2f} MB")

    targets = [idx for idx, size in rows[: args.top] if size / 1024 / 1024 >= args.min_mb]
    print(f"[제거 대상] {targets}")
    if args.dry_run or not targets:
        return
    remove_slides(path, targets)
    print(f"[완료] {len(targets)}개 슬라이드 제거 → {path}")


if __name__ == "__main__":
    main()
