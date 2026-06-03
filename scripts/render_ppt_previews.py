"""PPTX → PDF → 슬라이드별 PNG 변환.

GitHub에서 PPT 변경을 즉시 확인할 수 있도록 미리보기 자산을 생성한다.
- PDF: GitHub가 인라인 렌더링 지원
- PNG: PR 본문/README에 임베드해 슬라이드 단위 미리보기 제공

요구 사항(런타임): libreoffice, poppler-utils (pdftoppm)
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def require(cmd: str) -> str:
    path = shutil.which(cmd)
    if not path:
        sys.exit(f"필수 명령을 찾을 수 없습니다: {cmd}")
    return path


def pptx_to_pdf(pptx: Path, out_dir: Path) -> Path:
    require("libreoffice")
    out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(pptx),
        ],
        check=True,
    )
    pdf = out_dir / (pptx.stem + ".pdf")
    if not pdf.exists():
        sys.exit(f"PDF 변환 실패: {pdf}")
    return pdf


def pdf_to_png(pdf: Path, out_dir: Path, dpi: int = 110) -> list[Path]:
    require("pdftoppm")
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / "slide"
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(prefix)],
        check=True,
    )
    return sorted(out_dir.glob("slide-*.png"))


def main() -> None:
    parser = argparse.ArgumentParser(description="PPT 미리보기(PDF/PNG) 생성")
    parser.add_argument("--input", required=True, help="입력 PPTX 경로")
    parser.add_argument("--out", required=True, help="미리보기 출력 디렉터리")
    parser.add_argument("--dpi", type=int, default=110)
    args = parser.parse_args()

    pptx = Path(args.input)
    out_dir = Path(args.out)
    if not pptx.exists():
        sys.exit(f"입력 PPT를 찾을 수 없습니다: {pptx}")

    # 출력 폴더 초기화 (오래된 슬라이드 잔재 제거)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    pdf = pptx_to_pdf(pptx, out_dir)
    pngs = pdf_to_png(pdf, out_dir, dpi=args.dpi)

    # README.md 생성: 미리보기 인덱스
    readme = out_dir / "README.md"
    lines = [
        f"# 미리보기 — {pptx.name}",
        "",
        f"- 원본: [`{pptx.as_posix()}`](../{pptx.as_posix()})",
        f"- PDF: [`{pdf.name}`](./{pdf.name})",
        "",
        "## 슬라이드",
        "",
    ]
    for i, png in enumerate(pngs, start=1):
        lines.append(f"### Slide {i}")
        lines.append("")
        lines.append(f"![slide {i}](./{png.name})")
        lines.append("")
    readme.write_text("\n".join(lines), encoding="utf-8")

    print(f"[render_ppt_previews] PDF={pdf.name}, PNG={len(pngs)}장 생성 → {out_dir}")


if __name__ == "__main__":
    main()
