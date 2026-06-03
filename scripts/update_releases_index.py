"""releases/ 폴더를 스캔해 인덱스(README.md) 및 루트 README 의
'최근 업데이트' / '최신 변경 내용' 섹션을 자동 갱신.

- releases/<YYYYMMDD-HHMMSS>-<stem>/ 구조를 가정
- 각 릴리즈 폴더 안에 .pptx, .pdf, slide-*.png, README.md, summary.md(선택) 가 존재
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from urllib.parse import quote

ROOT_TABLE_START = "<!-- RELEASES:START -->"
ROOT_TABLE_END = "<!-- RELEASES:END -->"
ROOT_LATEST_START = "<!-- LATEST:START -->"
ROOT_LATEST_END = "<!-- LATEST:END -->"
SLIDE_SORT_FALLBACK = sys.maxsize


def list_releases(releases_dir: Path) -> list[Path]:
    if not releases_dir.exists():
        return []
    items = [p for p in releases_dir.iterdir() if p.is_dir() and re.match(r"^\d{8}-\d{6}", p.name)]
    return sorted(items, reverse=True)


def _enc(path: str) -> str:
    return quote(path, safe="/")


def _link(label: str, path: str) -> str:
    return f"[{label}](./{_enc(path)})"


def _slide_sort_key(path: Path) -> tuple[int, str]:
    try:
        return (int(path.stem.split("-")[-1]), path.name)
    except ValueError:
        return (SLIDE_SORT_FALLBACK, path.name)


def _ts_display(name: str) -> str:
    m = re.match(r"^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})", name)
    if not m:
        return name
    y, mo, d, h, mi, s = m.groups()
    return f"{y}-{mo}-{d} {h}:{mi}:{s} KST"


def render_release_row(rel_dir: Path, base_dir: Path) -> str:
    pptx = next(rel_dir.glob("*.pptx"), None)
    pdf = next(rel_dir.glob("*.pdf"), None)
    pngs = sorted(rel_dir.glob("slide-*.png"), key=_slide_sort_key)
    readme = rel_dir / "README.md"

    ts_display = _ts_display(rel_dir.name)
    try:
        rel_path = rel_dir.resolve().relative_to(base_dir.resolve()).as_posix()
    except ValueError:
        rel_path = rel_dir.as_posix()

    links: list[str] = []
    if readme.exists():
        links.append(_link("🖼 미리보기", f"{rel_path}/README.md"))
    if pdf:
        links.append(_link("📄 PDF", f"{rel_path}/{pdf.name}"))
    if pptx:
        links.append(_link("📊 PPTX", f"{rel_path}/{pptx.name}"))

    title = pptx.stem if pptx else rel_dir.name
    return f"| `{ts_display}` | {title} | {len(pngs)}장 | {' · '.join(links)} |"


def build_table(releases: list[Path], base_dir: Path, limit: int | None = None) -> str:
    header = (
        "| 시각 (KST) | 자료 | 슬라이드 | 보기 |\n"
        "|------------|------|----------|------|"
    )
    rows = [render_release_row(r, base_dir) for r in (releases[:limit] if limit else releases)]
    if not rows:
        return "_아직 자동 생성된 업데이트가 없습니다. 워크플로를 한 번 실행해 주세요._"
    return header + "\n" + "\n".join(rows)


def build_latest_section(latest: Path | None, base_dir: Path) -> str:
    """최신 릴리즈의 요약 본문과 '추가된 슬라이드' 이미지를 인라인 마크다운으로."""
    if latest is None:
        return "_아직 업데이트가 없습니다._"

    pptx = next(latest.glob("*.pptx"), None)
    pdf = next(latest.glob("*.pdf"), None)
    pngs = sorted(latest.glob("slide-*.png"), key=_slide_sort_key)
    summary = latest / "summary.md"
    try:
        rel_path = latest.resolve().relative_to(base_dir.resolve()).as_posix()
    except ValueError:
        rel_path = latest.as_posix()

    title = pptx.stem if pptx else latest.name
    ts_display = _ts_display(latest.name)

    parts: list[str] = []
    parts.append(f"**🆕 {ts_display}** — `{title}`")
    parts.append("")
    nav: list[str] = []
    if (latest / "README.md").exists():
        nav.append(_link("🖼 전체 슬라이드 미리보기", f"{rel_path}/README.md"))
    if pdf:
        nav.append(_link("📄 PDF", f"{rel_path}/{pdf.name}"))
    if pptx:
        nav.append(_link("📊 PPTX 다운로드", f"{rel_path}/{pptx.name}"))
    if summary.exists():
        nav.append(_link("📝 요약(summary.md)", f"{rel_path}/summary.md"))
    if nav:
        parts.append(" · ".join(nav))
        parts.append("")

    # 요약 본문 인라인
    if summary.exists():
        body = summary.read_text(encoding="utf-8").strip()
        if body:
            parts.append("#### ✍️ LLM 요약")
            parts.append("")
            parts.append("<blockquote>")
            parts.append("")
            parts.append(body)
            parts.append("")
            parts.append("</blockquote>")
            parts.append("")

    # 추가된 슬라이드(마지막 슬라이드 = Latest Azure Updates) 인라인 표시
    if pngs:
        last_png = pngs[-1]
        parts.append("#### 🆕 추가된 슬라이드")
        parts.append("")
        parts.append(f"![최신 추가 슬라이드](./{_enc(rel_path + '/' + last_png.name)})")
        parts.append("")
        if len(pngs) > 1:
            parts.append("<details><summary>📑 전체 슬라이드 펼쳐보기</summary>")
            parts.append("")
            for i, png in enumerate(pngs, start=1):
                parts.append(f"**Slide {i}**")
                parts.append("")
                parts.append(f"![slide {i}](./{_enc(rel_path + '/' + png.name)})")
                parts.append("")
            parts.append("</details>")
            parts.append("")

    return "\n".join(parts)


def write_release_readme(rel_dir: Path) -> None:
    """각 릴리즈 폴더 README — 추가 슬라이드를 맨 앞으로, summary 인라인."""
    pptx = next(rel_dir.glob("*.pptx"), None)
    pdf = next(rel_dir.glob("*.pdf"), None)
    pngs = sorted(rel_dir.glob("slide-*.png"), key=_slide_sort_key)
    summary = rel_dir / "summary.md"

    title = pptx.stem if pptx else rel_dir.name
    lines = [
        f"# 📦 {title} — {_ts_display(rel_dir.name)}",
        "",
    ]
    nav = []
    if pdf:
        nav.append(f"[📄 PDF](./{_enc(pdf.name)})")
    if pptx:
        nav.append(f"[📊 PPTX 다운로드](./{_enc(pptx.name)})")
    if summary.exists():
        nav.append(f"[📝 summary.md](./summary.md)")
    if (rel_dir / "updates.json").exists():
        nav.append("[🔗 updates.json](./updates.json)")
    if nav:
        lines.append(" · ".join(nav))
        lines.append("")

    if summary.exists():
        body = summary.read_text(encoding="utf-8").strip()
        if body:
            lines += ["## ✍️ LLM 요약", "", body, ""]

    if pngs:
        lines += [
            "## 🆕 추가된 슬라이드 (Latest Azure Updates)",
            "",
            f"![추가된 슬라이드](./{_enc(pngs[-1].name)})",
            "",
            "## 📑 전체 슬라이드",
            "",
        ]
        for i, png in enumerate(pngs, start=1):
            lines += [f"### Slide {i}", "", f"![slide {i}](./{_enc(png.name)})", ""]

    (rel_dir / "README.md").write_text("\n".join(lines), encoding="utf-8")


def write_releases_index(releases_dir: Path) -> None:
    releases = list_releases(releases_dir)
    body = [
        "# 📦 자동 생성 릴리즈",
        "",
        "각 폴더는 워크플로 실행 시각(KST)을 이름으로 가지며, 그 시점의 PPTX·PDF·슬라이드 PNG·요약(summary.md) 이 함께 저장됩니다.",
        "",
        build_table(releases, base_dir=releases_dir),
        "",
    ]
    (releases_dir / "README.md").write_text("\n".join(body), encoding="utf-8")


def _replace_block(text: str, start: str, end: str, new_body: str) -> str:
    pattern = re.compile(re.escape(start) + r"[\s\S]*?" + re.escape(end))
    new_block = f"{start}\n{new_body}\n{end}"
    return pattern.sub(new_block, text)


def update_root_readme(readme_path: Path, releases_dir: Path, repo_root: Path, limit: int) -> bool:
    if not readme_path.exists():
        return False
    text = readme_path.read_text(encoding="utf-8")
    if ROOT_TABLE_START not in text:
        return False

    releases = list_releases(releases_dir)
    latest = releases[0] if releases else None

    table = build_table(releases, base_dir=repo_root, limit=limit) + \
            f"\n\n전체 이력: [`releases/`](./releases/)"
    new_text = _replace_block(text, ROOT_TABLE_START, ROOT_TABLE_END, table)

    if ROOT_LATEST_START in new_text:
        new_text = _replace_block(
            new_text, ROOT_LATEST_START, ROOT_LATEST_END,
            build_latest_section(latest, repo_root),
        )

    if new_text != text:
        readme_path.write_text(new_text, encoding="utf-8")
        return True
    return False


def main() -> None:
    ap = argparse.ArgumentParser(description="릴리즈 인덱스/README 갱신")
    ap.add_argument("--releases", default="releases")
    ap.add_argument("--readme", default="README.md")
    ap.add_argument("--limit", type=int, default=10)
    args = ap.parse_args()

    repo_root = Path(".").resolve()
    releases_dir = Path(args.releases)
    releases_dir.mkdir(exist_ok=True)

    # 각 릴리즈 폴더 README 갱신 (추가 슬라이드 강조)
    for r in list_releases(releases_dir):
        write_release_readme(r)

    write_releases_index(releases_dir)
    changed = update_root_readme(Path(args.readme), releases_dir, repo_root, args.limit)
    print(
        f"[update_releases_index] releases={len(list_releases(releases_dir))}, "
        f"root_readme_updated={changed}"
    )


if __name__ == "__main__":
    main()
