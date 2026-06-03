"""releases/ 폴더를 스캔해 인덱스(README.md) 및 루트 README 의
'최근 업데이트' 섹션을 자동 갱신.

- releases/<YYYYMMDD-HHMMSS>-<stem>/ 구조를 가정
- 각 릴리즈 폴더 안에 .pptx, .pdf, slide-*.png, README.md 가 존재
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from urllib.parse import quote

ROOT_MARK_START = "<!-- RELEASES:START -->"
ROOT_MARK_END = "<!-- RELEASES:END -->"


def list_releases(releases_dir: Path) -> list[Path]:
    if not releases_dir.exists():
        return []
    # 폴더 이름이 YYYYMMDD-HHMMSS 로 시작하는 것만, 최신 순
    items = [p for p in releases_dir.iterdir() if p.is_dir() and re.match(r"^\d{8}-\d{6}", p.name)]
    return sorted(items, reverse=True)


def _link(label: str, path: str) -> str:
    return f"[{label}](./{quote(path, safe='/')})"


def render_release_row(rel_dir: Path, base_dir: Path) -> str:
    pptx = next(rel_dir.glob("*.pptx"), None)
    pdf = next(rel_dir.glob("*.pdf"), None)
    pngs = sorted(rel_dir.glob("slide-*.png"))
    readme = rel_dir / "README.md"

    ts_match = re.match(r"^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})", rel_dir.name)
    if ts_match:
        y, mo, d, h, mi, s = ts_match.groups()
        ts_display = f"{y}-{mo}-{d} {h}:{mi}:{s} KST"
    else:
        ts_display = rel_dir.name

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


def write_releases_index(releases_dir: Path) -> None:
    releases = list_releases(releases_dir)
    body = [
        "# 📦 자동 생성 릴리즈",
        "",
        "각 폴더는 워크플로 실행 시각(KST)을 이름으로 가지며, 그 시점의 PPTX·PDF·슬라이드 PNG·미리보기 README 가 함께 저장됩니다.",
        "",
        build_table(releases, base_dir=releases_dir),
        "",
    ]
    (releases_dir / "README.md").write_text("\n".join(body), encoding="utf-8")


def update_root_readme(readme_path: Path, releases_dir: Path, repo_root: Path, limit: int) -> bool:
    if not readme_path.exists():
        return False
    text = readme_path.read_text(encoding="utf-8")
    if ROOT_MARK_START not in text or ROOT_MARK_END not in text:
        return False
    releases = list_releases(releases_dir)
    table = build_table(releases, base_dir=repo_root, limit=limit)
    new_section = f"{ROOT_MARK_START}\n{table}\n\n전체 이력: [`releases/`](./releases/)\n{ROOT_MARK_END}"
    pattern = re.compile(
        re.escape(ROOT_MARK_START) + r"[\s\S]*?" + re.escape(ROOT_MARK_END)
    )
    new_text = pattern.sub(new_section, text)
    if new_text != text:
        readme_path.write_text(new_text, encoding="utf-8")
        return True
    return False


def main() -> None:
    ap = argparse.ArgumentParser(description="릴리즈 인덱스/README 갱신")
    ap.add_argument("--releases", default="releases", help="릴리즈 루트 (기본: releases)")
    ap.add_argument("--readme", default="README.md", help="갱신 대상 루트 README")
    ap.add_argument("--limit", type=int, default=10, help="루트 README 에 노출할 최신 항목 수")
    args = ap.parse_args()

    repo_root = Path(".").resolve()
    releases_dir = Path(args.releases)
    releases_dir.mkdir(exist_ok=True)
    write_releases_index(releases_dir)
    changed = update_root_readme(Path(args.readme), releases_dir, repo_root, args.limit)
    print(
        f"[update_releases_index] releases={len(list_releases(releases_dir))}, "
        f"root_readme_updated={changed}"
    )


if __name__ == "__main__":
    main()
