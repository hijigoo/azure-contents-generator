"""samples/*.pptx -> PPT별 폴더로 스크롤형 HTML 덱 빌드 + 버전 히스토리 관리.

출력 구조:
    docs/samples/
      index.html                  전체 PPT 목록
      <slug>/                     PPT 별 폴더
        latest.html               최신 버전 (안정 URL)
        index.html                이 PPT 의 버전 히스토리 목록
        v001-<YYYYMMDD-HHMMSS>.html ...  버전별 스냅샷 (보존)
        versions.json             버전 기록 (해시/시각/파일명)
        titles.json               (선택) 목차 타이틀 {"1":"제목",...}

버전 판단:
    PPTX 내용의 sha256 해시가 직전 버전과 다르면 **새 버전** 파일을 추가한다.
    과거 버전 HTML 은 그대로 보존된다(과거 원본 pptx 가 없으므로 재생성하지 않음).
    전체 히스토리는 각 PPT 폴더의 index.html 에서 항상 최신 상태로 확인할 수 있다.

순수 Python + python-pptx 만 사용한다. build_scroll_deck.py 의 추출/렌더 유틸을 재사용.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime
from html import escape
from pathlib import Path

from build_scroll_deck import (
    CSS,
    extract_deck,
    load_content,
    load_titles,
    resolve_title,
    _slide_section_html,
)


def slugify(name: str) -> str:
    s = name.strip()
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"[^0-9A-Za-z가-힣\-_]", "", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s or "deck"


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()[:12]


def load_versions(folder: Path) -> list[dict]:
    f = folder / "versions.json"
    if f.exists():
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def save_versions(folder: Path, versions: list[dict]) -> None:
    (folder / "versions.json").write_text(
        json.dumps(versions, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# --------------------------------------------------------------------------- #
# 렌더
# --------------------------------------------------------------------------- #
def render_version_page(deck_name: str, deck: list[dict], titles: dict,
                        version: dict, total_versions: int, content: dict) -> str:
    nav_items, sections = [], []
    for s in deck:
        title = resolve_title(s, titles)
        sid = f"slide-{s['num']}"
        nav_items.append(
            f'<li><a href="#{sid}" data-target="{sid}">'
            f'<span class="n">{s["num"]}</span><span>{escape(title)}</span></a></li>'
        )
        sections.append(_slide_section_html(s, title, len(deck), content))

    badge = f'v{version["n"]:03d} · {escape(version["ts_human"])}'
    return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(deck_name)} ({badge})</title>
<style>{CSS}</style></head><body>
<div class="topbar"><button onclick="document.querySelector('aside').classList.toggle('open')">&#9776;</button><span class="t">{escape(deck_name)}</span></div>
<div class="layout">
<aside>
<div class="deck-title">{escape(deck_name)}</div>
<div class="ver">{badge}<br><a href="index.html" style="color:var(--accent)">전체 버전 {total_versions}개 &rarr;</a></div>
<nav><ol>{''.join(nav_items)}</ol></nav>
<div class="ver" style="padding-top:16px"><a href="../index.html" style="color:var(--accent)">&larr; 전체 PPT 목록</a></div>
</aside>
<main>{''.join(sections)}</main>
</div>
<script>
const links=[...document.querySelectorAll('nav a')];
const map=new Map(links.map(a=>[a.dataset.target,a]));
const obs=new IntersectionObserver((es)=>{{
 es.forEach(e=>{{if(e.isIntersecting){{
  links.forEach(a=>a.classList.remove('active'));
  const a=map.get(e.target.id); if(a){{a.classList.add('active');
   a.scrollIntoView({{block:'nearest'}});}}
 }}}});
}},{{rootMargin:'-45% 0px -50% 0px'}});
document.querySelectorAll('.slide-wrap').forEach(s=>obs.observe(s));
links.forEach(a=>a.addEventListener('click',()=>document.querySelector('aside').classList.remove('open')));
</script>
</body></html>"""


def render_deck_index(deck_name: str, versions: list[dict]) -> str:
    rows = []
    for i, v in enumerate(reversed(versions)):
        tag = " <span style='color:var(--accent)'>(최신)</span>" if i == 0 else ""
        rows.append(
            f'<li><a href="{v["file"]}">v{v["n"]:03d}</a> '
            f'<span class="meta">· {escape(v["ts_human"])} · <code>{v["hash"]}</code></span>{tag}</li>'
        )
    return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(deck_name)} — 버전 히스토리</title>
<style>{CSS}
body{{padding:0;}} .wrap{{max-width:760px;margin:0 auto;padding:48px 24px;}}
h1{{font-size:24px;margin:0 0 4px;}} .sub{{color:var(--muted);margin-bottom:28px;}}
ul.vers{{list-style:none;margin:0;padding:0;}}
ul.vers li{{padding:10px 0;border-bottom:1px solid var(--border);}}
ul.vers a{{color:var(--accent);text-decoration:none;font-weight:600;}}
.meta{{color:var(--muted);font-size:13px;}}
code{{background:var(--side);padding:1px 6px;border-radius:6px;font-size:12px;}}
.cta{{display:inline-block;margin:18px 0;padding:10px 16px;background:var(--accent);
color:#fff;border-radius:10px;text-decoration:none;}}
</style></head><body>
<div class="wrap">
<h1>{escape(deck_name)}</h1>
<div class="sub">버전 히스토리 · 총 {len(versions)}개</div>
<a class="cta" href="latest.html">최신 버전 보기 &rarr;</a>
<ul class="vers">{''.join(rows)}</ul>
<p><a href="../index.html" style="color:var(--accent)">&larr; 전체 PPT 목록</a></p>
</div></body></html>"""


def render_root_index(decks: list[dict]) -> str:
    cards = []
    for d in sorted(decks, key=lambda x: x["name"]):
        latest = d["versions"][-1]
        cards.append(
            f'<article class="card"><h2><a href="{d["slug"]}/latest.html">{escape(d["name"])}</a></h2>'
            f'<div class="meta">버전 {len(d["versions"])}개 · 최신 {escape(latest["ts_human"])}</div>'
            f'<a href="{d["slug"]}/index.html" style="color:var(--accent);font-size:14px">버전 히스토리 &rarr;</a></article>'
        )
    return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>샘플 덱 목록</title>
<style>{CSS}
body{{padding:0;}} .wrap{{max-width:880px;margin:0 auto;padding:48px 24px;}}
h1{{font-size:28px;margin:0 0 8px;}} .sub{{color:var(--muted);margin-bottom:32px;}}
.card{{border:1px solid var(--border);border-radius:14px;padding:20px 24px;margin:16px 0;background:var(--card);}}
.card h2{{margin:0 0 4px;font-size:20px;}} .card h2 a{{color:var(--fg);text-decoration:none;}}
.card .meta{{color:var(--muted);font-size:13px;margin-bottom:10px;}}
</style></head><body>
<div class="wrap">
<h1>샘플 덱</h1>
<div class="sub">samples/ 의 PPTX 내용을 HTML 로 변환. PPT 별 폴더에서 최신/히스토리 관리.</div>
{''.join(cards) or '<p class="empty">변환할 PPTX 가 없습니다.</p>'}
</div></body></html>"""


# --------------------------------------------------------------------------- #
# 빌드
# --------------------------------------------------------------------------- #
def build(samples_dir: Path, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_files = sorted(p for p in samples_dir.glob("*.pptx"))
    if not pptx_files:
        print(f"PPTX 없음: {samples_dir}")
        return

    decks = []
    now = datetime.now()
    ts_key = now.strftime("%Y%m%d-%H%M%S")
    ts_human = now.strftime("%Y-%m-%d %H:%M:%S")

    for pptx in pptx_files:
        name = pptx.stem
        slug = slugify(name)
        folder = out_dir / slug
        folder.mkdir(parents=True, exist_ok=True)

        digest = file_hash(pptx)
        versions = load_versions(folder)

        if not versions or versions[-1]["hash"] != digest:
            n = (versions[-1]["n"] + 1) if versions else 1
            version = {
                "n": n,
                "hash": digest,
                "ts_key": ts_key,
                "ts_human": ts_human,
                "file": f"v{n:03d}-{ts_key}.html",
            }
            versions.append(version)
            is_new = True
        else:
            version = versions[-1]
            is_new = False

        titles = load_titles(samples_dir / f"{name}.titles.json")
        if not titles:
            titles = load_titles(folder / "titles.json")
        content = load_content(samples_dir / f"{name}.content.json",
                               folder / "content.json")
        deck = extract_deck(pptx)
        page = render_version_page(name, deck, titles, version, len(versions), content)

        (folder / version["file"]).write_text(page, encoding="utf-8")
        (folder / "latest.html").write_text(page, encoding="utf-8")
        (folder / "index.html").write_text(
            render_deck_index(name, versions), encoding="utf-8")
        save_versions(folder, versions)

        mark = "NEW" if is_new else "재생성"
        print(f"  + {slug}/  v{version['n']:03d} [{mark}]  ({len(deck)} slides)")
        decks.append({"name": name, "slug": slug, "versions": versions})

    (out_dir / "index.html").write_text(render_root_index(decks), encoding="utf-8")
    print(f"빌드 완료: {len(decks)}개 PPT -> {out_dir}")


def main() -> None:
    ap = argparse.ArgumentParser(description="samples/*.pptx 를 PPT별 폴더 HTML 덱으로 빌드")
    ap.add_argument("--samples", default="samples", help="입력 폴더 (기본: samples)")
    ap.add_argument("--out", default="docs/samples", help="출력 폴더 (기본: docs/samples)")
    args = ap.parse_args()
    build(Path(args.samples), Path(args.out))


if __name__ == "__main__":
    main()
