"""GitHub Pages 정적 사이트 빌드.

releases/ 폴더를 스캔해 site/ 에 다음을 생성한다:

- site/index.html        : 전체 릴리즈 카드 목록 (최신순, slide-01.png 썸네일)
- site/releases/<id>/index.html : 릴리즈별 슬라이드 갤러리 + 다운로드 링크
- site/releases/<id>/slide-*.png : 미리보기 PNG 복사
- site/assets/style.css  : 공통 스타일 (라이트/다크 자동)

PPTX / PDF 같은 큰 바이너리는 사이트에 복사하지 않고 GitHub raw URL 로 링크해서
Pages 아티팩트를 가볍게 유지한다. 환경변수:

    GITHUB_REPOSITORY  : owner/repo (Actions 에서 자동 주입; 로컬은 default)
    GITHUB_REF_NAME    : 기본 브랜치명 (Actions 에서 자동 주입; 로컬은 main)
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
from html import escape
from pathlib import Path
from urllib.parse import quote

RELEASE_NAME_RE = re.compile(r"^(\d{8})-(\d{6})-(.+)$")
SLIDE_RE = re.compile(r"^slide-(\d+)\.png$")


def repo_raw_base() -> str:
    repo = os.environ.get("GITHUB_REPOSITORY", "hijigoo/azure-contents-generator")
    ref = os.environ.get("GITHUB_REF_NAME", "main")
    return f"https://raw.githubusercontent.com/{repo}/{ref}"


def list_releases(releases_dir: Path) -> list[Path]:
    if not releases_dir.exists():
        return []
    items = [p for p in releases_dir.iterdir() if p.is_dir() and RELEASE_NAME_RE.match(p.name)]
    return sorted(items, key=lambda p: p.name, reverse=True)


def parse_release(folder: Path) -> dict:
    m = RELEASE_NAME_RE.match(folder.name)
    date_s, time_s, stem = m.group(1), m.group(2), m.group(3)
    ts_human = f"{date_s[:4]}-{date_s[4:6]}-{date_s[6:8]} {time_s[:2]}:{time_s[2:4]}:{time_s[4:6]} KST"
    slides = sorted(
        (p for p in folder.iterdir() if SLIDE_RE.match(p.name)),
        key=lambda p: int(SLIDE_RE.match(p.name).group(1)),
    )
    pptx = next((p for p in folder.iterdir() if p.suffix.lower() == ".pptx"), None)
    pdf = next((p for p in folder.iterdir() if p.suffix.lower() == ".pdf"), None)
    summary = folder / "summary.md"
    return {
        "folder": folder,
        "id": folder.name,
        "stem": stem,
        "ts_human": ts_human,
        "ts_key": f"{date_s}-{time_s}",
        "slides": slides,
        "pptx": pptx,
        "pdf": pdf,
        "summary": summary if summary.exists() else None,
    }


CSS = """:root {
  color-scheme: light dark;
  --bg: #ffffff;
  --fg: #0f172a;
  --muted: #64748b;
  --border: #e2e8f0;
  --card-bg: #ffffff;
  --accent: #2563eb;
  --accent-fg: #ffffff;
  --code-bg: #f1f5f9;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #0b1020;
    --fg: #e2e8f0;
    --muted: #94a3b8;
    --border: #1e293b;
    --card-bg: #0f172a;
    --accent: #60a5fa;
    --accent-fg: #0b1020;
    --code-bg: #1e293b;
  }
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: var(--bg); color: var(--fg);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Apple SD Gothic Neo",
    "Noto Sans KR", "Malgun Gothic", sans-serif;
  line-height: 1.55; }
.wrap { max-width: 1120px; margin: 0 auto; padding: 32px 20px 80px; }
header.site { display: flex; align-items: baseline; justify-content: space-between;
  gap: 16px; flex-wrap: wrap; margin-bottom: 24px; }
header.site h1 { margin: 0; font-size: 1.6rem; }
header.site .sub { color: var(--muted); font-size: 0.95rem; }
header.site nav { display: flex; gap: 16px; align-items: center; }
header.site nav a { color: var(--accent); text-decoration: none; font-size: 0.95rem; }
header.site nav a:hover { text-decoration: underline; }

/* sample landing */
.hero { padding: 56px 0 40px; text-align: center; }
.hero .eyebrow { display: inline-block; padding: 4px 12px; border-radius: 999px;
  background: var(--code-bg); color: var(--muted); font-size: 0.8rem; margin-bottom: 16px;
  letter-spacing: 0.04em; text-transform: uppercase; }
.hero h1 { font-size: clamp(2rem, 5vw, 3.2rem); line-height: 1.15; margin: 0 0 16px;
  letter-spacing: -0.02em; }
.hero h1 .grad { background: linear-gradient(135deg, var(--accent), #a855f7);
  -webkit-background-clip: text; background-clip: text; color: transparent; }
.hero p.lead { font-size: 1.15rem; color: var(--muted); max-width: 640px; margin: 0 auto 28px; }
.hero .ctas { display: inline-flex; gap: 12px; flex-wrap: wrap; justify-content: center; }
.features { display: grid; gap: 20px; padding: 40px 0;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); }
.feature { background: var(--card-bg); border: 1px solid var(--border); border-radius: 12px;
  padding: 20px; }
.feature .icon { font-size: 1.8rem; margin-bottom: 8px; }
.feature h3 { margin: 0 0 6px; font-size: 1.05rem; }
.feature p { margin: 0; color: var(--muted); font-size: 0.92rem; }
.cta-band { background: var(--card-bg); border: 1px solid var(--border); border-radius: 16px;
  padding: 36px 28px; text-align: center; margin: 24px 0; }
.cta-band h2 { margin: 0 0 8px; font-size: 1.5rem; }
.cta-band p { margin: 0 0 20px; color: var(--muted); }
.cards { display: grid; gap: 20px;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); }
.card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 12px;
  overflow: hidden; display: flex; flex-direction: column; transition: transform .12s ease;
  text-decoration: none; color: inherit; }
.card:hover { transform: translateY(-2px); border-color: var(--accent); }
.card .thumb { aspect-ratio: 16/9; background: #000; overflow: hidden;
  display: flex; align-items: center; justify-content: center; }
.card .thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
.card .body { padding: 14px 16px 16px; display: flex; flex-direction: column; gap: 6px; }
.card .title { font-weight: 600; font-size: 1rem; word-break: keep-all; }
.card .meta { color: var(--muted); font-size: 0.85rem; }
.card .badges { margin-top: 8px; display: flex; gap: 6px; flex-wrap: wrap; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 999px;
  background: var(--code-bg); color: var(--muted); font-size: 0.75rem; }

article.release header { margin-bottom: 20px; }
article.release h1 { margin: 0 0 6px; font-size: 1.5rem; word-break: keep-all; }
article.release .meta { color: var(--muted); font-size: 0.9rem; margin-bottom: 12px; }
.actions { display: flex; gap: 10px; flex-wrap: wrap; margin: 8px 0 24px; }
.btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px;
  border-radius: 8px; background: var(--accent); color: var(--accent-fg);
  text-decoration: none; font-weight: 500; font-size: 0.9rem; border: 1px solid var(--accent); }
.btn.secondary { background: transparent; color: var(--accent); }
.btn:hover { filter: brightness(1.05); }
.slides { display: grid; gap: 16px;
  grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); }
.slide { background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px;
  overflow: hidden; cursor: zoom-in; }
.slide img { width: 100%; height: auto; display: block; }
.slide .num { padding: 6px 10px; color: var(--muted); font-size: 0.78rem; font-variant-numeric: tabular-nums; }
.back { display: inline-block; margin-bottom: 16px; color: var(--accent); text-decoration: none; }
.back:hover { text-decoration: underline; }

/* lightbox */
dialog.lightbox { padding: 0; border: 0; background: transparent; width: 96vw; max-width: 1400px; }
dialog.lightbox::backdrop { background: rgba(0,0,0,.85); }
dialog.lightbox img { width: 100%; height: auto; display: block; border-radius: 8px; }
dialog.lightbox .close { position: absolute; top: 12px; right: 16px; color: #fff;
  background: rgba(0,0,0,.4); border: 0; padding: 6px 12px; border-radius: 6px; cursor: pointer; }
footer.site { margin-top: 60px; color: var(--muted); font-size: 0.85rem; text-align: center; }
footer.site a { color: var(--accent); }
"""


LIGHTBOX_JS = """
(function() {
  const dlg = document.querySelector('dialog.lightbox');
  if (!dlg) return;
  const img = dlg.querySelector('img');
  document.querySelectorAll('.slide img').forEach(el => {
    el.parentElement.addEventListener('click', () => {
      img.src = el.src;
      if (typeof dlg.showModal === 'function') dlg.showModal();
    });
  });
  dlg.addEventListener('click', e => {
    if (e.target === dlg || e.target.classList.contains('close')) dlg.close();
  });
})();
"""


def page_shell(title: str, body: str, *, css_href: str = "assets/style.css") -> str:
    return f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{escape(title)}</title>
<link rel="stylesheet" href="{escape(css_href)}" />
</head>
<body>
<div class="wrap">
{body}
</div>
</body>
</html>
"""


def render_index(releases: list[dict], repo: str) -> str:
    cards: list[str] = []
    for r in releases:
        thumb_rel = f"releases/{quote(r['id'])}/slide-01.png" if r["slides"] else ""
        thumb_html = (
            f'<div class="thumb"><img loading="lazy" src="{escape(thumb_rel)}" alt="" /></div>'
            if thumb_rel
            else '<div class="thumb"></div>'
        )
        n = len(r["slides"])
        cards.append(
            f'''<a class="card" href="releases/{escape(quote(r["id"]))}/">
              {thumb_html}
              <div class="body">
                <div class="title">{escape(r["stem"])}</div>
                <div class="meta">{escape(r["ts_human"])}</div>
                <div class="badges">
                  <span class="badge">슬라이드 {n}장</span>
                  {"<span class='badge'>PPTX</span>" if r["pptx"] else ""}
                  {"<span class='badge'>PDF</span>" if r["pdf"] else ""}
                </div>
              </div>
            </a>'''
        )
    cards_html = "\n".join(cards) if cards else '<p style="color:var(--muted)">아직 릴리즈가 없습니다.</p>'
    body = f"""<header class="site">
  <div>
    <h1>📊 Azure Contents Generator</h1>
    <div class="sub">최신 Azure / Microsoft 업데이트가 반영된 PPT 릴리즈 카탈로그</div>
  </div>
  <nav>
    <a href="sample/">소개</a>
    <a href="https://github.com/{escape(repo)}" target="_blank" rel="noopener">GitHub ↗</a>
  </nav>
</header>
<section class="cards">
{cards_html}
</section>
<footer class="site">
  자동 생성 · <a href="https://github.com/{escape(repo)}">{escape(repo)}</a>
</footer>
"""
    return page_shell("Azure Contents Generator — Releases", body, css_href="assets/style.css")


def render_release(r: dict, repo: str, raw_base: str) -> str:
    slides_html: list[str] = []
    for p in r["slides"]:
        num = int(SLIDE_RE.match(p.name).group(1))
        slides_html.append(
            f'<figure class="slide"><img loading="lazy" src="{escape(p.name)}" alt="slide {num}" />'
            f'<figcaption class="num">슬라이드 {num:02d}</figcaption></figure>'
        )
    slides_block = "\n".join(slides_html) or '<p style="color:var(--muted)">슬라이드 미리보기가 없습니다.</p>'

    raw_folder = f"{raw_base}/releases/{quote(r['id'])}"
    actions: list[str] = []
    if r["pptx"]:
        actions.append(
            f'<a class="btn" href="{escape(raw_folder)}/{escape(quote(r["pptx"].name))}" download>📊 PPTX 다운로드</a>'
        )
    if r["pdf"]:
        actions.append(
            f'<a class="btn secondary" href="{escape(raw_folder)}/{escape(quote(r["pdf"].name))}" download>📄 PDF 다운로드</a>'
        )
    actions.append(
        f'<a class="btn secondary" href="https://github.com/{escape(repo)}/tree/main/releases/{escape(quote(r["id"]))}" target="_blank" rel="noopener">📁 GitHub 폴더 ↗</a>'
    )
    actions_html = "\n".join(actions)

    body = f"""<a class="back" href="../../">← 모든 릴리즈</a>
<article class="release">
  <header>
    <h1>{escape(r["stem"])}</h1>
    <div class="meta">{escape(r["ts_human"])} · 슬라이드 {len(r["slides"])}장</div>
    <div class="actions">{actions_html}</div>
  </header>
  <section class="slides">
{slides_block}
  </section>
</article>
<dialog class="lightbox"><button class="close">닫기 ✕</button><img alt="" /></dialog>
<script>{LIGHTBOX_JS}</script>
<footer class="site">
  자동 생성 · <a href="https://github.com/{escape(repo)}">{escape(repo)}</a>
</footer>
"""
    return page_shell(f"{r['stem']} — {r['ts_human']}", body, css_href="../../assets/style.css")


def render_sample(repo: str) -> str:
    body = f"""<header class="site">
  <div>
    <h1><a href="../" style="color:inherit;text-decoration:none">📊 Azure Contents Generator</a></h1>
    <div class="sub">샘플 랜딩 페이지</div>
  </div>
  <nav>
    <a href="../">릴리즈</a>
    <a href="https://github.com/{escape(repo)}" target="_blank" rel="noopener">GitHub ↗</a>
  </nav>
</header>

<section class="hero">
  <span class="eyebrow">Sample Page</span>
  <h1>Azure 최신 소식을 <span class="grad">PPT로 자동</span> 정리합니다</h1>
  <p class="lead">
    매주 Microsoft 공식 RSS를 수집하고, GitHub Models 또는 Copilot Coding Agent가
    한국어 슬라이드를 생성합니다. PR 한 번이면 배포까지 완료됩니다.
  </p>
  <div class="ctas">
    <a class="btn" href="../">📚 릴리즈 보기</a>
    <a class="btn secondary" href="https://github.com/{escape(repo)}" target="_blank" rel="noopener">⭐ GitHub</a>
  </div>
</section>

<section class="features">
  <div class="feature">
    <div class="icon">🤖</div>
    <h3>LLM 기반 자동 요약</h3>
    <p>GitHub Models 무료 티어로 한국어 슬라이드 본문을 매주 자동 생성합니다.</p>
  </div>
  <div class="feature">
    <div class="icon">🎨</div>
    <h3>원본 디자인 유지</h3>
    <p>python-pptx로 슬라이드 마스터와 테마를 그대로 보존하며 콘텐츠만 갱신합니다.</p>
  </div>
  <div class="feature">
    <div class="icon">📦</div>
    <h3>릴리즈 카탈로그</h3>
    <p>모든 산출물(PPTX · PDF · PNG)이 releases/ 폴더에 시간순으로 누적됩니다.</p>
  </div>
  <div class="feature">
    <div class="icon">🌗</div>
    <h3>다크 모드 자동</h3>
    <p>OS 테마를 따라 라이트/다크가 자동 전환됩니다. 외부 의존성 없는 순수 Python 빌드.</p>
  </div>
</section>

<section class="cta-band">
  <h2>이 프로젝트가 마음에 드시나요?</h2>
  <p>레포를 fork하거나 issue로 의견을 남겨주세요.</p>
  <a class="btn" href="https://github.com/{escape(repo)}" target="_blank" rel="noopener">GitHub에서 보기 ↗</a>
</section>

<footer class="site">
  자동 생성 · <a href="https://github.com/{escape(repo)}">{escape(repo)}</a>
</footer>
"""
    return page_shell("샘플 페이지 — Azure Contents Generator", body, css_href="../assets/style.css")


def build(releases_dir: Path, out_dir: Path) -> None:
    repo = os.environ.get("GITHUB_REPOSITORY", "hijigoo/azure-contents-generator")
    raw_base = repo_raw_base()

    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    (out_dir / "assets").mkdir(parents=True, exist_ok=True)
    (out_dir / "assets" / "style.css").write_text(CSS, encoding="utf-8")
    (out_dir / ".nojekyll").write_text("", encoding="utf-8")

    releases = [parse_release(p) for p in list_releases(releases_dir)]

    (out_dir / "index.html").write_text(render_index(releases, repo), encoding="utf-8")

    sample_dir = out_dir / "sample"
    sample_dir.mkdir(parents=True, exist_ok=True)
    (sample_dir / "index.html").write_text(render_sample(repo), encoding="utf-8")

    for r in releases:
        rel_out = out_dir / "releases" / r["id"]
        rel_out.mkdir(parents=True, exist_ok=True)
        for slide in r["slides"]:
            shutil.copy2(slide, rel_out / slide.name)
        (rel_out / "index.html").write_text(render_release(r, repo, raw_base), encoding="utf-8")

    print(f"built {len(releases)} release page(s) into {out_dir}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build GitHub Pages site from releases/")
    parser.add_argument("--releases", default="releases", help="releases directory")
    parser.add_argument("--out", default="site", help="output directory")
    args = parser.parse_args()
    build(Path(args.releases), Path(args.out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
