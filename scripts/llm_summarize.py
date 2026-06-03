"""GitHub Models (Copilot이 쓰는 동일 모델 인프라)를 호출해
Azure 최신 업데이트를 슬라이드용 한국어 요약으로 변환한다.

- 인증: GITHUB_TOKEN (Actions 기본 토큰) 만 사용
- 가이드라인: Anthropic 의 오픈소스 pptx skill (SKILL.md / editing.md) 을
  system prompt 로 주입해 "skill을 따르는 코파일럿" 처럼 동작하게 함
- 출력: 마크다운/일반 텍스트 (update_ppt.py 가 슬라이드 상단에 삽입)
"""
from __future__ import annotations
import argparse, json, os, sys, urllib.request, urllib.error
from pathlib import Path

ENDPOINT = "https://models.github.ai/inference/chat/completions"
DEFAULT_MODEL = "openai/gpt-4o-mini"


def load_skill_context(skill_dir: Path) -> str:
    parts: list[str] = []
    for fname in ("SKILL.md", "editing.md"):
        p = skill_dir / fname
        if p.exists():
            parts.append(f"# {fname}\n\n" + p.read_text(encoding="utf-8"))
    return "\n\n---\n\n".join(parts)


def call_github_models(token: str, model: str, system: str, user: str) -> str:
    body = json.dumps(
        {
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": 0.3,
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        ENDPOINT,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/vnd.github+json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        sys.exit(f"GitHub Models 호출 실패: {e.code} {e.read().decode('utf-8', 'ignore')[:500]}")
    return payload["choices"][0]["message"]["content"].strip()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--updates", required=True, help="fetch_azure_updates.py 산출물 (JSON)")
    ap.add_argument("--skill", required=True, help="anthropics/skills/skills/pptx 디렉터리")
    ap.add_argument("--out", required=True, help="요약 마크다운 출력 경로")
    ap.add_argument("--model", default=os.environ.get("GH_MODELS_MODEL", DEFAULT_MODEL))
    ap.add_argument("--limit", type=int, default=8)
    args = ap.parse_args()

    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        sys.exit("GITHUB_TOKEN 환경 변수가 필요합니다 (GitHub Actions 기본 발급).")

    skill_dir = Path(args.skill)
    updates = json.loads(Path(args.updates).read_text(encoding="utf-8"))
    items = updates.get("items", [])[: args.limit]

    skill_ctx = load_skill_context(skill_dir)
    system = (
        "너는 Azure 기술 자료를 관리하는 시니어 테크니컬 라이터다. "
        "아래는 Anthropic 의 공식 pptx skill 가이드라인이다. "
        "이 가이드라인의 슬라이드 작성 원칙(간결한 bullet, 명확한 제목, "
        "출처 표기 등)을 따르되, LLM 답변은 마크다운으로만 작성한다.\n\n"
        + skill_ctx
    )
    user_payload = {
        "task": "다음 Azure 최신 업데이트들을 한 장의 PPT 슬라이드 본문으로 정리해줘.",
        "requirements": [
            "한국어로 작성",
            "맨 위에 1줄짜리 헤드라인",
            "그 아래 4~6개 bullet (각 60자 이내), 핵심 가치 위주로",
            "출처는 bullet 끝에 (Azure Updates) 같은 짧은 출처 태그만",
            "마크다운 헤더(#)는 사용하지 말 것",
        ],
        "items": items,
    }
    summary = call_github_models(
        token=token,
        model=args.model,
        system=system,
        user=json.dumps(user_payload, ensure_ascii=False),
    )
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(summary, encoding="utf-8")
    print(f"[llm_summarize] {len(items)}건 요약 → {out} ({len(summary)}자)")


if __name__ == "__main__":
    main()
