"""GitHub Models (Copilot이 쓰는 동일 모델 인프라)를 호출해
Azure 최신 업데이트를 슬라이드용 한국어 요약으로 변환한다.

- 인증: GITHUB_TOKEN (Actions 기본 토큰) 만 사용
- 가이드라인: Anthropic 의 오픈소스 pptx skill (SKILL.md / editing.md) 을
  system prompt 로 주입해 "skill을 따르는 코파일럿" 처럼 동작하게 함
- 출력: 마크다운/일반 텍스트 (update_ppt.py 가 슬라이드 상단에 삽입)
- GitHub Models 접근이 불가능한 환경에서는 로컬 규칙 기반 요약으로 폴백
"""
from __future__ import annotations
import argparse, json, os, re, sys, urllib.error, urllib.request
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
        raise RuntimeError(f"GitHub Models 호출 실패: {e.code} {e.read().decode('utf-8', 'ignore')[:500]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"GitHub Models 연결 실패: {e.reason}") from e
    return payload["choices"][0]["message"]["content"].strip()


def normalize_title(title: str) -> str:
    title = re.sub(r"^\[[^\]]+\]\s*", "", title).strip()
    prefixes = (
        "Generally Available:",
        "General Availability:",
        "Public Preview:",
        "Preview:",
        "In preview:",
        "Launched:",
    )
    for prefix in prefixes:
        if title.lower().startswith(prefix.lower()):
            return title[len(prefix):].strip()
    return title


def score_item(item: dict) -> int:
    title = item.get("title", "").lower()
    score = 0
    keywords = {
        "foundry": 5,
        "agent": 4,
        "model router": 4,
        "unified model api": 4,
        "voice live": 4,
        "ai": 3,
        "cosmos db": 3,
        "api management": 2,
        "translator": 2,
        "monitor": 2,
        "ptu": 2,
    }
    for keyword, weight in keywords.items():
        if keyword in title:
            score += weight
    return score


def pick_items(items: list[dict], limit: int = 5) -> list[dict]:
    ranked = sorted(enumerate(items), key=lambda row: (-score_item(row[1]), row[0]))
    selected = [item for _, item in ranked[:limit]]
    if len(selected) < min(limit, len(items)):
        seen = {item.get("title", "") for item in selected}
        for item in items:
            title = item.get("title", "")
            if title in seen:
                continue
            selected.append(item)
            seen.add(title)
            if len(selected) >= min(limit, len(items)):
                break
    return selected


def fallback_bullet(item: dict) -> str:
    title = item.get("title", "")
    lower = title.lower()
    if "foundry agents" in lower and "vs code" in lower:
        return "- **Foundry Agents** 관측 기능이 VS Code에서 미리보기 제공 (Azure Updates)"
    if "cosmos db" in lower and "agent kit" in lower:
        return "- **Azure Cosmos DB Agent Kit** 가 GA로 개발 모범사례를 기본화 (Azure Updates)"
    if "model router" in lower and "foundry models" in lower:
        return "- **Foundry Models** Model Router 거버넌스가 미리보기 확대 (Azure Updates)"
    if "voice live" in lower and "foundry agent service" in lower:
        return "- **Foundry Agent Service** 와 Voice Live 연동이 GA로 확장 (Azure Updates)"
    if "unified model api" in lower:
        return "- **Azure API Management** 가 멀티모델 통합 API를 미리보기 공개 (Azure Updates)"
    if "ptu" in lower:
        return "- **Global PTU Reservations** 가 리전 독립형으로 운영 유연성 강화 (Azure Updates)"
    if "monitor" in lower and "alert" in lower:
        return "- **Azure Monitor** 단순 로그 경보가 GA로 운영 진입장벽 완화 (Azure Updates)"
    if "translator" in lower:
        return "- **Azure AI Translator** 통합 Text Translation API 가 GA 도달 (Azure Updates)"

    clean = normalize_title(title)
    product = clean.split(" for ", 1)[0].split(" now ", 1)[0].strip(" .:")
    if len(product) > 36:
        product = product[:33].rstrip() + "..."
    return f"- **{product or 'Azure'}** 관련 최신 기능이 공개됨 (Azure Updates)"


def build_local_summary(items: list[dict]) -> str:
    selected = pick_items(items, limit=5)
    lines = ["Agentic AI 최신 업데이트", *[fallback_bullet(item) for item in selected]]
    return "\n".join(lines)


def ensure_issue_closure(body: str, issue_number: str | None) -> str:
    text = body.rstrip()
    if not issue_number or f"Closes #{issue_number}" in text:
        return text
    return f"{text}\n\nCloses #{issue_number}"


def build_local_pr_body(items: list[dict], issue_number: str | None) -> str:
    bullets = [fallback_bullet(item) for item in pick_items(items, limit=5)]
    body = [
        "## TL;DR",
        "",
        "이번 업데이트는 **Microsoft Foundry** 중심의 Agentic AI 변경 사항과 운영·거버넌스 보강 항목을 한 장 슬라이드로 정리했습니다.",
        "새 릴리즈 폴더에는 업데이트된 PPT와 함께 검토용 미리보기 자산을 함께 담았습니다.",
        "",
        "## 🆕 주요 업데이트",
        *bullets,
        "",
        "## ✅ 리뷰 체크리스트",
        "- [ ] **Microsoft Foundry** / Azure 제품명이 원문과 일치하는지 확인",
        "- [ ] 마지막 슬라이드가 `Latest Azure Updates` 한 장만 추가되었는지 확인",
        "- [ ] `samples/*.pptx` 원본이 변경되지 않았는지 확인",
        "- [ ] 릴리즈 폴더의 PPTX·PDF·PNG·updates.json 링크가 정상 동작하는지 확인",
        "",
        "## 🔗 데이터 소스",
        "- Microsoft Release Communications RSS, Azure Updates, Microsoft Tech Community RSS",
    ]
    return ensure_issue_closure("\n".join(body), issue_number)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--updates", required=True, help="fetch_azure_updates.py 산출물 (JSON)")
    ap.add_argument("--skill", required=True, help="anthropics/skills/skills/pptx 디렉터리")
    ap.add_argument("--out", required=True, help="요약 마크다운 출력 경로")
    ap.add_argument("--pr-body", default=None, help="PR 본문 마크다운 출력 경로 (선택)")
    ap.add_argument("--issue-number", default=None, help="PR 본문 마지막에 붙일 이슈 번호 (선택)")
    ap.add_argument("--model", default=os.environ.get("GH_MODELS_MODEL", DEFAULT_MODEL))
    ap.add_argument("--limit", type=int, default=8)
    args = ap.parse_args()

    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        sys.exit("GITHUB_TOKEN 환경 변수가 필요합니다 (GitHub Actions 기본 발급).")

    skill_dir = Path(args.skill)
    updates = json.loads(Path(args.updates).read_text(encoding="utf-8"))
    all_items = updates.get("items", [])
    items = all_items[: args.limit]

    skill_ctx = load_skill_context(skill_dir)
    issue_number = str(args.issue_number) if args.issue_number else None
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
            "제품명/기능명은 굵게(`**...**`) 표기",
            "출처는 bullet 끝에 (Azure Updates) 같은 짧은 출처 태그만",
            "마크다운 헤더(#)는 사용하지 말 것",
        ],
        "items": items,
    }
    try:
        summary = call_github_models(
            token=token,
            model=args.model,
            system=system,
            user=json.dumps(user_payload, ensure_ascii=False),
        )
        print(f"[llm_summarize] GitHub Models 요약 사용: {args.model}")
    except RuntimeError as exc:
        print(f"[llm_summarize] GitHub Models 요약 실패 → 로컬 폴백 사용: {exc}")
        summary = build_local_summary(all_items)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(summary, encoding="utf-8")
    print(f"[llm_summarize] {len(items)}건 요약 → {out} ({len(summary)}자)")

    if args.pr_body:
        pr_user = {
            "task": (
                "방금 생성한 슬라이드 요약을 바탕으로, GitHub Pull Request 본문을 "
                "작성하라. 리뷰어가 머지 전에 변경 사항을 빠르게 파악할 수 있어야 한다."
            ),
            "requirements": [
                "한국어 마크다운",
                "맨 위에 2~3문장의 변경 개요(TL;DR)",
                "그 아래 '주요 업데이트' 섹션에 bullet 5개 이내",
                "제품명/기능명은 굵게(`**...**`) 표기",
                "그 아래 '리뷰 체크리스트' 섹션 (제품명 정확성, 한국어 표기, 출처 링크 등 3~5개 항목)",
                "마지막에 '데이터 소스' 1줄 (RSS / Microsoft Learn)",
                "이모지는 섹션 제목에만 1개씩",
                *(["마지막 줄에 'Closes #<issue_number>' 추가"] if issue_number else []),
            ],
            "slide_summary_markdown": summary,
            "items": items,
            "issue_number": issue_number,
        }
        try:
            pr_body = call_github_models(
                token=token, model=args.model, system=system,
                user=json.dumps(pr_user, ensure_ascii=False),
            )
            print(f"[llm_summarize] GitHub Models PR 본문 사용: {args.model}")
        except RuntimeError as exc:
            print(f"[llm_summarize] GitHub Models PR 본문 실패 → 로컬 폴백 사용: {exc}")
            pr_body = build_local_pr_body(all_items, issue_number)
        pr_body = ensure_issue_closure(pr_body, issue_number)
        pr_path = Path(args.pr_body)
        pr_path.parent.mkdir(parents=True, exist_ok=True)
        pr_path.write_text(pr_body, encoding="utf-8")
        print(f"[llm_summarize] PR 본문 → {pr_path} ({len(pr_body)}자)")


if __name__ == "__main__":
    main()
