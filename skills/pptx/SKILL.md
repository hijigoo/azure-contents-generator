# Azure PPT Skill

이 폴더는 **클라우드 에이전트 / LLM 이 슬라이드를 작성할 때 따를 가이드라인**입니다.
`scripts/update_ppt.py` 와 `scripts/llm_summarize.py` 가 이 폴더를 기본 컨텍스트로 주입하고,
`.github/workflows/delegate-to-copilot-agent.yml` 의 Copilot Coding Agent 도 이 폴더를 먼저 읽습니다.

> ⚖️ **라이선스 메모 — 중요.**
> Anthropic 의 공식 PPTX skill ([`anthropics/skills/skills/pptx`](https://github.com/anthropics/skills/tree/main/skills/pptx))
> 은 *Proprietary* 라이선스이며 *"reproduce, copy, create derivative works, distribute"* 가
> 명시적으로 금지되어 있습니다. 따라서 본 레포에는 **그 콘텐츠를 복사/이식하지 않습니다.**
> 이 폴더의 모든 문서는 **저자가 직접 작성**한 자체 가이드이며 본 레포의 라이선스를 따릅니다.
> 공식 skill 은 *참고용 외부 링크*로만 안내하며, Claude 서비스 안에서 사용하실 수 있습니다.

## 파일 구성
- `SKILL.md` — *(이 파일)* 작업 개요와 트리거 규칙
- `editing.md` — `python-pptx` 로 슬라이드를 안전하게 편집하는 방법

## 트리거 규칙
다음 중 하나라도 해당하면 이 skill 을 사용한다:
- 사용자가 `.pptx`, "슬라이드", "덱", "프레젠테이션" 을 언급
- `samples/*.pptx` 의 내용을 갱신해야 할 때
- 한 번의 PR 로 PPT + 미리보기를 함께 검토할 수 있도록 만들어야 할 때

## 작성 원칙 (요약)
| 항목 | 규칙 |
|------|------|
| 언어 | 한국어. 제품·기능 이름은 영문 표기 유지 |
| 헤드라인 | 1줄, 35자 이내, 명사형 종결 |
| 본문 bullet | 4~6개, 각 60자 이내, 핵심 가치 위주 |
| 출처 | bullet 끝에 `(Azure Updates)` 처럼 짧은 태그만 |
| 마크다운 헤더 | `#` 금지 — bullet 과 인라인 강조(`**굵게**`) 만 |
| 빈 슬라이드 | 만들지 말 것. 최소 1개의 시각 요소 또는 본문 필요 |

## 워크플로 안의 역할
```text
fetch_azure_updates.py  →  .cache/updates.json
                                ↓
                  llm_summarize.py (skill=skills/pptx)
                                ↓  .cache/summary.md + .cache/pr_body.md
                  update_ppt.py     (skill=skills/pptx)
                                ↓  releases/<TS>/<자료명>.pptx
                  render_ppt_previews.py
                                ↓  *.pdf, slide-*.png, README.md
                  update_releases_index.py → 루트 README 자동 갱신
                                ↓
                  peter-evans/create-pull-request  (LLM-written body)
```

자세한 편집 로직은 [`editing.md`](./editing.md) 를 참고하세요.
