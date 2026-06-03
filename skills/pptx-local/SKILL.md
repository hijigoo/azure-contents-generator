# Azure PPT Skill (로컬, 라이선스 안전)

이 폴더는 **클라우드 에이전트 / LLM 이 슬라이드를 작성할 때 따를 가이드라인**입니다.
`scripts/update_ppt.py` 와 `scripts/llm_summarize.py` 가 이 폴더를 기본 컨텍스트로 주입합니다.

> ⚖️ **라이선스 메모.** Anthropic 의 `anthropics/skills` 의 pptx skill 은 *Proprietary*
> 라이선스이므로 본 레포에 복사할 수 없습니다. 본 폴더는 그 아이디어를 참고해
> **저자 직접 작성**한 자체 가이드이며 자유롭게 수정/재배포 가능합니다.
> (선택) 워크플로에서 `anthropics/skills` 를 런타임 sparse-checkout 으로 추가 참조할
> 수 있고, 그 경우에도 본 폴더가 우선 컨텍스트입니다.

## 파일 구성
- `SKILL.md` — *(이 파일)* 작업 개요와 트리거 규칙
- `editing.md` — `python-pptx` 로 슬라이드를 안전하게 편집하는 방법
- 향후 추가 가능: `pptxgenjs.md`, `clean.py`, `thumbnail.py` 등

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
                  llm_summarize.py (skill=skills/pptx-local)
                                ↓  .cache/summary.md + .cache/pr_body.md
                  update_ppt.py     (skill=skills/pptx-local)
                                ↓  releases/<TS>/<자료명>.pptx
                  render_ppt_previews.py
                                ↓  *.pdf, slide-*.png, README.md
                  update_releases_index.py → 루트 README 자동 갱신
                                ↓
                  peter-evans/create-pull-request  (LLM-written body)
```

자세한 편집 로직은 [`editing.md`](./editing.md) 를 참고하세요.
