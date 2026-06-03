# Azure PowerPoint 자동 업데이트 CI/CD 설정 가이드

## 📌 프로젝트 개요

이 프로젝트는 Azure 관련 최신 정보를 기반으로 PowerPoint 자료를 **정기적으로 자동 업데이트**하는 CI/CD 파이프라인을 구축하는 것을 목표로 합니다.

- **테스트 대상 파일**: `samples/Microsoft Foundry를 활용한 Agentic AI.pptx`
- **CI/CD 플랫폼**: GitHub Actions
- **PPT 처리 엔진**: Anthropic PPT Skill (pptx 생성/편집 스킬) 기반
- **트리거 방식**: 정기 스케줄 (cron) + 수동 실행 (workflow_dispatch)

---

## 🎯 목표

1. GitHub 저장소에서 PPT 원본 파일을 버전 관리한다.
2. 정기적으로 Azure의 최신 정보(릴리스 노트, 문서, 블로그 등)를 수집한다.
3. Anthropic의 PPT 스킬을 활용해 슬라이드 내용을 최신 정보로 갱신한다.
4. 갱신된 PPT를 자동으로 커밋하거나 Pull Request로 제출한다.

---

## 🏗️ 아키텍처

```
┌────────────────────┐
│  GitHub Repository │
│  └─ samples/*.pptx │
└──────────┬─────────┘
           │ (schedule / manual)
           ▼
┌────────────────────────────────┐
│      GitHub Actions Runner     │
│  1. PPT 파일 체크아웃           │
│  2. Azure 최신 정보 수집        │
│     (Microsoft Learn / RSS 등) │
│  3. Anthropic PPT Skill 실행   │
│     - 슬라이드 분석             │
│     - 콘텐츠 업데이트           │
│  4. 변경된 PPT 저장             │
│  5. PR 생성 또는 커밋           │
└────────────────────────────────┘
```

---

## 📁 디렉터리 구조 (계획)

```
azure-contents-generator/
├── .github/
│   └── workflows/
│       └── update-ppt.yml        # GitHub Actions 워크플로
├── samples/
│   └── Microsoft Foundry를 활용한 Agentic AI.pptx
├── scripts/
│   ├── fetch_azure_updates.py    # Azure 최신 정보 수집
│   ├── update_ppt.py             # Anthropic PPT 스킬 호출
│   ├── render_ppt_previews.py    # PDF/PNG 미리보기 생성
│   └── requirements.txt
├── .skills/                      # (gitignore) anthropics/skills sparse-checkout
├── previews/                     # PPT 미리보기 (PDF + 슬라이드 PNG)
└── setup.md
```

---

## ⚙️ 필요한 사전 준비

### 1. GitHub Secrets 등록

| Secret 이름 | 용도 |
|-------------|------|
| `ANTHROPIC_API_KEY` | Anthropic Claude API 호출 (PPT skill 실행) |
| `AZURE_DOCS_SOURCE` *(선택)* | Azure 정보 소스 URL 또는 토큰 |
| `GITHUB_TOKEN` | (기본 제공) PR 생성용 |

### 2. Anthropic PPT Skill

- [`anthropics/skills`](https://github.com/anthropics/skills) 공개 레포의
  `skills/pptx/` 를 베이스로 사용합니다.
- 우리 저장소에는 포함하지 않고, **CI 단계에서 sparse-checkout** 으로 받아
  `.skills/skills/pptx` 경로에 둡니다 (라이선스: source-available).
- Claude API가 이 skill 메타데이터/스크립트를 참조해 슬라이드를 갱신합니다.

### 3. 실행 환경

- Python 3.11+
- 주요 패키지: `python-pptx`, `anthropic`, `requests`, `feedparser`

---

## 🔁 GitHub Actions 워크플로 설계

`.github/workflows/update-ppt.yml`:

```yaml
name: Update Azure PPT

on:
  schedule:
    - cron: "0 0 * * 1"   # 매주 월요일 09:00 KST
  workflow_dispatch:
    inputs:
      target_file:
        description: "업데이트할 PPT 경로"
        default: "samples/Microsoft Foundry를 활용한 Agentic AI.pptx"

jobs:
  update-ppt:
    runs-on: ubuntu-latest
    permissions:
      contents: write
      pull-requests: write

    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install -r scripts/requirements.txt

      - name: Fetch latest Azure updates
        run: python scripts/fetch_azure_updates.py --out ./.cache/updates.json

      - name: Update PPT via Anthropic skill
        env:
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
        run: |
          python scripts/update_ppt.py \
            --input  "${{ github.event.inputs.target_file || 'samples/Microsoft Foundry를 활용한 Agentic AI.pptx' }}" \
            --updates ./.cache/updates.json \
            --skill   ./skills/pptx

      - name: Create Pull Request
        uses: peter-evans/create-pull-request@v6
        with:
          commit-message: "chore(ppt): Azure 최신 정보 반영"
          branch: auto/ppt-update-${{ github.run_id }}
          title: "📊 Azure PPT 자동 업데이트"
          body: |
            Anthropic PPT skill로 자동 생성된 업데이트입니다.
            - 소스: Microsoft Learn / Azure Updates
            - 실행 시각: ${{ github.event.repository.updated_at }}
```

---

## 🧪 테스트 시나리오

1. **로컬 검증**
   - `python scripts/update_ppt.py --input samples/...pptx --dry-run` 으로 변경 사항 미리보기
2. **수동 트리거**
   - GitHub Actions에서 `workflow_dispatch`로 직접 실행하여 PR 생성 여부 확인
3. **정기 실행 검증**
   - cron 일정에 따라 자동 PR이 생성되는지 모니터링
4. **회귀 테스트**
   - 슬라이드 마스터/레이아웃/이미지가 깨지지 않는지 `python-pptx`로 차이 비교

---

## 📅 단계별 진행 계획

| 단계 | 작업 | 산출물 |
|------|------|--------|
| 1 | Anthropic pptx skill을 CI에서 sparse-checkout으로 연결 | `.github/workflows/update-ppt.yml` |
| 2 | Azure 정보 수집 스크립트 작성 | `scripts/fetch_azure_updates.py` |
| 3 | PPT 업데이트 스크립트 작성 | `scripts/update_ppt.py` |
| 4 | GitHub Actions 워크플로 작성 | `.github/workflows/update-ppt.yml` |
| 5 | 샘플 PPT로 end-to-end 테스트 | PR 자동 생성 |
| 6 | 스케줄 운영 및 품질 모니터링 | 주간 업데이트 PR |

---

## 👀 GitHub에서 PPT 미리보기

GitHub는 `.pptx`를 직접 렌더링하지 않으므로, CI에서 두 가지 형태로 변환해
저장소에 함께 커밋합니다.

| 포맷 | 도구 | GitHub에서의 동작 |
|------|------|-------------------|
| PDF  | LibreOffice (`--convert-to pdf`) | 파일 클릭 시 인라인 렌더링, PR Files changed 탭에서도 표시 |
| PNG (슬라이드별) | poppler `pdftoppm` | `previews/<파일명>/README.md` 에 모두 임베드되어 한 페이지에서 확인 가능 |

스크립트: `scripts/render_ppt_previews.py`
출력 경로: `previews/<pptx 파일명>/`

워크플로는 PPT 업데이트 후 자동으로 이 스크립트를 실행하고, 생성된
미리보기 파일을 PR에 포함시킵니다. 리뷰어는 PR에서 바로 슬라이드를
확인할 수 있습니다.

## ⚠️ 고려 사항

- **API 비용**: Anthropic API 호출 비용을 cron 주기에 맞춰 관리
- **PPT 포맷 보존**: 폰트/테마/마스터 슬라이드 유지가 핵심
- **검토 프로세스**: 자동 커밋보다는 **PR 기반 검토**를 권장 (사람 확인 후 머지)
- **출처 표기**: 인용된 Azure 문서 URL을 슬라이드 노트에 기록
- **민감 정보**: Anthropic으로 전송되는 데이터에 비공개 정보 포함되지 않도록 필터링
