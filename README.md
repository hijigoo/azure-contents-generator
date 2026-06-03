# Azure Contents Generator

Azure 관련 PPT 자료를 **최신 정보로 자동 업데이트**하는 CI/CD 파이프라인 프로젝트입니다.
자세한 설계와 운영 가이드는 [`setup.md`](./setup.md)를 참고하세요.

---

## 🚀 GitHub에 푸시하기 (최초 1회)

이 디렉터리는 아직 git 저장소가 아니거나 원격이 연결되어 있지 않습니다.
아래 순서대로 실행해 GitHub에 올립니다.

### 1) GitHub에서 빈 저장소 생성

브라우저에서 https://github.com/new 로 이동해 다음과 같이 만듭니다.

- Repository name: `azure-contents-generator`
- README, .gitignore, license 체크 **모두 해제** (이미 로컬에 있음)
- Private/Public는 자유

### 2) 로컬 초기화 & 첫 푸시

```bash
cd /Users/kichul/Documents/project/azure-contents-generator/azure-contents-generator

git init -b main
git add .
git commit -m "chore: initial scaffold for Azure PPT auto-update pipeline"

# 본인 GitHub 사용자명/조직으로 교체
git remote add origin https://github.com/<USER>/azure-contents-generator.git
git push -u origin main
```

> 이미 위의 `git init` / `git commit` 까지는 CLI가 미리 수행해 두었습니다.
> 그래도 안전을 위해 `git status` 로 한 번 확인 후 진행하세요.

### 3) GitHub Secret 등록

푸시한 저장소의 **Settings → Secrets and variables → Actions → New repository secret**:

| Name | Value |
|------|-------|
| `ANTHROPIC_API_KEY` | Anthropic 콘솔에서 발급한 API 키 |

### 4) Actions 권한 확인

**Settings → Actions → General → Workflow permissions** 에서
`Read and write permissions` + `Allow GitHub Actions to create and approve pull requests`
를 활성화합니다. (PR 자동 생성 권한 필요)

---

## 🔁 PPT 업데이트 실행 & 결과 확인 방법

### A. 자동 실행 (스케줄)

`.github/workflows/update-ppt.yml` 에 정의된 cron 에 따라 **매주 월요일 09:00 KST**
자동 실행됩니다. 실행 결과는 PR로 도착합니다.

### B. 수동 실행 (원할 때 바로)

1. GitHub 저장소 페이지 → 상단 **Actions** 탭
2. 좌측 목록에서 **Update Azure PPT** 선택
3. 우측 **Run workflow** ▼ 클릭
4. 필요 시 `target_file`, `dry_run` 입력 후 **Run workflow**
5. 잠시 후 워크플로 실행이 나타나고, 완료되면 **Pull requests** 탭에
   `📊 Azure PPT 자동 업데이트 (...)` PR이 생성됨

### C. 업데이트된 PPT 파일 확인하기

GitHub는 `.pptx` 를 인라인 렌더링하지 않아, 워크플로가 자동으로
**PDF + 슬라이드별 PNG** 미리보기를 함께 커밋합니다.

| 보고 싶은 것 | 위치 / 방법 |
|--------------|-------------|
| 전체 슬라이드 한 페이지에 | PR에서 [`previews/<파일명>/README.md`](./previews) 클릭 → 모든 슬라이드 PNG가 한 화면에 렌더링 |
| PDF 인라인 보기 | PR **Files changed** 탭에서 `previews/<파일명>/<파일명>.pdf` 클릭 (GitHub가 PDF 자동 렌더) |
| 슬라이드별 이미지 | `previews/<파일명>/slide-*.png` 각각 클릭 |
| 원본 .pptx 다운로드 | `samples/<파일명>.pptx` 의 **Download raw file** 버튼 |
| PR diff (텍스트 변경) | PR **Files changed** 탭에 `.pptx` 는 바이너리로만 표시됨 → 위 PNG/PDF로 검토 |

### D. 머지 후 main에서 보기

PR을 머지하면 `main` 의 `samples/` (원본) 와 `previews/` (미리보기)가
함께 갱신됩니다. 저장소 메인 페이지에서 `previews/<파일명>/` 로 들어가면
바로 슬라이드를 확인할 수 있습니다.

---

## 🧪 로컬에서 빠르게 시험

```bash
# 1. 의존성
pip install -r scripts/requirements.txt
# (macOS) brew install libreoffice poppler   # 미리보기 변환용

# 2. Anthropic 공식 pptx skill sparse-checkout
git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/anthropics/skills.git .skills
git -C .skills sparse-checkout set skills/pptx
export SKILL_PATH="$(pwd)/.skills/skills/pptx"

# 3. Azure 업데이트 수집
python scripts/fetch_azure_updates.py --out .cache/updates.json

# 4. PPT 업데이트 (dry-run 으로 미리 확인)
export ANTHROPIC_API_KEY=sk-ant-...
python scripts/update_ppt.py \
  --input  "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --updates .cache/updates.json \
  --skill   "$SKILL_PATH" \
  --dry-run

# 5. 미리보기(PDF/PNG) 생성
python scripts/render_ppt_previews.py \
  --input "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --out   "previews/Microsoft Foundry를 활용한 Agentic AI"
```

---

## 📁 구성

- `.github/workflows/update-ppt.yml` — 주간 cron + 수동 실행 워크플로
- `scripts/fetch_azure_updates.py` — Azure RSS/문서 수집
- `scripts/update_ppt.py` — Anthropic PPT skill 호출 및 슬라이드 반영
- `scripts/render_ppt_previews.py` — PPT → PDF/PNG 미리보기 생성
- `.skills/` *(gitignore)* — `anthropics/skills` sparse-checkout 캐시
- `samples/` — 테스트용 PPT
- `previews/` — PPT 미리보기 (PDF + 슬라이드별 PNG)

## Anthropic PPT Skill

본 프로젝트는 [`anthropics/skills`](https://github.com/anthropics/skills) 레포의
`skills/pptx/` 를 그대로 사용합니다. 우리 저장소에는 포함하지 않고
CI(또는 로컬 setup)에서 sparse-checkout 으로 가져옵니다.
