# Azure Contents Generator

Azure 관련 PPT 자료를 **최신 정보로 자동 업데이트**하는 CI/CD 파이프라인 프로젝트입니다.
업데이트 본문은 **GitHub Models (GitHub Copilot이 사용하는 모델 인프라)** 가
Anthropic 의 오픈소스 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx)
가이드라인을 따라 작성하며, **별도 API 키 없이 `GITHUB_TOKEN` 만으로** 동작합니다.

자세한 설계는 [`setup.md`](./setup.md) 참고.

---

## 📰 최근 업데이트

각 워크플로 실행은 `releases/<YYYYMMDD-HHMMSS>-<자료명>/` 폴더에
업데이트된 **PPTX · PDF · 슬라이드별 PNG · 요약(summary.md)** 을 모아 저장합니다.
아래 두 블록은 자동 갱신됩니다 (편집하지 마세요).

### 🆕 최신 변경 내용

<!-- LATEST:START -->
_아직 업데이트가 없습니다._
<!-- LATEST:END -->

### 🗂 전체 릴리즈 이력

<!-- RELEASES:START -->
_아직 자동 생성된 업데이트가 없습니다. 워크플로를 한 번 실행해 주세요._

전체 이력: [`releases/`](./releases/)
<!-- RELEASES:END -->

---

## 🚀 GitHub에 푸시하기 (최초 1회)

### 1) GitHub에서 빈 저장소 생성

브라우저에서 https://github.com/new 로 이동:
- Repository name: `azure-contents-generator`
- README/.gitignore/license 체크 **모두 해제**

### 2) 첫 푸시

```bash
cd /Users/kichul/Documents/project/azure-contents-generator/azure-contents-generator

# (이미 git init + 첫 커밋은 완료되어 있음. 안전을 위해 확인)
git status
git log --oneline

# 본인 GitHub 사용자명/조직으로 교체
git remote add origin https://github.com/<USER>/azure-contents-generator.git
git push -u origin main
```

### 3) Actions 권한 확인

**Settings → Actions → General → Workflow permissions**:
- `Read and write permissions`
- `Allow GitHub Actions to create and approve pull requests`

LLM 사용 방식은 두 가지 중 선택:

| 엔진 | 필요 권한 | 비고 |
|------|-----------|------|
| **GitHub Models** *(기본)* | 워크플로의 `permissions: { models: read }` 만으로 자동 | **별도 Secret 불필요.** 무료 티어/엔터프라이즈 공통 |
| **GitHub Copilot CLI** *(Cloud Agent)* | `COPILOT_GITHUB_TOKEN` 시크릿 (Fine-grained PAT, *Copilot Requests: Read* 권한) | **Copilot Enterprise/Business/Pro+ 라이선스 필요.** 진짜 자율 에이전트가 skill 을 직접 읽고 실행 |

> Copilot CLI 를 쓰려면 **Settings → Secrets and variables → Actions → New repository secret**
> 에서 `COPILOT_GITHUB_TOKEN` 이름으로 PAT 를 등록한 뒤,
> 워크플로 실행 시 `engine = copilot` 을 선택하세요.

---

## 🔁 PPT 업데이트 실행 & 결과 확인

### A. 자동 실행 (스케줄)

`.github/workflows/update-ppt.yml` 에 정의된 cron 에 따라
**매주 월요일 09:00 KST** 자동 실행되어 PR을 생성합니다.

### B. 수동 실행

1. GitHub 저장소 → **Actions** 탭
2. **Update Azure PPT** 워크플로 선택
3. 우측 **Run workflow** ▼
   - `target_file`: 업데이트할 PPT 경로
   - `use_llm`: `true` (기본) 면 LLM 으로 한국어 요약, `false` 면 RSS 원문 bullet
   - `engine`:
     - `models` *(기본)* — **GitHub Models**, `GITHUB_TOKEN` 만으로 동작
     - `copilot` — **GitHub Copilot CLI Cloud Agent**, 엔터프라이즈 라이선스 + `COPILOT_GITHUB_TOKEN` 시크릿 필요
4. 완료 후 **Pull requests** 탭에 `📊 Azure PPT 자동 업데이트 …` PR 생성

### C. 업데이트된 PPT를 GitHub에서 바로 보기

각 실행마다 `releases/<YYYYMMDD-HHMMSS>-<자료명>/` 폴더가 새로 생성되며,
그 안에 다음이 모두 들어갑니다 (원본 `samples/*.pptx` 는 변경되지 않습니다).

| 보고 싶은 것 | 위치 / 방법 |
|--------------|-------------|
| 최근 실행 목록 | 위 [`📰 최근 업데이트`](#-최근-업데이트) 표 또는 [`releases/README.md`](./releases) |
| 전체 슬라이드 한 페이지에 | `releases/<TS>-<자료명>/README.md` — 모든 슬라이드 PNG 인라인 |
| PDF 인라인 보기 | `releases/<TS>-<자료명>/<자료명>.pdf` (GitHub가 PDF 자동 렌더) |
| 슬라이드별 이미지 | `releases/<TS>-<자료명>/slide-*.png` |
| 업데이트된 .pptx 다운로드 | `releases/<TS>-<자료명>/<자료명>.pptx` 의 **Download raw file** |
| LLM 요약 원문 | `releases/<TS>-<자료명>/summary.md` |
| 그때 사용한 RSS 원본 | `releases/<TS>-<자료명>/updates.json` |

PR을 머지하면 `main` 에 새 릴리즈 폴더와 갱신된 README 가 추가됩니다.

---

## 🧪 로컬에서 빠르게 시험

```bash
# 1. 의존성
pip install -r scripts/requirements.txt
# (macOS) brew install libreoffice poppler

# 2. Anthropic pptx skill 받기 (오픈소스, 인증 불필요)
git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/anthropics/skills.git .skills
git -C .skills sparse-checkout set skills/pptx

# 3. Azure 업데이트 수집
python scripts/fetch_azure_updates.py --out .cache/updates.json

# 4. (선택) GitHub Models 로 한국어 요약 생성 — Copilot 같은 동작
export GITHUB_TOKEN=ghp_...   # repo / models:read 권한이 있는 토큰
python scripts/llm_summarize.py \
  --updates .cache/updates.json \
  --skill   .skills/skills/pptx \
  --out     .cache/summary.md

# 5. 타임스탬프 릴리즈 폴더 준비
TS=$(TZ=Asia/Seoul date +'%Y%m%d-%H%M%S')
STEM="Microsoft Foundry를 활용한 Agentic AI"
REL="releases/${TS}-${STEM}"
mkdir -p "$REL"

# 6. PPT 에 업데이트 슬라이드 삽입 → 릴리즈 폴더로 출력 (원본 samples/ 보존)
python scripts/update_ppt.py \
  --input  "samples/${STEM}.pptx" \
  --updates .cache/updates.json \
  --skill   .skills/skills/pptx \
  --summary .cache/summary.md \
  --output  "${REL}/${STEM}.pptx"

# 7. 미리보기(PDF/PNG/README) 를 같은 릴리즈 폴더에 생성
python scripts/render_ppt_previews.py \
  --input "${REL}/${STEM}.pptx" \
  --out   "${REL}"

# 8. 릴리즈 인덱스 + 루트 README '최근 업데이트' 표 갱신
cp .cache/summary.md .cache/updates.json "${REL}/" 2>/dev/null || true
python scripts/update_releases_index.py
```

---

## 📁 구성

- `.github/workflows/update-ppt.yml` — 주간 cron + 수동 실행 워크플로
- `scripts/fetch_azure_updates.py` — Azure RSS/문서 수집
- `scripts/llm_summarize.py` — **GitHub Models(코파일럿 모델)** 로 한국어 요약 생성. skill SKILL.md/editing.md 를 system prompt 로 주입
- `scripts/update_ppt.py` — pptx skill 활용해 'Latest Azure Updates' 슬라이드 삽입 (LLM 요약 또는 RSS bullet)
- `scripts/render_ppt_previews.py` — PPT → PDF/PNG 미리보기 생성
- `scripts/update_releases_index.py` — `releases/` 스캔 → 인덱스 및 루트 README '최근 업데이트' 표 갱신
- `.skills/` *(gitignore)* — `anthropics/skills` sparse-checkout 캐시
- `samples/` — 원본 PPT (자동 워크플로에서 **변경되지 않음**)
- `releases/<YYYYMMDD-HHMMSS>-<자료명>/` — 실행 시각별 업데이트 결과물 (PPTX·PDF·PNG·요약 묶음)
