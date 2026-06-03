# Azure Contents Generator

Azure 관련 PPT 자료를 **최신 정보로 자동 업데이트**하는 CI/CD 파이프라인 프로젝트입니다.
업데이트 본문은 **GitHub Models (GitHub Copilot이 사용하는 모델 인프라)** 가
Anthropic 의 오픈소스 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx)
가이드라인을 따라 작성하며, **별도 API 키 없이 `GITHUB_TOKEN` 만으로** 동작합니다.

자세한 설계는 [`setup.md`](./setup.md) 참고.

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

GitHub Models 호출 권한은 워크플로의 `permissions: { models: read }` 로
자동 부여됩니다. **별도 Secret 등록은 필요 없습니다.**

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
   - `use_llm`: `true` (기본) 면 GitHub Models 로 한국어 요약 본문 생성,
     `false` 면 RSS 원문을 그대로 bullet 으로 삽입
4. 완료 후 **Pull requests** 탭에 `📊 Azure PPT 자동 업데이트 (...)` PR 생성

### C. 업데이트된 PPT를 GitHub에서 바로 보기

GitHub는 `.pptx` 를 인라인 렌더링하지 않으므로,
워크플로가 자동으로 **PDF + 슬라이드별 PNG** 미리보기를 함께 커밋합니다.

| 보고 싶은 것 | 위치 / 방법 |
|--------------|-------------|
| 전체 슬라이드 한 페이지에 | [`previews/<파일명>/README.md`](./previews) 클릭 — 모든 슬라이드 PNG가 한 화면에 |
| PDF 인라인 보기 | `previews/<파일명>/<파일명>.pdf` 클릭 (GitHub가 PDF 자동 렌더) |
| 슬라이드별 이미지 | `previews/<파일명>/slide-*.png` |
| 원본 .pptx 다운로드 | `samples/<파일명>.pptx` 의 **Download raw file** |

PR을 머지하면 `main` 의 `samples/` 와 `previews/` 가 함께 갱신됩니다.

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

# 5. PPT 에 업데이트 슬라이드 삽입 (요약이 있으면 본문에 사용)
python scripts/update_ppt.py \
  --input  "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --updates .cache/updates.json \
  --skill   .skills/skills/pptx \
  --summary .cache/summary.md

# 6. 미리보기 생성 (PDF + 슬라이드별 PNG)
python scripts/render_ppt_previews.py \
  --input "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --out   "previews/Microsoft Foundry를 활용한 Agentic AI"
```

---

## 📁 구성

- `.github/workflows/update-ppt.yml` — 주간 cron + 수동 실행 워크플로
- `scripts/fetch_azure_updates.py` — Azure RSS/문서 수집
- `scripts/llm_summarize.py` — **GitHub Models(코파일럿 모델)** 로 한국어 요약 생성. skill SKILL.md/editing.md 를 system prompt 로 주입
- `scripts/update_ppt.py` — pptx skill 활용해 'Latest Azure Updates' 슬라이드 삽입 (LLM 요약 또는 RSS bullet)
- `scripts/render_ppt_previews.py` — PPT → PDF/PNG 미리보기 생성
- `.skills/` *(gitignore)* — `anthropics/skills` sparse-checkout 캐시
- `samples/` — 테스트용 PPT
- `previews/` — PPT 미리보기 (PDF + 슬라이드별 PNG)
