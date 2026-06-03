# Azure Contents Generator

Azure 관련 PPT 자료를 **최신 정보로 자동 업데이트**하는 CI/CD 파이프라인 프로젝트입니다.
**외부 LLM API 인증이 전혀 필요 없습니다** — Anthropic의 오픈소스 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 의 스크립트와 `python-pptx` 만으로 동작합니다.

자세한 설계는 [`setup.md`](./setup.md) 참고.

---

## ✨ 동작 요약

1. 매주 월요일 (또는 수동 실행) GitHub Actions 가 트리거
2. Azure Updates / Microsoft 블로그 **RSS 피드**에서 최신 항목 수집
3. `anthropics/skills` 레포의 `skills/pptx` 를 sparse-checkout
4. 대상 PPT의 마지막 슬라이드에 **"Latest Azure Updates"** 슬라이드를
   기계적으로 삽입 (멱등 — 매번 같은 자리 갱신)
5. 슬라이드를 **PDF + 슬라이드별 PNG** 로 변환해 `previews/` 에 저장
6. 변경 사항을 PR로 자동 제출

> 인증이 필요한 곳: **없음.** (`GITHUB_TOKEN` 은 GitHub Actions가 자동 발급)
>
> 만약 슬라이드 본문을 LLM으로 재작성하고 싶다면, 별도 Step에서
> `anthropics/claude-code-action` 이나 `actions/ai-inference` 를 추가하면 됩니다.
> 본 기본 구성은 그런 키 없이 RSS 데이터를 그대로 보여줍니다.

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

### 3) Actions 권한만 확인

**Settings → Actions → General → Workflow permissions**:
- `Read and write permissions`
- `Allow GitHub Actions to create and approve pull requests`

→ 위 두 가지만 체크하면 끝. **Secret 등록 필요 없음.**

---

## 🔁 PPT 업데이트 실행 & 결과 확인

### A. 자동 실행 (스케줄)

`.github/workflows/update-ppt.yml` 에 정의된 cron 에 따라
**매주 월요일 09:00 KST** 자동 실행되어 PR을 생성합니다.

### B. 수동 실행

1. GitHub 저장소 → **Actions** 탭
2. **Update Azure PPT** 워크플로 선택
3. 우측 **Run workflow** ▼ → 필요 시 `target_file` 변경 → **Run workflow**
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

# 4. PPT 에 업데이트 슬라이드 삽입
python scripts/update_ppt.py \
  --input  "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --updates .cache/updates.json \
  --skill   .skills/skills/pptx

# 5. 미리보기 생성 (PDF + 슬라이드별 PNG)
python scripts/render_ppt_previews.py \
  --input "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --out   "previews/Microsoft Foundry를 활용한 Agentic AI"
```

---

## 📁 구성

- `.github/workflows/update-ppt.yml` — 주간 cron + 수동 실행 워크플로
- `scripts/fetch_azure_updates.py` — Azure RSS/문서 수집
- `scripts/update_ppt.py` — pptx skill 활용해 'Latest Azure Updates' 슬라이드 삽입
- `scripts/render_ppt_previews.py` — PPT → PDF/PNG 미리보기 생성
- `.skills/` *(gitignore)* — `anthropics/skills` sparse-checkout 캐시
- `samples/` — 테스트용 PPT
- `previews/` — PPT 미리보기 (PDF + 슬라이드별 PNG)
