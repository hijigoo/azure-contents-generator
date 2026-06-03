# Azure Contents Generator

Azure 관련 PowerPoint 자료를 **최신 정보로 정기 업데이트**하는 GitHub Actions 파이프라인입니다.
샘플 PPT(`samples/`)에 매주 최신 Azure 업데이트 슬라이드를 자동 갱신하고, PR로 결과를 받습니다.

> 외부 LLM API 키 없이 동작합니다. Anthropic의 오픈소스 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 스크립트와 `python-pptx` 만 사용합니다.

---

## ✨ 동작

1. 매주 월요일 09:00 KST (cron) 또는 수동 트리거로 워크플로 실행
2. Azure Updates / Microsoft 블로그 RSS 에서 최신 항목 수집
3. `anthropics/skills` 의 `skills/pptx` 를 sparse-checkout
4. 대상 PPT 끝에 **"Latest Azure Updates"** 슬라이드를 멱등 삽입
   (본문 = 최신 항목 제목·날짜, 슬라이드 노트 = 링크·요약)
5. PPT 를 PDF + 슬라이드별 PNG 로 변환해 `previews/` 에 커밋
6. 변경 사항을 PR로 자동 제출

---

## 🚀 사용 방법

### 자동 실행

`.github/workflows/update-ppt.yml` 의 cron 에 따라 매주 월요일 자동 실행됩니다.

### 수동 실행

1. GitHub 저장소 → **Actions** 탭
2. **Update Azure PPT** 워크플로 선택
3. 우측 **Run workflow** ▼ → 필요 시 `target_file` 변경 → **Run workflow**
4. 완료되면 **Pull requests** 탭에 `📊 Azure PPT 자동 업데이트 (...)` PR 생성

### 업데이트된 PPT를 GitHub에서 바로 보기

`.pptx` 는 GitHub가 인라인 렌더링하지 않아, 워크플로가 자동으로 **PDF + 슬라이드별 PNG** 미리보기를 함께 커밋합니다.

| 보고 싶은 것 | 위치 |
|--------------|------|
| 모든 슬라이드 한 화면에 | [`previews/<파일명>/README.md`](./previews) |
| PDF 인라인 보기 | `previews/<파일명>/<파일명>.pdf` (클릭 시 GitHub 자동 렌더) |
| 슬라이드별 이미지 | `previews/<파일명>/slide-*.png` |
| 원본 .pptx 다운로드 | `samples/<파일명>.pptx` → **Download raw file** |

PR 머지 후 `main` 의 `samples/` 와 `previews/` 가 함께 갱신됩니다.

---

## 🛠️ 로컬에서 동일 작업 실행

```bash
# 의존성
pip install -r scripts/requirements.txt
# macOS: brew install libreoffice poppler

# pptx skill 받기 (오픈소스, 인증 불필요)
git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/anthropics/skills.git .skills
git -C .skills sparse-checkout set skills/pptx

# 1) Azure 업데이트 수집
python scripts/fetch_azure_updates.py --out .cache/updates.json

# 2) PPT 에 'Latest Azure Updates' 슬라이드 삽입
python scripts/update_ppt.py \
  --input  "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --updates .cache/updates.json \
  --skill   .skills/skills/pptx

# 3) 미리보기(PDF + 슬라이드별 PNG) 생성
python scripts/render_ppt_previews.py \
  --input "samples/Microsoft Foundry를 활용한 Agentic AI.pptx" \
  --out   "previews/Microsoft Foundry를 활용한 Agentic AI"
```

---

## 📁 컨텐츠 업데이트 관련 파일

| 경로 | 역할 |
|------|------|
| `.github/workflows/update-ppt.yml` | 주간 cron + 수동 실행 워크플로 |
| `scripts/fetch_azure_updates.py`   | Azure RSS 최신 항목 수집 |
| `scripts/update_ppt.py`            | PPT에 'Latest Azure Updates' 슬라이드 삽입 (멱등) |
| `scripts/render_ppt_previews.py`   | PPT → PDF + 슬라이드별 PNG 변환 |
| `samples/`                         | 업데이트 대상 PPT |
| `previews/`                        | 자동 생성된 미리보기 (PR/머지 후 확인용) |
