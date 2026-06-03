# Editing Guide — Azure PPT 슬라이드 갱신

본 문서는 `python-pptx` 로 기존 PPT 에 **"Latest Azure Updates" 한 장**을 추가/갱신할 때
LLM·클라우드 에이전트가 따라야 하는 편집 규칙입니다. `scripts/update_ppt.py` 의 동작과 1:1 대응합니다.

## 1. 멱등 업데이트
같은 제목(`Latest Azure Updates`) 슬라이드가 이미 있으면 먼저 제거하고 새로 1장 추가.
이렇게 해야 매주 실행해도 슬라이드가 누적되지 않습니다.

```python
xml_slides = prs.slides._sldIdLst
for idx, slide in enumerate(list(prs.slides)):
    title = slide.shapes.title
    if title and title.has_text_frame and title.text_frame.text.strip() == "Latest Azure Updates":
        xml_slides.remove(list(xml_slides)[idx])
```

## 2. 레이아웃 선택
- 가능하면 `prs.slide_layouts[1]` (Title and Content) 사용.
- 회사 템플릿 PPT 의 경우 layout 인덱스가 다를 수 있으므로,
  `prs.slide_layouts` 를 순회해 `Title and Content` / `제목 및 내용` 이름을 우선 매칭.

## 3. 본문 구성 우선순위
1. **`summary_text` 가 제공되면** (LLM 결과) 본문에 그대로 표시 — 한국어 자연어 bullet.
2. 그렇지 않으면 `updates.json` 의 `items[:limit]` 을 bullet 으로 변환.

```python
body_tf.text = first_line
for para in remaining_lines:
    p = body_tf.add_paragraph()
    p.text = para
    p.level = 0
```

## 4. 노트(Speaker Notes) 에는 원문 메타 보존
LLM 요약을 본문에 쓰더라도, 노트에는 원문 RSS 메타(제목·날짜·URL)를 보존해야
머지 후에도 추적이 가능합니다.

```python
notes = slide.notes_slide.notes_text_frame
for item in items:
    notes.add_paragraph().text = f"- {item['title']} ({item['published']}) {item['link']}"
```

## 5. 폰트·색상
- 헤드라인: 28~32pt bold (템플릿 폰트 유지)
- 본문: 16~18pt
- 새 색상을 임의로 도입하지 말 것. 템플릿 마스터의 색을 그대로 사용.

## 6. QA 체크리스트 (PR 생성 직전)
- [ ] 추가된 슬라이드 1장만 늘었는가? (전체 `len(prs.slides)` 비교)
- [ ] 본문이 비어 있지 않은가?
- [ ] `samples/*.pptx` 원본은 손대지 않았는가? (`--output` 으로만 새 파일 작성)
- [ ] `releases/<TS>/<자료명>.pptx` 가 LibreOffice 로 PDF 변환되는가?
- [ ] PDF 가 28~30장 사이인가? (이전 대비 +1)

## 7. 절대 하지 말 것
- `prs.slides` 전체를 재구성하지 말 것. 기존 레이아웃·관계가 깨질 위험.
- 임의로 master/layout 의 `.rels` 를 수정하지 말 것 (PowerPoint 호환성 깨짐).
- 외부 폰트/이미지 다운로드 금지. 인터넷 접근은 RSS fetch 단계에서만 허용.
- 동영상·오디오 임베드 금지 (LFS 크기 제약).
