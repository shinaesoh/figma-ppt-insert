---
name: figma-ppt-insert
description: |
  Insert Figma-exported screen images into a PowerPoint template at marker shapes
  whose text is "[화면 교체 위치: {화면ID}]", or — on screen-definition decks without
  markers — replace the existing screenshot on slides whose table has a "화면ID" cell.
  Matching is by file name = Figma frame name = screen ID.
  Fits each image into the box (similar ratio: full width, top-aligned; very different
  ratio such as popups: centered; optional fill/crop),
  removes the marker, archives the previous result, and appends a match report to
  insert_log.md. Re-running on an already-filled deck replaces the inserted screens
  in the same box. Triggers: "피그마 화면 넣어줘", "피그마 이미지 PPT에 삽입",
  "화면 교체 위치에 이미지 넣어줘", "피그마 export 반영해줘", "화면 이미지 최신으로 교체",
  "figma screens into ppt". Do NOT use for generating or redesigning slides, filling
  text into templates (ppt-template-fill), or extracting marker lists to Excel.
---

# figma-ppt-insert — 피그마 화면 자동 삽입 (1차)

피그마에서 export한 화면 이미지를 PPT 안의 마커 도형 위치·크기에 맞춰 넣는다.
비개발자용 사용 안내는 저장소 루트의 `README.md`에 있다.

## 작업 폴더

저장소 루트 (스크립트 기본값).

| 경로 | 내용 |
|---|---|
| `01_template/` | 마커가 들어간 원본 PPT. 정확히 1개 |
| `02_images/YYYY-MM-DD/` | export 날짜별 이미지 (`.png` / `.jpg`) |
| `03_output/{원본파일명}_최종.pptx` | 항상 최신 1개 |
| `03_output/_archive/` | 이전 최종본 `{원본파일명}_YYYYMMDD_HHMM.pptx`, 최근 10개 보관 |
| `insert_log.md` | 실행 이력 + 매칭 리포트 (누적, 최신이 맨 아래) |

## 규칙

| 항목 | 규칙 |
|---|---|
| 화면 ID | 피그마 프레임명 = 화면 ID (예: `SCR-MAIN-001`) |
| 파일명 | `{화면ID}.png` / `.jpg`. 피그마 배율 접미사 `@2x` 등은 자동 제거. 대소문자 무시 |
| 마커 | 텍스트 전체가 `[화면 교체 위치: {화면ID}]`인 도형(사각형·텍스트상자·placeholder). 전각 콜론·공백 차이 허용 |
| 화면ID 표 (마커 없는 슬라이드) | 표에 `화면ID` 칸이 있으면 바로 오른쪽 값이 ID. 슬라이드에서 가장 큰 그림(1 sq in 이상, 로고 제외)이 기존 화면으로 보고 그 영역·레이어에서 교체. 이미지가 없으면 기존 화면 유지 |
| 우선순위 | 슬라이드마다 마커·이전 삽입 이미지 → 없을 때만 화면ID 표 |
| 중복 ID | 같은 ID가 여러 슬라이드에 있으면 모두 같은 이미지로 교체 |
| 최신 선택 | 같은 ID가 여러 날짜 폴더에 있으면 가장 최근 날짜 폴더. 같은 폴더에 png·jpg가 둘 다 있으면 png |
| 배치 (기본 `fit`, 비율 유지) | 영역과 비율 차이 ±15% 이내(예: 16:9 ↔ 16:10)면 **상단 맞춤**: 영역 너비에 맞추고 좌상단·우상단을 영역과 일치시킴. 더 짧으면 아래 여백, 더 길면 아래쪽을 잘라 영역 안에 둠. 차이가 크면(팝업·모바일) **가운데**: 영역 안에 맞추고 가운데 정렬 |
| 배치 (`--mode fill`) | 영역을 채우고 가운데 기준으로 잘림 |
| z-순서 | 이미지는 마커가 있던 레이어 위치에 들어가고 마커는 삭제 |
| 재실행 교체 | 삽입 이미지는 이름 `FIGMA:{화면ID}`, 대체 텍스트에 원래 영역(inch)을 기록. 마커가 없어도 이 이미지를 같은 영역에서 새 export로 교체 |
| 단위 | 위치·크기 계산은 inch. 픽셀은 가로세로 비율과 해상도 점검에만 사용 |
| 해상도 점검 | 표시 크기 기준 150ppi 미만이면 `확인 필요`에 경고 (삽입은 진행) |
| 폰트 | 1차는 텍스트를 추가하지 않음. 추가 시 Pretendard |

## 실행

저장소 루트에서:

```bash
python .claude/skills/figma-ppt-insert/scripts/insert.py
```

| 옵션 | 의미 |
|---|---|
| `--mode fill` | 영역 채우기(잘림) |
| `--dry-run` | 매칭 결과만 출력. 저장·보관·로그 기록 안 함 |
| `--root <경로>` | 다른 작업 폴더 사용 |

순서: 템플릿 1개 확인 → 날짜 폴더에서 ID별 최신 이미지 수집 → 마커·기존 삽입 이미지 탐색 → 삽입 →
기존 `_최종.pptx`를 `_archive`로 이동(10개 초과분 삭제) → 새 최종본 저장 → `insert_log.md`에 기록.

## 실행 후 보고

스크립트 출력(= 이번 로그 항목)을 바탕으로 사용자에게 다음을 짧게 알린다.

1. 결과 파일 경로와 삽입 성공 개수
2. 매칭 실패: 마커는 있는데 이미지 없음 / 이미지는 있는데 마커·화면ID 없음. 대부분 파일명·프레임명 오타이므로 ID를 그대로 보여준다.
   화면ID 슬라이드 중 이미지가 없어 기존 화면을 유지한 개수도 알린다 (일부 화면만 export한 경우 정상)
3. `확인 필요` 항목(저해상도 이미지, 날짜 형식이 아닌 폴더, 그룹 안 마커, 회전된 마커, 중복 파일).
   저해상도 경고가 있으면 피그마에서 배율 2x로 다시 export하도록 안내한다

오류 처리:

- `01_template` 파일이 0개 또는 2개 이상이면 → 1개만 남기도록 안내
- 쓰기 권한 오류 → PowerPoint에서 `_최종.pptx`를 닫고 다시 실행하도록 안내

## 1차 범위 밖

텍스트 자동 채우기, 다중 템플릿, 그룹 안 마커, 회전 반영, ID 표시 없는 수동 배치 이미지 인식 등은 구현하지 않는다.
요청이 들어오면 `README.md`의 '2차 후보'에 추가하고 사용자에게 범위 밖임을 알린다.
