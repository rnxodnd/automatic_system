# 상세페이지 자동화 시스템 작업 문서

## 1. 현재 최종 워크플로우 구조

```text
입력_파일 업로드(Webhook)
↓
python 활용 excel 파일 분석(HTTP Request)
↓
기본 프롬프트 틀 생성(Code)
↓
최종 프롬프트 작성(Message a model / Gemini)
↓
Respond to Webhook
```


이유: `app.py`에서 `targetProductCode` 기준으로 상품 기본정보, 정보표, 사이즈표를 모두 파싱하므로 n8n에서 다시 품번 검색하지 않는다.

---

## 2. 주요 파일 및 경로

```text
프로젝트 루트:
 /Users/rnxodnd9/Desktop/automatic_system

n8n docker-compose.yml:
 /Users/rnxodnd9/Desktop/automatic_system/docker-compose.yml

n8n 데이터 폴더:
 /Users/rnxodnd9/Desktop/automatic_system/n8n_data

n8n DB:
 /Users/rnxodnd9/Desktop/automatic_system/n8n_data/database.sqlite

Python FastAPI 서버:
 /Users/rnxodnd9/Desktop/xlsx-parser/app.py

테스트 엑셀:
 /Users/rnxodnd9/Desktop/automatic_system/test_sample/sample.xlsx

테스트 이미지:
 /Users/rnxodnd9/Desktop/automatic_system/test_sample/sample/IMG_7638.JPG

n8n 접속:
 http://localhost:5678

Python health:
 http://localhost:8010/health

Python parser endpoint:
 http://localhost:8010/parse-xlsx

n8n test webhook:
 http://localhost:5678/webhook-test/upload
```

---

## 3. docker-compose.yml

경로:

```text
/Users/rnxodnd9/Desktop/automatic_system/docker-compose.yml
```

내용:

```yaml
services:
  n8n:
    image: n8nio/n8n:latest
    container_name: n8n_automation
    ports:
      - "5678:5678"
    volumes:
      - /Users/rnxodnd9/Desktop/automatic_system/n8n_data:/home/node/.n8n
    restart: unless-stopped
```

---

## 4. Docker 실행 명령어

### n8n 실행

```bash
cd /Users/rnxodnd9/Desktop/automatic_system
docker compose up -d
```

### n8n 종료

```bash
cd /Users/rnxodnd9/Desktop/automatic_system
docker compose down
```

### n8n 재시작

```bash
cd /Users/rnxodnd9/Desktop/automatic_system
docker compose down
docker compose up -d
```

### 컨테이너 삭제 후 재실행

```bash
cd /Users/rnxodnd9/Desktop/automatic_system
docker rm -f n8n_automation
docker compose up -d
```

### 마운트 확인

```bash
docker inspect n8n_automation | grep Mounts -A 10
```

정상 기준:

```text
Source: /Users/rnxodnd9/Desktop/automatic_system/n8n_data
Destination: /home/node/.n8n
```

---

## 5. Python 서버 실행

경로:

```text
/Users/rnxodnd9/Desktop/xlsx-parser/app.py
```

실행:

```bash
cd /Users/rnxodnd9/Desktop/xlsx-parser
python3 -m uvicorn app:app --host 0.0.0.0 --port 8010
```

정상 확인:

```bash
curl http://localhost:8010/health
```

정상 응답:

```json
{"status":"ok"}
```

---

## 6. app.py 전체 코드

```python
from fastapi import FastAPI, UploadFile, File, Form
from openpyxl import load_workbook
import tempfile
import os
import re

app = FastAPI()

@app.get("/health")
def health():
    return {"status": "ok"}

def clean_key(value):
    return re.sub(r"\s+", "", str(value or "").strip())

def to_text(value):
    if value is None:
        return ""
    return str(value).strip()

def is_dark_font(cell):
    color = cell.font.color
    if color is None:
        return True
    if color.type == "rgb" and color.rgb:
        rgb = color.rgb
        if len(rgb) == 8:
            rgb = rgb[2:]
        if len(rgb) != 6:
            return True
        r = int(rgb[0:2], 16)
        g = int(rgb[2:4], 16)
        b = int(rgb[4:6], 16)
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        return brightness < 160
    if color.type in ["theme", "indexed"]:
        return True
    return True

def parse_product_info(wb, target_code):
    code_keys = ["품번", "상품코드", "상품 코드", "productCode"]
    for ws in wb.worksheets:
        if ws.title in ["정보표", "사이즈표"]:
            continue
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [to_text(v) for v in rows[0]]
        for row in rows[1:]:
            data = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                value = row[i] if i < len(row) else None
                if value is not None and to_text(value) != "":
                    data[header] = to_text(value)
            for key in code_keys:
                for data_key, data_value in data.items():
                    if clean_key(data_key) == clean_key(key):
                        if to_text(data_value) == target_code:
                            return data
    return {}

def parse_info_table(wb, target_code):
    if "정보표" not in wb.sheetnames:
        return {}, ""
    ws = wb["정보표"]
    sections = {}
    current_code = ""
    for row in ws.iter_rows():
        cells = []
        for cell in row:
            if cell.value is None:
                continue
            value = to_text(cell.value)
            if value == "":
                continue
            cells.append({
                "value": value,
                "dark": is_dark_font(cell)
            })
        if not cells:
            continue
        first_value = cells[0]["value"]
        first_is_code = bool(re.match(r"^[A-Z0-9]{4,}", first_value))
        if first_is_code:
            current_code = first_value
            if current_code != target_code:
                continue
            if len(cells) < 2:
                continue
            key = cells[1]["value"]
            option_cells = cells[2:]
        else:
            if current_code != target_code:
                continue
            key = cells[0]["value"]
            option_cells = cells[1:]
        selected_values = [
            cell["value"]
            for cell in option_cells
            if cell["dark"]
        ]
        if key and selected_values:
            sections[key] = selected_values
    text = "\n".join(
        f"- {key}: {', '.join(values)}"
        for key, values in sections.items()
    )
    return sections, text

def parse_size_table(wb, target_code):
    if "사이즈표" not in wb.sheetnames:
        return {}, ""
    ws = wb["사이즈표"]
    result = {}
    current_code = ""
    collecting = False
    for row in ws.iter_rows(values_only=True):
        row_values = list(row)
        first_cell = row_values[0] if len(row_values) > 0 else None
        if first_cell is not None and to_text(first_cell) != "":
            current_code = to_text(first_cell)
            collecting = current_code == target_code
        if not collecting:
            continue
        if len(row_values) < 3:
            continue
        label = to_text(row_values[1])
        if not label:
            continue
        values = [
            to_text(value)
            for value in row_values[2:]
            if value is not None and to_text(value) != ""
        ]
        if values:
            result[label] = values
    text = "\n".join(
        f"- {label}: {', '.join(values)}"
        for label, values in result.items()
    )
    return result, text

@app.post("/parse-xlsx")
async def parse_xlsx(
    excel: UploadFile = File(...),
    image: UploadFile = File(None),
    targetProductCode: str = Form("")
):
    target_code = to_text(targetProductCode)
    suffix = os.path.splitext(excel.filename)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await excel.read())
        tmp_path = tmp.name
    try:
        wb = load_workbook(tmp_path, data_only=True)
        product_info = parse_product_info(wb, target_code)
        sheet3_sections, sheet3_text = parse_info_table(wb, target_code)
        size_table, size_table_text = parse_size_table(wb, target_code)
        return {
            "targetProductCode": target_code,
            "imageFileName": image.filename if image else "",
            "excelFileName": excel.filename,
            **product_info,
            "productInfo": product_info,
            "sheet3ProductCode": target_code,
            "sheet3Sections": sheet3_sections,
            "sheet3Text": sheet3_text,
            "sizeTable": size_table,
            "sizeTableText": size_table_text
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
```

---

## 7. 실행 curl

### Python parser 직접 테스트

```bash
curl -X POST "http://localhost:8010/parse-xlsx" \
  -F "excel=@/Users/rnxodnd9/Desktop/automatic_system/test_sample/sample.xlsx" \
  -F "image=@/Users/rnxodnd9/Desktop/automatic_system/test_sample/sample/IMG_7638.JPG" \
  -F "targetProductCode=DQMLJK03"
```

### n8n Webhook 테스트

n8n에서 `Listen for test event` 또는 `Execute workflow`를 누른 뒤 실행한다.

```bash
curl -X POST "http://localhost:5678/webhook-test/upload" \
  -F "excel=@/Users/rnxodnd9/Desktop/automatic_system/test_sample/sample.xlsx" \
  -F "image=@/Users/rnxodnd9/Desktop/automatic_system/test_sample/sample/IMG_7638.JPG" \
  -F "targetProductCode=DQMLJK03"
```

---

## 8. n8n 노드별 설정

### 8-1. 입력_파일 업로드

노드 타입:

```text
Webhook
```

설정:

```text
HTTP Method: POST
Path: upload
Authentication: None
Respond: Using 'Respond to Webhook' Node
Test URL: http://localhost:5678/webhook-test/upload
```

입력 form-data:

```text
excel: xlsx 파일
image: jpg/png 파일
targetProductCode: 선택 품번
```

---

### 8-2. python 활용 excel 파일 분석

노드 타입:

```text
HTTP Request
```

설정:

```text
Method: POST
URL: http://host.docker.internal:8010/parse-xlsx
Authentication: None
Send Query Parameters: OFF
Send Headers: OFF
Send Body: ON
Body Content Type: Form-Data
```

Body 1:

```text
Name: excel
Type: n8n Binary File
Input Data Field Name: excel
```

Body 2:

```text
Name: image
Type: n8n Binary File
Input Data Field Name: image
```

Body 3:

```text
Name: targetProductCode
Type: Form Data / Text
Value: {{ $json.body.targetProductCode }}
```

정상 output 필드 예시:

```text
targetProductCode
imageFileName
excelFileName
품번
색상
사이즈
소재
제조사
제조국
제조일
productInfo
sheet3ProductCode
sheet3Sections
sheet3Text
sizeTable
sizeTableText
```

---

### 8-3. 기본 프롬프트 틀 생성

노드 타입:

```text
Code
```

설정:

```text
Mode: Run Once for All Items
Language: JavaScript
```

코드:

```javascript
const 품번 = $json["품번"] || $json.targetProductCode || "";
const 색상 = $json["색상"] || "";
const 사이즈 = $json["사이즈"] || "";
const 소재 = $json["소재"] || "";
const 제조사 = $json["제조사"] || "";
const 제조국 = $json["제조국"] || "";
const 제조일 = $json["제조일"] || "";

const sheet3Text = $json.sheet3Text || "";
const sizeTableText = $json.sizeTableText || "";

const draftPrompt = `
상품 기본 정보:
- 품번: ${품번}
- 색상: ${색상}
- 사이즈: ${사이즈}
- 소재: ${소재}
- 제조사: ${제조사}
- 제조국: ${제조국}
- 제조일: ${제조일}

정보표 선택 정보:
${sheet3Text}

사이즈 실측 정보:
${sizeTableText}

작성 요청:
- 업로드된 상품 이미지와 상품 기본 정보를 함께 참고할 것
- 엑셀 파일 안에 있는 상품 정보를 사실 기반 정보로 반영할 것
- 정보표 선택 정보는 상품 착용감, 계절감, 신축성, 안감, 촉감 설명에 반영할 것
- 사이즈 실측 정보는 상세페이지 사이즈 안내에 반영할 것
- 사이즈 실측 정보는 마지막에 표로 표시할 것
- 이미지에서 확인되는 상품의 형태, 분위기, 디테일, 착용 인상을 반영할 것
- 엑셀에 없는 내용은 임의로 만들지 말 것
- 특정 스타일을 미리 단정하지 말고, 상품에 맞는 방향으로 상세페이지 프롬프트를 구성할 것
- 과장 없이 상품 특성을 잘 전달할 수 있는 방향으로 작성할 것
- 상품의 색감에 변동이 없어야 하며, 전달되는 색상이 모두 포함될 것

목적:
온라인 의류 쇼핑몰용 상세페이지 제작을 위한 프롬프트 초안 정리
`.trim();

return [
  {
    json: {
      품번,
      색상,
      사이즈,
      소재,
      제조사,
      제조국,
      제조일,
      sheet3Text,
      sizeTableText,
      draftPrompt
    }
  }
];
```

---

### 8-4. 최종 프롬프트 작성

노드 타입:

```text
Message a model
```

설정:

```text
Credential: Google Gemini(PaLM) Api account
Resource: Text
Operation: Message a Model
Model: models/gemini-3-flash-preview
Simplify Output: ON
Output Content as JSON: OFF
```

Messages:

```text
Role: User
```

Prompt:

```text
너는 온라인 의류 쇼핑몰 상세페이지 제작용 프롬프트를 작성하는 전문가다.

아래 입력 정보는 엑셀 상품 정보, 정보표 선택 정보, 사이즈 실측 정보, 업로드된 상품 이미지를 바탕으로 생성된 기초 자료다.

너의 역할은 이 기초 자료를 바탕으로, 상세페이지 제작 AI 또는 이미지 생성 AI에 바로 입력할 수 있는 “최종 상세페이지 제작 프롬프트”를 작성하는 것이다.

[입력 정보]
{{$json["draftPrompt"]}}

[작성 원칙]
1. 엑셀에 있는 정보는 사실 정보로만 사용한다.
2. 엑셀에 없는 상품명, 가격, 원산지 외 정보, 기능성, 체형 보정 효과, 고급 소재감 등은 임의로 만들지 않는다.
3. 색상은 반드시 입력 정보에 있는 색상만 사용한다.
4. 상품의 실제 색감이 변하지 않도록 작성한다.
5. 이미지에서 확인 가능한 형태, 실루엣, 디테일, 분위기만 보완 설명으로 사용한다.
6. 과장 광고, 근거 없는 표현, 허위 기능 표현은 사용하지 않는다.
7. 정보표 선택 정보는 계절감, 착용핏, 신축성, 안감, 촉감 설명에 자연스럽게 반영한다.
8. 사이즈 실측 정보는 상세페이지 안의 사이즈 안내 섹션에 반드시 포함한다.
9. 최종 결과는 “상세페이지 제작자가 바로 사용할 수 있는 프롬프트” 형태로 작성한다.
10. 소비자에게 직접 보여줄 최종 상세페이지 문구가 아니라, 상세페이지를 만들 AI에게 지시하는 제작 프롬프트로 작성한다.

[상품 정보 반영 기준]
- 품번: 반드시 포함한다.
- 색상: 입력된 색상을 모두 포함한다.
- 사이즈: 입력된 사이즈를 그대로 사용한다.
- 소재: 입력된 혼용률을 그대로 사용한다.
- 제조사, 제조국, 제조일 정보가 있으면 하단 상품 정보 영역에 반영한다.
- 상품명은 입력값에 없으면 작성하지 않는다.
- 가격은 입력값에 없으면 작성하지 않는다.

[상세페이지 구성 방식]
아래 순서로 상세페이지를 구성하도록 프롬프트를 작성한다.

1. 첫 화면 / 후킹 영역
- 상품의 핵심 인상을 짧고 명확하게 보여준다.
- 입력 정보와 이미지 기준으로 확인 가능한 분위기만 사용한다.
- 과장된 카피는 피하고, 담백한 의류 쇼핑몰 톤으로 작성한다.

2. 상품 핵심 특징 영역
- 색상, 소재, 핏, 계절감, 착용감을 중심으로 설명한다.
- 루즈핏이면 여유 있는 실루엣 중심으로 표현한다.
- 신축성이 없음이면 편안한 신축성 표현을 하지 않는다.
- 안감이 없음이면 안감 없음 정보를 그대로 반영한다.
- 촉감이 보통이면 부드럽다, 고급스럽다 같은 과장 표현은 피한다.

3. 이미지 기반 디테일 영역
- 업로드된 상품 이미지를 참고하여 형태, 넥라인, 소매, 기장감, 실루엣, 전면 디테일 등 확인 가능한 요소를 설명에 반영한다.
- 이미지에서 확인되지 않는 디테일은 만들지 않는다.

4. 컬러 안내 영역
- 입력된 모든 색상을 포함한다.
- 실제 상품 색감이 바뀌지 않도록 자연광/과보정/색감 변경 지시를 피한다.
- 입력 색상명을 그대로 사용한다.

5. 사이즈 안내 영역
- 사이즈 실측 정보를 표 또는 목록 형태로 정리한다.
- 입력된 실측 값을 그대로 사용한다.
- 임의의 추천 키, 몸무게, 체형별 추천 정보는 만들지 않는다.

6. 상품 정보 영역
- 품번, 색상, 사이즈, 소재, 제조사, 제조국, 제조일을 정리한다.
- 없는 항목은 표시하지 않는다.

[출력 형식]
아래 형식으로 출력한다.

제목:
상세페이지 제작용 최종 프롬프트

내용:
상세페이지 제작 AI에게 전달할 최종 프롬프트를 작성한다.

최종 프롬프트 안에는 다음 문장을 반드시 포함한다.
- “업로드된 상품 이미지를 기준으로 상품의 실제 형태와 색감을 유지한다.”
- “엑셀에 없는 정보는 임의로 생성하지 않는다.”
- “입력된 색상, 소재, 사이즈, 실측 정보를 사실 그대로 반영한다.”

[주의]
최종 출력에는 설명이나 분석을 덧붙이지 말고, 바로 사용할 수 있는 최종 프롬프트만 작성한다.
```

주의:

```text
Gemini 응답에 thoughtSignature가 보여도 정상이다.
이 값은 내부 메타데이터라 무시한다.
실제 결과 텍스트는 보통 content.parts[0].text에 있다.
```

---

### 8-5. Respond to Webhook

노드 타입:

```text
Respond to Webhook
```

설정:

```text
Respond With: JSON
Response Code: 200
```

Response Body:

```json
{
  "ok": true,
  "finalPrompt": "{{ $json.content.parts[0].text }}"
}
```

만약 Gemini output이 `$json.text` 형태로 나오면 아래로 변경한다.

```json
{
  "ok": true,
  "finalPrompt": "{{ $json.text }}"
}
```

---


## 9. 현재 시스템 기준 핵심 원칙

```text
1. 품번 검색은 Python(app.py)에서 처리한다.
2. n8n은 파일 전달, 프롬프트 조립, Gemini 호출만 담당한다.
3. Excel에 없는 정보는 생성하지 않는다.
4. 상품명/가격/관리자 코멘트는 현재 엑셀 기준에서 사용하지 않는다.
5. 사이즈표와 정보표는 app.py에서 같은 품번 기준으로 추출한다.
6. Gemini는 최종 프롬프트 작성 전용이다.
```
