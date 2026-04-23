from __future__ import annotations

import base64
import io
import os
from pathlib import Path

import fitz  # pymupdf
from openai import OpenAI

from .schema import Receipt, ReceiptExtraction

MODEL = "gpt-41-mini"

SYSTEM_PROMPT = """너는 한국 공무원 출장 교통 영수증을 판독해 e-사람 여비 정산 필드로 추출하는 어시스턴트다.

규칙:
1. 금액은 원(KRW) 정수로만. 쉼표·원 기호 제거. 총 결제금액을 사용.
2. 운임일자(travel_date)는 승차일 기준 YYYY-MM-DD.
3. 출발지·도착지는 역/터미널/공항명만 (예: "서울역"→"서울", "동대구역"→"동대구", "인천공항"→"인천").
4. 영수증 1장에 왕복 2건이 있으면 2개의 receipt 객체로 분리.
5. 판독 불가 필드는 null. 추측 금지.
6. **transport (교통편)** 는 아래 중 반드시 택1:
   - "철도"  (KTX·새마을호·무궁화호·ITX 등 코레일 열차 전부)
   - "SRT"
   - "고속버스"
   - "고속(호남)버스"
   - "대한항공", "아시아나", "기타항공", "한성항공"
   - "선박", "시외버스", "자가용차량", "공무용 차량"
   - "기타"
7. **travel_class (등급)** 예시: "KTX일반실", "KTX특실", "SRT일반실", "SRT특실",
   "새마을호특실", "새마을호보통", "무궁화호특실", "무궁화호보통",
   "우등", "일반", "이코노미", "비즈니스", "일등석".
   판독 불가 시 null.
8. **quantity (매수)** 는 영수증상 인원/매수. 미기재면 1.
9. 반드시 응답 스키마(JSON)만 출력."""

USER_INSTRUCTION = (
    "이 영수증 이미지에서 교통 탑승 내역을 모두 추출해줘. "
    "교통편은 정해진 값 중 택1, 등급은 예시 참고. 스키마에 맞춰 JSON으로만 답변."
)

PDF_DPI = 200


def _client() -> OpenAI:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY 환경변수가 설정되어 있지 않습니다.")
    return OpenAI(api_key=api_key)


def _pdf_to_png_list(data: bytes) -> list[bytes]:
    pages: list[bytes] = []
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=PDF_DPI)
            pages.append(pix.tobytes("png"))
    finally:
        doc.close()
    return pages


def _image_to_data_url(data: bytes, media_type: str) -> str:
    b64 = base64.standard_b64encode(data).decode("ascii")
    return f"data:{media_type};base64,{b64}"


def _file_to_image_data_urls(filename: str, data: bytes) -> list[str]:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        return [_image_to_data_url(img, "image/png") for img in _pdf_to_png_list(data)]
    media_map = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
    }
    media_type = media_map.get(ext, "image/jpeg")
    return [_image_to_data_url(data, media_type)]


def extract_from_file(filename: str, data: bytes) -> list[Receipt]:
    """단일 영수증 파일(이미지 또는 PDF)에서 Receipt 목록을 추출한다."""
    client = _client()
    image_urls = _file_to_image_data_urls(filename, data)

    content: list[dict] = [{"type": "text", "text": USER_INSTRUCTION}]
    for url in image_urls:
        content.append({"type": "image_url", "image_url": {"url": url}})

    completion = client.beta.chat.completions.parse(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        response_format=ReceiptExtraction,
        max_tokens=2048,
    )

    parsed = completion.choices[0].message.parsed
    if parsed is None:
        return []
    for r in parsed.receipts:
        r.source_file = filename
    return parsed.receipts
