from __future__ import annotations

from datetime import date
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class TransportType(str, Enum):
    RAIL = "철도"
    SRT = "SRT"
    EXPRESS_BUS = "고속버스"
    EXPRESS_BUS_HONAM = "고속(호남)버스"
    KOREAN_AIR = "대한항공"
    ASIANA = "아시아나"
    OTHER_AIR = "기타항공"
    SHIP = "선박"
    INTERCITY_BUS = "시외버스"
    PRIVATE_CAR = "자가용차량"
    HANSUNG_AIR = "한성항공"
    OFFICIAL_CAR = "공무용 차량"
    OTHER = "기타"


TRANSPORT_OPTIONS: list[str] = [t.value for t in TransportType]


TRAVEL_CLASS_SUGGESTIONS: list[str] = [
    "KTX일반실",
    "KTX특실",
    "SRT일반실",
    "SRT특실",
    "새마을호특실",
    "새마을호보통",
    "무궁화호특실",
    "무궁화호보통",
    "우등",
    "일반",
    "이코노미",
    "비즈니스",
    "일등석",
]


class Receipt(BaseModel):
    """영수증 1건에서 추출한 정산 필드. e-사람 양식 기반."""

    transport: TransportType = Field(description="교통편 (필수)")
    travel_class: Optional[str] = Field(default=None, description="등급 (예: KTX일반실)")
    quantity: int = Field(default=1, ge=1, description="매수")
    travel_date: Optional[date] = Field(default=None, description="운임일자 (YYYY-MM-DD)")
    amount: int = Field(ge=0, description="금액 (원, 총액)")
    origin: Optional[str] = Field(default=None, description="출발지")
    destination: Optional[str] = Field(default=None, description="도착지")
    source_file: Optional[str] = Field(default=None, description="원본 파일명")


class ReceiptExtraction(BaseModel):
    """LLM 응답 스키마. 파일 1개에 여러 건이 있을 수 있음 (왕복 등)."""

    receipts: list[Receipt]
