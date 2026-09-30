# -*- coding: utf-8 -*-
"""주소 열의 아파트·빌라 동·호수를 읽어 그 평형 매매가격을 조회한다.

주소만 단지 검색에 보내고 성명·주민등록번호는 전송하지 않는다.
동·호수로 그 동의 전용·공급면적을 맞춘 뒤, KB 시세와 아실·국토교통부 실거래를 그 면적에 맞춰 산출한다.
국토교통부 자료는 호수를 공개하지 않으므로, 같은 단지·같은 전용면적·같은 층 거래를 사용한다.
"""

from __future__ import annotations

import json
import os
import queue
import re
import sys
import threading
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable

import requests
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

ZIGBANG_SEARCH = "https://apis.zigbang.com/v2/search"
ZIGBANG_DANJI = "https://apis.zigbang.com/apt/danjis/{danji_id}"
KB_LIST = "https://api.kbland.kr/land-complex/complexComm/hscmList"
KB_TYPES = "https://api.kbland.kr/land-complex/complex/typInfo"
KB_PRICE = "https://api.kbland.kr/land-price/price/WholQuotList"
KB_DONG = "https://api.kbland.kr/land-complex/complex/pubLandPriceByDong"
KB_SIGUNGU = "https://api.kbland.kr/land-complex/map/siGunGuAreaNameList"
KB_LEGAL_DONG = "https://api.kbland.kr/land-complex/map/stutDongAreaNameList"
KB_MAIN = "https://api.kbland.kr/land-complex/complex/main"
ASIL_LIST = "https://asil.kr/asil/include/apt_li.jsp"
ASIL_AREAS = "https://asil.kr/asil/apt_price_2020.jsp"
ASIL_TRADES = "https://asil.kr/app/price_detail_ver_3_9.jsp"
MOLIT_APT = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
MOLIT_HOUSE = "https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade"
MOLIT_VILLA = "https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade"
MOLIT_OFFI = "https://apis.data.go.kr/1613000/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade"
MOLIT_URLS = {"apt": MOLIT_APT, "sh": MOLIT_HOUSE, "rh": MOLIT_VILLA, "offi": MOLIT_OFFI}
HOGANG_SUGGEST = "https://hogangnono.com/api/region/suggest"

SIDO_ALIAS = {
    "서울": "서울특별시",
    "서울시": "서울특별시",
    "서울특별시": "서울특별시",
    "부산": "부산광역시",
    "부산시": "부산광역시",
    "부산광역시": "부산광역시",
    "대구": "대구광역시",
    "대구시": "대구광역시",
    "대구광역시": "대구광역시",
    "인천": "인천광역시",
    "인천시": "인천광역시",
    "인천광역시": "인천광역시",
    "광주": "광주광역시",
    "광주광역시": "광주광역시",
    "대전": "대전광역시",
    "대전시": "대전광역시",
    "대전광역시": "대전광역시",
    "울산": "울산광역시",
    "울산시": "울산광역시",
    "울산광역시": "울산광역시",
    "세종": "세종특별자치시",
    "세종시": "세종특별자치시",
    "세종특별자치시": "세종특별자치시",
    "경기": "경기도",
    "경기도": "경기도",
    "강원": "강원특별자치도",
    "강원도": "강원특별자치도",
    "강원특별자치도": "강원특별자치도",
    "충북": "충청북도",
    "충청북도": "충청북도",
    "충남": "충청남도",
    "충청남도": "충청남도",
    "전북": "전북특별자치도",
    "전라북도": "전북특별자치도",
    "전북특별자치도": "전북특별자치도",
    "전남": "전라남도",
    "전라남도": "전라남도",
    "경북": "경상북도",
    "경상북도": "경상북도",
    "경남": "경상남도",
    "경상남도": "경상남도",
    "제주": "제주특별자치도",
    "제주도": "제주특별자치도",
    "제주특별자치도": "제주특별자치도",
}

SIDO_EQUIV = {
    "서울특별시": ("서울특별시", "서울시"),
    "부산광역시": ("부산광역시", "부산시"),
    "대구광역시": ("대구광역시", "대구시"),
    "인천광역시": ("인천광역시", "인천시"),
    "광주광역시": ("광주광역시", "전남광주통합특별시", "전남광주시"),
    "대전광역시": ("대전광역시", "대전시"),
    "울산광역시": ("울산광역시", "울산시"),
    "세종특별자치시": ("세종특별자치시", "세종시"),
    "경기도": ("경기도",),
    "강원특별자치도": ("강원특별자치도", "강원도"),
    "충청북도": ("충청북도",),
    "충청남도": ("충청남도",),
    "전북특별자치도": ("전북특별자치도", "전라북도"),
    "전라남도": ("전라남도", "전남광주통합특별시", "전남광주시"),
    "경상북도": ("경상북도",),
    "경상남도": ("경상남도",),
    "제주특별자치도": ("제주특별자치도", "제주도"),
}

MATCH_VERSION = "4"

RESULT_HEADERS = [
    "동",
    "호",
    "추정층",
    "매칭단지명",
    "주택유형",
    "전용면적_㎡",
    "공급면적_㎡",
    "공급평",
    "전용평",
    "면적근거",
    "산출가격_만원",
    "산출가격_표기",
    "산출기준",
    "KB매매일반_만원",
    "KB매매하한_만원",
    "KB매매상한_만원",
    "KB기준년월",
    "아실최근실거래_만원",
    "아실계약일",
    "국토부최근실거래_만원",
    "국토부계약일",
    "국토부층",
    "호갱노노",
    "조회상태",
]

HEADER_FILL = PatternFill("solid", fgColor="FFF2CC")
PRICE_COMMENT = (
    "주소의 동·호수로 맞춘 면적의 가격입니다. "
    "KB는 그 평형의 매매 일반거래가이고, "
    "아실·국토교통부 금액은 그 전용면적의 최근 매매 실거래입니다. "
    "국토부 자료에는 호수가 없어 같은 면적과 같은 층 거래를 사용합니다."
)


@dataclass
class ParsedAddress:
    raw: str
    sido: str = ""
    sigungu: str = ""
    dong: str = ""
    road: str = ""
    road_no: str = ""
    names: list[str] = field(default_factory=list)
    unit_dong: str = ""
    unit_ho: str = ""
    unit_floor: str = ""

    @property
    def key(self) -> str:
        return re.sub(r"\s+", " ", self.raw).strip()


@dataclass
class PriceResult:
    unit_dong: str = ""
    unit_ho: str = ""
    unit_floor: str = ""
    complex_name: str = ""
    housing_type: str = ""
    exclusive_m2: str = ""
    supply_m2: str = ""
    supply_pyeong: str = ""
    exclusive_pyeong: str = ""
    area_basis: str = ""
    price_man: str = ""
    price_text: str = ""
    price_basis: str = ""
    kb_price: str = ""
    kb_low: str = ""
    kb_high: str = ""
    kb_month: str = ""
    asil_price: str = ""
    asil_date: str = ""
    molit_price: str = ""
    molit_date: str = ""
    molit_floor: str = ""
    molit_checked: str = ""
    hogang: str = ""
    status: str = ""

    def as_row(self) -> list[str]:
        return [
            self.unit_dong,
            self.unit_ho,
            self.unit_floor,
            self.complex_name,
            self.housing_type,
            self.exclusive_m2,
            self.supply_m2,
            self.supply_pyeong,
            self.exclusive_pyeong,
            self.area_basis,
            self.price_man,
            self.price_text,
            self.price_basis,
            self.kb_price,
            self.kb_low,
            self.kb_high,
            self.kb_month,
            self.asil_price,
            self.asil_date,
            self.molit_price,
            self.molit_date,
            self.molit_floor,
            self.hogang,
            self.status,
        ]


def format_manwon(value) -> str:
    number = int(round(float(value)))
    eok, man = divmod(abs(number), 10000)
    sign = "-" if number < 0 else ""
    if eok and man:
        return f"{sign}{eok}억 {man:,}만원"
    if eok:
        return f"{sign}{eok}억원"
    return f"{sign}{man:,}만원"


def pyeong_text(square_meters) -> str:
    if square_meters in (None, ""):
        return ""
    value = float(square_meters) / 3.305785
    text = f"{value:.1f}"
    if text.endswith(".0"):
        return text[:-2]
    return text


def norm_dong_no(value) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits.lstrip("0") or ("0" if digits else "")


def floor_from_ho(ho: str) -> str:
    if not ho.isdigit():
        return ""
    number = int(ho)
    if number < 100:
        return ""
    floor = number // 100
    if 1 <= floor <= 70:
        return str(floor)
    return ""


def extract_unit(text: str) -> tuple[str, str, str]:
    unit_dong = ""
    unit_ho = ""
    dongs = re.findall(r"(?<![가-힣\d])(\d{1,4})\s*동", text)
    hos = re.findall(r"(?<![가-힣\d])(\d{1,5})\s*호", text)
    if dongs:
        unit_dong = dongs[-1]
    if hos:
        unit_ho = hos[-1]
    if not unit_dong or not unit_ho:
        for match in re.finditer(r"(?<!\d)(\d{1,3})\s*-\s*(\d{3,4})(?!\d)", text):
            before = text[max(0, match.start() - 1) : match.start()]
            if before in {"로", "길"}:
                continue
            unit_dong = unit_dong or match.group(1)
            unit_ho = unit_ho or match.group(2)
            break
    return unit_dong, unit_ho, floor_from_ho(unit_ho)


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_molit_key() -> str:
    env_key = os.environ.get("MOLIT_SERVICE_KEY", "").strip()
    if env_key:
        return env_key
    path = app_dir() / "국토부인증키.txt"
    if not path.exists():
        return ""
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            return line
    return ""


def recent_year_months(count: int = 6) -> list[str]:
    year = datetime.now().year
    month = datetime.now().month
    values = []
    for _ in range(count):
        values.append(f"{year}{month:02d}")
        month -= 1
        if month == 0:
            month = 12
            year -= 1
    return values


def parse_amount(value) -> int | None:
    if value in (None, ""):
        return None
    text = str(value).strip().replace(",", "")
    if not text.isdigit():
        return None
    amount = int(text)
    if amount <= 0:
        return None
    return amount


def _households(area: dict) -> int:
    try:
        return int(area.get("세대수") or 0)
    except (TypeError, ValueError):
        return 0


def _most_households(areas: list[dict]) -> dict | None:
    usable = [area for area in areas if area.get("면적일련번호")]
    if not usable:
        return None
    return max(usable, key=_households)


def _areas_in_supply_range(areas: list[dict], low, high) -> list[dict]:
    try:
        lower = float(low) - 1
        upper = float(high) + 1
    except (TypeError, ValueError):
        return []
    chosen = []
    for area in areas:
        try:
            supply = float(area.get("공급면적") or 0)
        except (TypeError, ValueError):
            continue
        if lower <= supply <= upper:
            chosen.append(area)
    return chosen


def _area_choice(area: dict, basis: str) -> dict:
    return {
        "exclusive": _plain_number(area.get("전용면적")),
        "supply": _plain_number(area.get("공급면적")),
        "area_id": area.get("면적일련번호"),
        "basis": basis,
    }


def closest_area(areas: list[dict], supply=None, exclusive=None) -> dict | None:
    best = None
    best_gap = 9999.0
    for area in areas:
        try:
            if supply not in (None, ""):
                gap = abs(float(area.get("공급면적") or 0) - float(supply))
            elif exclusive not in (None, ""):
                gap = abs(float(area.get("전용면적") or 0) - float(exclusive))
            else:
                continue
        except (TypeError, ValueError):
            continue
        if gap < best_gap:
            best = area
            best_gap = gap
    if best is None or best_gap > 1.5:
        return None
    return best


def core_name(name: str) -> str:
    text = re.sub(r"[\s·,.\-_/]+", "", name or "")
    for word in ("아파트", "오피스텔", "연립", "다세대", "단지"):
        text = text.replace(word, "")
    text = re.sub(r"\d+단지", "", text)
    return text.lower()


def name_similarity(left: str, right: str) -> float:
    a, b = core_name(left), core_name(right)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        short, long = (a, b) if len(a) <= len(b) else (b, a)
        if len(short) >= 4 or len(short) / len(long) >= 0.5:
            return 1.0
        return 0.45
    return SequenceMatcher(None, a, b).ratio()


def parse_address(raw) -> ParsedAddress:
    text = "" if raw is None else str(raw).strip()
    if not text or text.lower() == "none":
        return ParsedAddress(text)

    text = text.replace("\u00a0", " ").replace("，", ",").replace("（", "(").replace("）", ")")
    parens = re.findall(r"\(([^)]*)\)", text)
    flat = text.replace("(", " ").replace(")", " ").replace(",", " ")
    flat = re.sub(r"\s+", " ", flat).strip()
    tokens = flat.split(" ")

    sido = ""
    index = 0
    while index < len(tokens):
        canon = SIDO_ALIAS.get(tokens[index])
        if not canon:
            break
        if sido and canon != sido:
            if canon.endswith(("광역시", "특별시", "특별자치시")):
                sido = canon
                index += 1
                continue
            break
        sido = canon
        index += 1

    sigungu = ""
    if index < len(tokens):
        current = tokens[index]
        nxt = tokens[index + 1] if index + 1 < len(tokens) else ""
        if current.endswith("시") and nxt.endswith("구") and not re.fullmatch(r"\d+구", nxt):
            sigungu = f"{current} {nxt}"
            index += 2
        elif current.endswith(("시", "군", "구")) and not re.fullmatch(r"\d+[동호구]", current):
            sigungu = current
            index += 1

    rest = tokens[index:]
    dong = ""
    for token in rest:
        if re.fullmatch(r"\d+동", token):
            continue
        if re.fullmatch(r"[가-힣]{1,10}\d*(?:동|읍|면|가)", token) or re.fullmatch(r"[가-힣]{1,10}\d*리", token):
            dong = token
            break

    road, road_no = _extract_road(" ".join(rest))
    names = _extract_names(parens, rest, road, dong, sigungu)
    unit_dong, unit_ho, unit_floor = extract_unit(flat)
    return ParsedAddress(
        text, sido, sigungu, dong, road, road_no, names, unit_dong, unit_ho, unit_floor
    )


def _extract_road(text: str) -> tuple[str, str]:
    match = re.search(
        r"([가-힣]{2,}(?:"
        r"\d{1,5}번길|"
        r"\d{0,5}대로\s*\d{1,5}길|"
        r"\d{0,4}로\s*\d{1,5}길|"
        r"\d{0,5}대로|"
        r"\d{0,4}로|"
        r"\d{0,4}길"
        r"))\s*(\d{1,5}(?:-\d{1,4})?)?",
        text,
    )
    if not match:
        return "", ""
    road = re.sub(r"\s+", "", match.group(1))
    return road, match.group(2) or ""


def looks_like_officetel(parsed: ParsedAddress) -> bool:
    raw = parsed.raw or ""
    if "아파트" in raw:
        return False
    if "오피" in raw:
        return True
    if any(word in raw for word in ("빌라", "맨션", "맨숀", "주공")):
        return False
    return bool(parsed.unit_ho) and not parsed.unit_dong


def is_individual_unit(parsed: ParsedAddress) -> bool:
    raw = parsed.raw or ""
    if "아파트" in raw or parsed.unit_dong:
        return False
    return bool(parsed.unit_ho) or "오피" in raw


def _extract_names(parens: list[str], rest: list[str], road: str, dong: str, sigungu: str) -> list[str]:
    admin = {dong, road, *sigungu.split()}
    admin.discard("")

    def clean(phrase: str) -> str:
        kept: list[str] = []
        for token in phrase.replace(",", " ").split():
            if token in admin or token in SIDO_ALIAS:
                continue
            if re.fullmatch(r"\d+동", token) or re.fullmatch(r"\d+호", token) or re.fullmatch(r"\d+층", token):
                continue
            if re.fullmatch(r"[A-Za-z가-힣]동", token) or re.fullmatch(r"\d+길", token):
                continue
            road_compact = road.replace(" ", "")
            token_compact = token.replace(" ", "")
            if road_compact and (
                token_compact == road_compact
                or road_compact.startswith(token_compact)
                or token_compact.startswith(road_compact)
            ):
                continue
            if re.fullmatch(r"\d+(?:-\d+)?", token):
                continue
            if token in {"번지", "번길"} or token.endswith(("번지", "번길")):
                continue
            if road and (token == road or token.startswith(road)):
                continue
            kept.append(token)
        phrase = " ".join(kept).strip()
        phrase = re.sub(r"\s*(아파트|오피스텔)\s*$", "", phrase).strip()
        return phrase

    names: list[str] = []
    for phrase in parens:
        name = clean(phrase)
        if len(re.sub(r"[^가-힣a-zA-Z]", "", name)) >= 2:
            names.append(name)
    leftover = clean(" ".join(rest))
    if leftover and leftover not in names and len(re.sub(r"[^가-힣a-zA-Z]", "", leftover)) >= 2:
        names.append(leftover)

    unique: list[str] = []
    for name in names:
        if name not in unique:
            unique.append(name)
    return unique


def search_queries(parsed: ParsedAddress) -> list[str]:
    loc = parsed.sigungu or parsed.sido
    queries: list[str] = []
    for name in parsed.names:
        compact = re.sub(r"\s+", "", name)
        spaced = name.strip()
        for candidate in (f"{loc} {spaced}".strip(), spaced, compact, f"{loc} {compact}".strip()):
            if candidate and candidate not in queries:
                queries.append(candidate)
    if parsed.dong:
        for name in parsed.names[:2]:
            candidate = f"{parsed.dong} {name}".strip()
            if candidate and candidate not in queries:
                queries.append(candidate)
    if parsed.road:
        road = f"{parsed.road} {parsed.road_no}".strip()
        for candidate in (f"{loc} {road}".strip(), road):
            if candidate and candidate not in queries:
                queries.append(candidate)
    return queries[:8]


def _number_matches(text: str, number: str) -> bool:
    compact = text.replace(" ", "")
    if "-" in number:
        return number in compact
    return re.search(rf"(?<!\d){re.escape(number)}(?![-.\d])", compact) is not None


def _sido_matches(sido: str, blob: str) -> bool:
    if not sido:
        return True
    return any(name and name in blob for name in SIDO_EQUIV.get(sido, (sido,)))


def accept_candidate(item: dict, parsed: ParsedAddress) -> bool:
    kind = item.get("type") or item.get("hint") or ""
    if kind and kind not in {"apartment", "officetel", "villa", "아파트", "오피스텔", "빌라"}:
        return False
    source = item.get("_source") or {}
    blob = " ".join(
        str(part)
        for part in (
            source.get("local1"),
            source.get("local2"),
            source.get("local3"),
            item.get("description"),
            source.get("신주소"),
            item.get("name"),
        )
        if part
    )
    if not _sido_matches(parsed.sido, blob):
        return False
    for part in parsed.sigungu.split():
        if part and part not in blob:
            return False

    road_text = f"{source.get('신주소') or ''} {item.get('description') or ''}"
    road_ok = False
    if parsed.road and parsed.road.replace(" ", "") in road_text.replace(" ", ""):
        road_ok = not parsed.road_no or _number_matches(road_text, parsed.road_no)

    name_ok = False
    official = str(item.get("name") or "")
    if parsed.names:
        name_ok = max(name_similarity(name, official) for name in parsed.names) >= 0.72

    if parsed.names and parsed.road:
        return name_ok or road_ok
    if parsed.names:
        return name_ok
    if parsed.road and parsed.road_no:
        return road_ok
    return False


def candidate_score(item: dict, parsed: ParsedAddress) -> float:
    kind = item.get("type") or item.get("hint") or ""
    if kind and kind not in {"apartment", "officetel", "villa", "아파트", "오피스텔", "빌라"}:
        return -1
    source = item.get("_source") or {}
    blob = " ".join(
        str(part)
        for part in (
            source.get("local1"),
            source.get("local2"),
            source.get("local3"),
            item.get("description"),
            source.get("신주소"),
            item.get("name"),
        )
        if part
    )
    if parsed.sido and not _sido_matches(parsed.sido, blob):
        return -1
    score = 0.0
    parts = [part for part in parsed.sigungu.split() if part]
    if parts and not any(part in blob for part in parts):
        return -1
    score += 2 * sum(1 for part in parts if part in blob)
    official = str(item.get("name") or "")
    if parsed.names:
        score += max(name_similarity(name, official) for name in parsed.names) * 10
    road_text = f"{source.get('신주소') or ''} {item.get('description') or ''}"
    if parsed.road and parsed.road.replace(" ", "") in road_text.replace(" ", ""):
        score += 4
        if not parsed.road_no or _number_matches(road_text, parsed.road_no):
            score += 3
    if parsed.dong and parsed.dong in blob:
        score += 2
    return score


class PriceLookup:
    def __init__(self, cache_path: Path, pause_seconds: float = 0.12):
        self.cache_path = cache_path
        self.pause_seconds = pause_seconds
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
                ),
                "Accept": "application/json, text/javascript, */*;q=0.8",
            }
        )
        self.cache = {
            "search": {},
            "danji": {},
            "kb_list": {},
            "kb_area": {},
            "kb_dong": {},
            "price": {},
            "asil_list": {},
            "asil_meta": {},
            "asil_trade": {},
            "molit": {},
            "result": {},
            "meta": {},
            "kb_sigungu": {},
            "kb_emd": {},
            "kb_main": {},
        }
        self.molit_key = load_molit_key()
        self._molit_disabled: set[str] = set()
        self._load_cache()
        self._unsaved = 0

    def _load_cache(self) -> None:
        if not self.cache_path.exists():
            return
        try:
            loaded = json.loads(self.cache_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        if isinstance(loaded, dict):
            for key in self.cache:
                if isinstance(loaded.get(key), dict):
                    self.cache[key] = loaded[key]

    def save_cache(self) -> None:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.cache_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self.cache, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary, self.cache_path)
        self._unsaved = 0

    def _touch_cache(self) -> None:
        self._unsaved += 1
        if self._unsaved >= 15:
            self.save_cache()

    def _get_response(self, url: str, params: dict | None = None, referer: str = "", timeout: int = 20):
        headers = {"Referer": referer} if referer else None
        last_error = "조회 실패"
        for attempt in range(3):
            try:
                response = self.session.get(url, params=params, headers=headers, timeout=timeout)
                if response.status_code == 200:
                    time.sleep(self.pause_seconds)
                    return response
                last_error = f"HTTP {response.status_code}"
            except requests.RequestException as exc:
                last_error = str(exc)
            time.sleep(0.4 * (attempt + 1))
        raise RuntimeError(last_error)

    def _get_json(self, url: str, params: dict | None = None, referer: str = "") -> dict:
        return self._get_response(url, params, referer).json()

    def _get_text(self, url: str, params: dict | None = None, referer: str = "", encoding: str = "utf-8") -> str:
        response = self._get_response(url, params, referer)
        response.encoding = encoding
        return response.text

    def lookup(self, address) -> PriceResult:
        parsed = parse_address(address)
        if not parsed.key:
            return PriceResult(status="주소 없음")
        cached = self.cache["result"].get(parsed.key)
        if isinstance(cached, dict) and cached.get("match_version") == MATCH_VERSION:
            fields = {key: cached[key] for key in PriceResult.__dataclass_fields__ if key in cached}
            try:
                previous = PriceResult(**fields)
            except TypeError:
                previous = None
            if previous and previous.price_man:
                pending_molit = bool(self.molit_key) and not previous.molit_price and previous.molit_checked != "1"
                unmatched = not previous.complex_name
                if not pending_molit and not unmatched:
                    return previous
        result = self._lookup_parsed(parsed)
        stored = result.__dict__.copy()
        stored["match_version"] = MATCH_VERSION
        self.cache["result"][parsed.key] = stored
        self._touch_cache()
        return result

    def _lookup_parsed(self, parsed: ParsedAddress) -> PriceResult:
        result = PriceResult(
            unit_dong=parsed.unit_dong,
            unit_ho=parsed.unit_ho,
            unit_floor=parsed.unit_floor,
            hogang=self._hogang_note(),
        )
        queries = search_queries(parsed)
        candidate = self._find_candidate(queries, parsed) if queries else None
        kb = None
        bcode = ""
        official_name = ""
        housing = ""
        if candidate is None:
            kb, bcode = self._kb_fallback(parsed)
            if not kb:
                result.status = (
                    "아파트·빌라 이름 또는 도로명 번호가 없어 단지를 특정할 수 없음"
                    if not queries
                    else "일치하는 아파트·빌라 단지를 찾지 못함"
                )
                return self._finish(result, parsed, "", "")
            official_name = str(kb.get("단지명") or "")
            housing = str(kb.get("매물종별구분명") or "아파트")
            result.complex_name = official_name
            result.housing_type = housing

        if candidate is not None:
            try:
                danji = self._danji(candidate.get("id"))
            except RuntimeError as exc:
                result.complex_name = str(candidate.get("name") or "")
                result.status = f"단지 상세 조회 실패: {exc}"
                return self._finish(result, parsed, result.complex_name, "")

            bcode = _bcode(danji)
            official_name = str(danji.get("name") or candidate.get("name") or "")
            housing = _housing_label(candidate.get("type") or danji.get("real_type") or "")
            result.complex_name = official_name
            result.housing_type = housing
            if not bcode:
                result.status = "법정동 코드를 찾지 못함"
                return self._finish(result, parsed, official_name, "")

            try:
                kb = _pick_kb_complex(self._kb_complexes(bcode), official_name, str(danji.get("bunji") or ""))
            except RuntimeError:
                kb = None
            if kb is None:
                kb, fallback_bcode = self._kb_fallback(parsed)
                if fallback_bcode:
                    bcode = fallback_bcode
        if kb:
            result.complex_name = str(kb.get("단지명") or official_name)
            if kb.get("매물종별구분명"):
                result.housing_type = str(kb.get("매물종별구분명"))

        asil = None
        try:
            asil = _pick_asil(self._asil_complexes(bcode), official_name, parsed.names)
        except RuntimeError:
            asil = None
        if asil and asil.get("villa"):
            result.housing_type = "빌라"

        try:
            area_info = self._resolve_area(parsed, kb, asil)
        except RuntimeError as exc:
            area_info = {"basis": f"면적 조회 실패: {exc}"}
        result.exclusive_m2 = area_info.get("exclusive") or ""
        result.supply_m2 = area_info.get("supply") or ""
        result.supply_pyeong = pyeong_text(result.supply_m2)
        result.exclusive_pyeong = pyeong_text(result.exclusive_m2)
        result.area_basis = area_info.get("basis") or ""

        if kb and area_info.get("area_id"):
            try:
                quote = self._kb_quote(kb["단지기본일련번호"], area_info["area_id"])
            except RuntimeError:
                quote = None
            if quote:
                result.kb_price = quote["price"]
                result.kb_low = quote["low"]
                result.kb_high = quote["high"]
                result.kb_month = quote["year_month"]
        if kb and not result.kb_price:
            try:
                extras = self._kb_areas(kb["단지기본일련번호"])
            except RuntimeError:
                extras = []
            for area in sorted(extras, key=_households, reverse=True):
                area_id = area.get("면적일련번호")
                if not area_id or str(area_id) == str(area_info.get("area_id")):
                    continue
                try:
                    quote = self._kb_quote(kb["단지기본일련번호"], area_id)
                except RuntimeError:
                    continue
                if not quote:
                    continue
                choice = _area_choice(area, f"{result.area_basis or '단지'} 중 시세가 있는 평형")
                result.exclusive_m2 = choice["exclusive"]
                result.supply_m2 = choice["supply"]
                result.supply_pyeong = pyeong_text(result.supply_m2)
                result.exclusive_pyeong = pyeong_text(result.exclusive_m2)
                result.area_basis = choice["basis"]
                result.kb_price = quote["price"]
                result.kb_low = quote["low"]
                result.kb_high = quote["high"]
                result.kb_month = quote["year_month"]
                break

        if asil and result.exclusive_m2:
            building = asil.get("building") or "apt"
            try:
                trade = self._asil_trade(asil["id"], building, result.exclusive_m2, bcode)
            except RuntimeError:
                trade = None
            if trade:
                if trade.get("price"):
                    result.asil_price = str(trade["price"])
                date_text = trade.get("date") or ""
                if trade.get("floor"):
                    date_text = f"{date_text} {trade['floor']}층".strip()
                if not trade.get("price"):
                    date_text = f"{date_text} 금액암호화".strip()
                result.asil_date = date_text

        if result.exclusive_m2 and self.molit_key:
            try:
                molit = self._molit_trade(
                    bcode[:5],
                    official_name,
                    result.exclusive_m2,
                    parsed.unit_dong,
                    parsed.unit_floor,
                    result.housing_type,
                )
            except RuntimeError:
                molit = None
            if molit:
                result.molit_price = str(molit["price"])
                result.molit_date = molit.get("date") or ""
                result.molit_floor = molit.get("floor") or ""

        _apply_chosen_price(result)
        if result.price_man:
            result.status = f"완료 ({result.area_basis})" if result.area_basis else "완료"
        elif "실패" not in result.area_basis and "확정하지" in result.area_basis:
            result.status = result.area_basis
        elif not result.exclusive_m2:
            result.status = result.area_basis or "동·호에 해당하는 면적을 확정하지 못함"
        else:
            result.status = "해당 평형의 매매가격을 찾지 못함"
        return self._finish(result, parsed, official_name or result.complex_name, bcode)

    def _find_candidate(self, queries: list[str], parsed: ParsedAddress) -> dict | None:
        best = None
        best_score = -1.0
        chosen = None
        chosen_score = -1.0
        for query in queries:
            items = self._search(query)[:8]
            for item in items:
                score = candidate_score(item, parsed)
                if score > best_score:
                    best = item
                    best_score = score
                if accept_candidate(item, parsed) and score > chosen_score:
                    chosen = item
                    chosen_score = score
        if chosen is not None:
            return chosen
        if best is not None and (best.get("type") or "") in {"officetel", "villa", "오피스텔", "빌라"} and best_score >= 6:
            return best
        if best is not None and best_score >= 8:
            return best
        return None

    def _kb_sigungu_rows(self, sido: str) -> list[dict]:
        if sido in self.cache["kb_sigungu"]:
            return self.cache["kb_sigungu"][sido]
        payload = self._get_json(KB_SIGUNGU, {"시도명": sido}, "https://kbland.kr/")
        rows = (payload.get("dataBody") or {}).get("data") or []
        if not isinstance(rows, list):
            rows = []
        self.cache["kb_sigungu"][sido] = rows
        self._touch_cache()
        return rows

    def _kb_legal_dongs(self, sido: str, sigungu_name: str) -> list[dict]:
        key = f"{sido}|{sigungu_name}"
        cached = self.cache["kb_emd"].get(key)
        if cached:
            return cached
        payload = self._get_json(
            KB_LEGAL_DONG,
            {"시도명": sido, "시군구명": sigungu_name},
            "https://kbland.kr/",
        )
        rows = (payload.get("dataBody") or {}).get("data") or []
        if not isinstance(rows, list):
            rows = []
        if rows:
            self.cache["kb_emd"][key] = rows
            self._touch_cache()
        return rows

    def _legal_dong_rows(self, parsed: ParsedAddress, sigungu_name: str) -> list[dict]:
        for sido_name in SIDO_EQUIV.get(parsed.sido, (parsed.sido,)):
            rows = self._kb_legal_dongs(sido_name, sigungu_name)
            if rows:
                return rows
        return []

    def _kb_road(self, complex_id) -> str:
        key = str(complex_id or "")
        if not key:
            return ""
        cached = self.cache["kb_main"].get(key)
        if isinstance(cached, str):
            return cached
        payload = self._get_json(KB_MAIN, {"단지기본일련번호": complex_id}, "https://kbland.kr/")
        data = (payload.get("dataBody") or {}).get("data") or {}
        road = str(data.get("신주소") or "") if isinstance(data, dict) else ""
        self.cache["kb_main"][key] = road
        self._touch_cache()
        return road

    def _sigungu_options(self, parsed: ParsedAddress) -> list[tuple[str, str, str]]:
        if not parsed.sido or not parsed.sigungu:
            return []
        target = parsed.sigungu.replace(" ", "")
        for sido_name in SIDO_EQUIV.get(parsed.sido, (parsed.sido,)):
            try:
                rows = self._kb_sigungu_rows(sido_name)
            except RuntimeError:
                continue
            matched = []
            for row in rows:
                name = str(row.get("시군구명") or "").replace(" ", "")
                if not name:
                    continue
                if name == target or target.endswith(name) or name.endswith(target) or name.startswith(target):
                    matched.append((sido_name, str(row.get("시군구명") or ""), str(row.get("법정동코드") or "")))
            if not matched:
                continue
            exact = [item for item in matched if item[1].replace(" ", "") == target]
            if exact:
                return exact
            strict = [
                item
                for item in matched
                if target.endswith(item[1].replace(" ", "")) or item[1].replace(" ", "").endswith(target)
            ]
            prefixed = [item for item in matched if item[1].replace(" ", "").startswith(target)]
            if strict and not prefixed:
                return strict[:1]
            return matched
        return []

    def _legal_bcode(self, parsed: ParsedAddress) -> str:
        options = self._sigungu_options(parsed)
        if not options:
            return ""
        if not parsed.dong:
            return options[0][2] if len(options) == 1 else ""
        aliases = set(legal_dong_aliases(parsed.dong))
        for _sido_name, sigungu_name, sigungu_code in options:
            try:
                dongs = self._legal_dong_rows(parsed, sigungu_name)
            except RuntimeError:
                continue
            for row in dongs:
                if str(row.get("법정동명") or "").replace(" ", "") in aliases:
                    return str(row.get("법정동코드") or "") or sigungu_code
        if len(options) == 1:
            return options[0][2]
        return ""

    def _approximate_lawds(self, parsed: ParsedAddress, bcode: str) -> list[str]:
        if len(bcode) >= 5:
            return [bcode[:5]]
        options = self._sigungu_options(parsed)
        resolved = self._legal_bcode(parsed)
        if parsed.dong and resolved and not resolved.endswith("00000"):
            return [resolved[:5]]
        lawds = []
        for _sido, _name, code in options:
            if len(code) >= 5 and code[:5] not in lawds:
                lawds.append(code[:5])
        if lawds:
            return lawds
        if len(resolved) >= 5:
            return [resolved[:5]]
        return []

    def _molit_deals(self, lawd: str, kinds: tuple[str, ...]) -> list[dict]:
        deals: list[dict] = []
        for kind in kinds:
            if kind in self._molit_disabled:
                continue
            for year_month in recent_year_months(6):
                deals.extend(self._molit_month(lawd, year_month, kind))
        return deals

    def _fill_approximate(self, result: PriceResult, parsed: ParsedAddress, name: str, bcode: str) -> None:
        if result.price_man or not self.molit_key:
            return
        lawds = self._approximate_lawds(parsed, bcode)
        if not lawds:
            return
        query_name = name or (parsed.names[0] if parsed.names else "")
        if query_name.replace(" ", "") in {"단독", "다가구", "주택", "단독주택", "다가구주택", "단독다가구"}:
            query_name = ""
        house_text = f"{parsed.raw} {result.housing_type}"
        explicit_officetel = "오피" in (parsed.raw or "") or "오피" in (result.housing_type or "")
        individual = not result.complex_name and is_individual_unit(parsed)
        house_first = any(word in house_text for word in ("단독", "다가구", "주택", "빌라")) and "아파트" not in house_text and not explicit_officetel
        attempts = approximate_kinds(
            explicit_officetel and not result.complex_name,
            individual,
            house_first,
        )
        summary = None
        used_house = False
        used_villa = False
        used_officetel = False
        for kinds in attempts:
            for lawd in lawds:
                deals = self._molit_deals(lawd, kinds)
                summary = summarize_deals(deals, query_name, parsed.dong, result.exclusive_m2)
                if summary is None and query_name and result.complex_name:
                    summary = summarize_deals(deals, query_name, "", result.exclusive_m2)
                if summary is None and query_name and not result.complex_name:
                    summary = summarize_deals(deals, "", parsed.dong, result.exclusive_m2)
                if summary is None and not result.complex_name and kinds in {("sh",), ("rh",)}:
                    summary = summarize_deals(deals, "", "", result.exclusive_m2)
                if result.complex_name:
                    if summary is None and parsed.dong:
                        summary = summarize_deals(deals, "", parsed.dong, result.exclusive_m2)
                    if summary is None:
                        summary = summarize_deals(deals, "", "", result.exclusive_m2)
                elif summary and not _usable_unmatched_summary(summary, parsed, individual, kinds):
                    summary = None
                if summary and kinds == ("sh",) and not unit_house_price_ok(summary["price"]):
                    if individual or "시군구" in summary["basis"]:
                        summary = None
                if summary:
                    used_house = kinds == ("sh",)
                    used_villa = kinds == ("rh",)
                    used_officetel = kinds == ("offi",)
                    break
            if summary:
                break
        if not summary:
            if (individual or explicit_officetel) and not result.complex_name:
                result.housing_type = result.housing_type or ("오피스텔" if explicit_officetel else "주택")
                result.price_man = ""
                result.price_text = ""
                result.price_basis = ""
                result.status = "개별 호실이라 건물 전체 거래가와 아파트 구 평균가는 넣지 않음"
            return
        result.price_man = str(summary["price"])
        result.price_text = format_manwon(summary["price"])
        result.price_basis = summary["basis"]
        result.molit_price = str(summary["price"])
        latest = summary.get("latest") or ""
        result.molit_date = f"{latest} {summary['count']}건 중앙값".strip()
        if used_officetel:
            result.price_basis = result.price_basis.replace("국토부", "국토부 오피스텔", 1)
            result.housing_type = result.housing_type or "오피스텔"
            result.status = f"대략 ({result.price_basis})"
        elif used_house:
            result.price_basis = result.price_basis.replace("국토부", "국토부 단독·다가구", 1)
            result.housing_type = result.housing_type or "주택"
            result.status = f"대략 ({result.price_basis})"
        elif used_villa and not result.complex_name:
            result.price_basis = result.price_basis.replace("국토부", "국토부 연립·다세대", 1)
            result.housing_type = result.housing_type or "연립·다세대"
            result.status = f"대략 ({result.price_basis})"
        if not result.area_basis:
            result.area_basis = "개별 면적 미상"
        if not result.status.startswith("대략"):
            result.status = f"대략 ({result.price_basis})"

    def _finish(self, result: PriceResult, parsed: ParsedAddress, name: str, bcode: str) -> PriceResult:
        if not result.price_man:
            self._fill_approximate(result, parsed, name, bcode)
        if self.molit_key:
            result.molit_checked = "1"
        elif "인증키 없음" not in (result.status or ""):
            result.status = f"{result.status} / 국토부 인증키 없음".strip()
        return result

    def _kb_fallback(self, parsed: ParsedAddress) -> tuple[dict | None, str]:
        if not parsed.names or not parsed.sido or not parsed.sigungu or not parsed.dong:
            return None, ""
        options = self._sigungu_options(parsed)
        if not options:
            return None, ""
        aliases = set(legal_dong_aliases(parsed.dong))
        best = None
        best_score = 0.0
        best_code = ""
        for _chosen_sido, sigungu_name, _sigungu_code in options:
            try:
                dongs = self._legal_dong_rows(parsed, sigungu_name)
            except RuntimeError:
                continue
            bcode = ""
            for row in dongs:
                if str(row.get("법정동명") or "").replace(" ", "") in aliases:
                    bcode = str(row.get("법정동코드") or "")
                    break
            if not bcode:
                continue
            try:
                complexes = self._kb_complexes(bcode)
            except RuntimeError:
                continue
            for name in parsed.names:
                picked = _pick_kb_complex(complexes, name, "")
                if not picked:
                    continue
                if parsed.road and parsed.road_no:
                    road = self._kb_road(picked.get("단지기본일련번호"))
                    compact = road.replace(" ", "")
                    same_road = parsed.road.replace(" ", "") in compact
                    if road and (not same_road or not _number_matches(road, parsed.road_no)):
                        continue
                score = name_similarity(name, str(picked.get("단지명") or ""))
                query = core_name(name)
                target = core_name(str(picked.get("단지명") or ""))
                if query and target and len(query) >= 3 and query in target:
                    score = max(score, 0.9)
                if score > best_score:
                    best = picked
                    best_score = score
                    best_code = bcode
        if best is None:
            return None, ""
        return best, best_code

    def _search(self, query: str) -> list[dict]:
        if query in self.cache["search"]:
            return self.cache["search"][query]
        payload = self._get_json(ZIGBANG_SEARCH, {"leaseYn": "N", "q": query}, "https://www.zigbang.com/")
        items = payload.get("items") or []
        if not isinstance(items, list):
            items = []
        self.cache["search"][query] = items
        self._touch_cache()
        return items

    def _danji(self, danji_id) -> dict:
        key = str(danji_id)
        if key in self.cache["danji"]:
            return self.cache["danji"][key]
        payload = self._get_json(ZIGBANG_DANJI.format(danji_id=danji_id), referer="https://www.zigbang.com/")
        self.cache["danji"][key] = payload
        self._touch_cache()
        return payload

    def _kb_complexes(self, bcode: str) -> list[dict]:
        if bcode in self.cache["kb_list"]:
            return self.cache["kb_list"][bcode]
        payload = self._get_json(KB_LIST, {"법정동코드": bcode}, "https://kbland.kr/")
        body = payload.get("dataBody") or {}
        data = body.get("data") or []
        if not isinstance(data, list):
            data = []
        self.cache["kb_list"][bcode] = data
        self._touch_cache()
        return data

    def _kb_areas(self, complex_id) -> list[dict]:
        key = str(complex_id)
        if key in self.cache["kb_area"]:
            return self.cache["kb_area"][key]
        payload = self._get_json(KB_TYPES, {"단지기본일련번호": complex_id}, "https://kbland.kr/")
        areas = (payload.get("dataBody") or {}).get("data") or []
        if not isinstance(areas, list):
            areas = []
        self.cache["kb_area"][key] = areas
        self._touch_cache()
        return areas

    def _kb_dongs(self, complex_id) -> list[dict]:
        key = str(complex_id)
        if key in self.cache["kb_dong"]:
            return self.cache["kb_dong"][key]
        payload = self._get_json(KB_DONG, {"단지기본일련번호": complex_id}, "https://kbland.kr/")
        body = payload.get("dataBody") or {}
        rows = body.get("data") or []
        if not isinstance(rows, list):
            rows = []
        self.cache["kb_dong"][key] = rows
        self._touch_cache()
        return rows

    def _kb_quote(self, complex_id, area_id) -> dict | None:
        key = f"{complex_id}:{area_id}"
        if key in self.cache["price"]:
            return self.cache["price"][key] or None
        year = datetime.now().year
        payload = self._get_json(
            KB_PRICE,
            {
                "단지기본일련번호": complex_id,
                "면적일련번호": area_id,
                "기준년": f"{year},{year - 1}",
            },
            "https://kbland.kr/",
        )
        quote = _latest_sale((payload.get("dataBody") or {}).get("data") or {})
        found = None
        if quote and quote.get("매매일반거래가"):
            found = {
                "price": str(int(quote["매매일반거래가"])),
                "low": str(int(quote["매매하한가"])) if quote.get("매매하한가") else "",
                "high": str(int(quote["매매상한가"])) if quote.get("매매상한가") else "",
                "year_month": str(quote.get("기준년월") or ""),
            }
        self.cache["price"][key] = found or {}
        self._touch_cache()
        return found

    def _resolve_area(self, parsed: ParsedAddress, kb: dict | None, asil: dict | None) -> dict:
        areas = self._kb_areas(kb["단지기본일련번호"]) if kb else []
        dongs = self._kb_dongs(kb["단지기본일련번호"]) if kb else []
        target = norm_dong_no(parsed.unit_dong)
        if target and dongs:
            matched = [row for row in dongs if norm_dong_no(row.get("동명")) == target]
            if len(matched) == 1:
                row = matched[0]
                supply_min = row.get("최소공급면적")
                supply_max = row.get("최대공급면적")
                try:
                    single_type = abs(float(supply_min) - float(supply_max)) <= 0.2
                except (TypeError, ValueError):
                    single_type = False
                if single_type:
                    area = closest_area(areas, supply=supply_min)
                    if area:
                        return _area_choice(
                            area,
                            f"{parsed.unit_dong}동 {parsed.unit_ho}호 전용 {_plain_number(area.get('전용면적'))}㎡".strip(),
                        )
                ranged = _areas_in_supply_range(areas, supply_min, supply_max)
                area = _most_households(ranged)
                if area is None:
                    try:
                        midpoint = (float(supply_min) + float(supply_max)) / 2
                    except (TypeError, ValueError):
                        midpoint = None
                    area = closest_area(areas, supply=midpoint) if midpoint is not None else _most_households(areas)
                if area:
                    return _area_choice(
                        area,
                        (
                            f"{parsed.unit_dong}동 공급 {supply_min}~{supply_max}㎡ 중 "
                            f"세대수 최다 전용 {_plain_number(area.get('전용면적'))}㎡"
                        ),
                    )
        if len(areas) == 1:
            area = areas[0]
            basis = "단지에 면적 유형이 하나"
            if target:
                basis = f"{parsed.unit_dong}동, 단지 단일 면적"
            return _area_choice(area, basis)
        if asil and not areas:
            asil_areas = self._asil_areas(asil["id"], asil.get("building") or "apt")
            if len(asil_areas) == 1:
                return {
                    "exclusive": asil_areas[0]["exclusive"],
                    "supply": "",
                    "area_id": None,
                    "basis": "아실 기준 단일 면적",
                }
        area = _most_households(areas)
        if area:
            if not target:
                basis = "동·호수 없음, 세대수가 가장 많은 평형"
            elif not dongs:
                basis = f"{parsed.unit_dong}동 면적표 없음, 세대수가 가장 많은 평형"
            else:
                basis = "해당 동 면적 불일치, 세대수가 가장 많은 평형"
            return _area_choice(area, basis)
        return {"basis": "면적 자료를 찾지 못함"}

    def _asil_complexes(self, bcode: str) -> list[dict]:
        if bcode in self.cache["asil_list"]:
            return self.cache["asil_list"][bcode]
        text = self._get_text(ASIL_LIST, {"code": bcode}, "https://asil.kr/", encoding="euc-kr")
        rows = []
        for match in re.finditer(r"searchApt\(this,(\d+),\s*\\'([^\\']+)\\'([\s\S]*?)</a>", text):
            chunk = match.group(3)
            villa = any(token in chunk for token in ("연립", "다세대", "buildingD"))
            rows.append(
                {
                    "id": match.group(1),
                    "name": match.group(2),
                    "villa": villa,
                    "building": "dasede" if villa else "apt",
                }
            )
        self.cache["asil_list"][bcode] = rows
        self._touch_cache()
        return rows

    def _asil_areas(self, apt_id, building: str) -> list[dict]:
        key = f"{building}:{apt_id}"
        if key in self.cache["asil_meta"]:
            return self.cache["asil_meta"][key]
        text = self._get_text(
            ASIL_AREAS,
            {"os": "pc", "building": building, "apt": apt_id},
            "https://asil.kr/",
            encoding="cp949",
        )
        areas = []
        seen = set()
        for match in re.finditer(r"setEvt\(this,\s*'([\d.]+)m2'\)[\s\S]{0,160}?(\d+)\s*평", text):
            exclusive = match.group(1)
            if exclusive in seen:
                continue
            seen.add(exclusive)
            areas.append({"exclusive": exclusive, "pyeong_label": match.group(2)})
        self.cache["asil_meta"][key] = areas
        self._touch_cache()
        return areas

    def _asil_trade(self, apt_id, building: str, exclusive: str, bcode: str) -> dict | None:
        key = f"json:{building}:{apt_id}:{exclusive}"
        if key in self.cache["asil_trade"]:
            return self.cache["asil_trade"][key] or None
        year = datetime.now().year
        best = None
        for deal_year in (year, year - 1):
            text = self._get_text(
                "https://asil.kr/app/data/apt_price_m2_newver_6.jsp",
                {
                    "sido": bcode[:2],
                    "dealmode": "1",
                    "building": building,
                    "seq": apt_id,
                    "m2": exclusive,
                    "year": str(deal_year),
                    "u": "0",
                    "start": "0",
                    "count": "20",
                    "order": "1",
                },
                "https://asil.kr/",
                encoding="utf-8",
            )
            best = _latest_asil_deal(text)
            if best:
                break
        self.cache["asil_trade"][key] = best or {}
        self._touch_cache()
        return best

    def _molit_trade(self, lawd: str, name: str, exclusive: str, unit_dong: str, unit_floor: str, housing: str) -> dict | None:
        if not self.molit_key:
            return None
        if "오피" in housing:
            kinds = ["offi"]
        elif any(word in housing for word in ("단독", "다가구", "주택")):
            kinds = ["sh"]
        elif any(word in housing for word in ("빌라", "연립", "다세대")):
            kinds = ["rh", "apt", "sh"]
        else:
            kinds = ["apt", "sh", "rh"]
        for kind in kinds:
            candidates = []
            for year_month in recent_year_months(6):
                for row in self._molit_month(lawd, year_month, kind):
                    if name_similarity(name, row.get("name") or "") < 0.72:
                        continue
                    try:
                        if abs(float(row.get("area") or 0) - float(exclusive)) > 2.0:
                            continue
                    except (TypeError, ValueError):
                        continue
                    score = 0
                    if unit_dong and norm_dong_no(row.get("dong")) == norm_dong_no(unit_dong):
                        score += 3
                    if unit_floor and str(row.get("floor") or "") == str(unit_floor):
                        score += 2
                    candidates.append((score, row.get("date") or "", row))
            if candidates:
                candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
                return candidates[0][2]
        return None

    def _molit_month(self, lawd: str, year_month: str, kind: str) -> list[dict]:
        cache_key = f"{kind}:{lawd}:{year_month}"
        cached = self.cache["molit"].get(cache_key)
        if isinstance(cached, list) and (not cached or "umd" in cached[0]):
            return cached
        if not self.molit_key:
            self.cache["molit"][cache_key] = []
            return []
        if kind in self._molit_disabled:
            return []
        url = MOLIT_URLS[kind]
        rows: list[dict] = []
        page = 1
        while page <= 4:
            try:
                response = self.session.get(
                    url,
                    params={
                        "serviceKey": self.molit_key,
                        "LAWD_CD": lawd,
                        "DEAL_YMD": year_month,
                        "pageNo": page,
                        "numOfRows": 1000,
                    },
                    headers={"Accept": "application/xml"},
                    timeout=20,
                )
            except requests.RequestException:
                return []
            if response.status_code == 403:
                self._molit_disabled.add(kind)
                return []
            if response.status_code != 200:
                return []
            parsed = _parse_molit_xml(response.content)
            rows.extend(parsed["items"])
            if page * 1000 >= parsed["total"] or not parsed["items"]:
                break
            page += 1
        self.cache["molit"][cache_key] = rows
        self._touch_cache()
        return rows

    def _hogang_note(self) -> str:
        meta = self.cache.setdefault("meta", {})
        if meta.get("hogang"):
            return meta["hogang"]
        note = "공개 검색 API 없음"
        try:
            response = self.session.get(
                HOGANG_SUGGEST,
                params={"query": "창포청구타운"},
                headers={"Referer": "https://hogangnono.com/", "Accept": "application/json"},
                timeout=8,
            )
            if response.status_code == 200:
                try:
                    payload = response.json()
                except ValueError:
                    payload = None
                if payload:
                    note = "검색 응답 있음"
            elif response.status_code == 404:
                note = "공개 검색 API 없음"
            else:
                note = f"HTTP {response.status_code}"
        except requests.RequestException:
            note = "연결 실패"
        meta["hogang"] = note
        self._touch_cache()
        return note


def _latest_sale(data: dict) -> dict | None:
    best = None
    best_month = ""
    for group in data.get("시세") or []:
        for item in group.get("items") or []:
            price = item.get("매매일반거래가")
            month = str(item.get("기준년월") or "")
            if price in (None, "", 0) or not month:
                continue
            if month > best_month:
                best = item
                best_month = month
    return best


def _bcode(danji: dict) -> str:
    bcode = str(danji.get("bcode") or "")
    if len(bcode) == 10 and bcode.isdigit():
        return bcode
    local = str(danji.get("local3_code") or "")
    if len(local) == 10 and local.isdigit():
        return local
    if len(local) == 8 and local.isdigit():
        return local + "00"
    return ""


def _housing_label(kind: str) -> str:
    return {
        "apartment": "아파트",
        "officetel": "오피스텔",
        "villa": "빌라",
        "아파트": "아파트",
        "오피스텔": "오피스텔",
        "빌라": "빌라",
    }.get(kind, kind)


def _match_sigungu_name(rows: list[dict], target: str) -> dict | None:
    best = None
    best_len = -1
    for row in rows:
        name = str(row.get("시군구명") or "").replace(" ", "")
        if not name:
            continue
        if name == target or target.endswith(name) or name.endswith(target):
            if len(name) > best_len:
                best = row
                best_len = len(name)
    return best


def _pick_kb_complex(complexes: list[dict], official_name: str, bunji: str) -> dict | None:
    best = None
    best_score = 0.0
    bunji_text = bunji.lstrip("0")
    for complex_row in complexes:
        score = name_similarity(official_name, str(complex_row.get("단지명") or ""))
        query = core_name(official_name)
        target = core_name(str(complex_row.get("단지명") or ""))
        if query and target and len(query) >= 3 and query in target:
            score = max(score, 0.9)
        row_bunji = str(complex_row.get("본번지내용") or complex_row.get("상세번지내용") or "").lstrip("0")
        if bunji_text and row_bunji and bunji_text == row_bunji:
            score += 0.35
        if score > best_score:
            best = complex_row
            best_score = score
    if best is not None and best_score >= 0.72:
        return best
    return None


def _pick_asil(rows: list[dict], official_name: str, names: list[str]) -> dict | None:
    best = None
    best_score = 0.0
    for row in rows:
        score = name_similarity(official_name, row.get("name") or "")
        for name in names:
            score = max(score, name_similarity(name, row.get("name") or ""))
        if score > best_score:
            best = row
            best_score = score
    if best is not None and best_score >= 0.72:
        return best
    return None


def _plain_number(value) -> str:
    if value in (None, ""):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    text = f"{number:.2f}".rstrip("0").rstrip(".")
    return text


def _apply_chosen_price(result: PriceResult) -> None:
    if result.kb_price:
        result.price_man = result.kb_price
        result.price_text = format_manwon(result.kb_price)
        result.price_basis = "KB부동산 해당 평형 매매일반거래가"
        return
    if result.molit_price:
        result.price_man = result.molit_price
        result.price_text = format_manwon(result.molit_price)
        result.price_basis = "국토교통부 해당 평형 최근 실거래"
        return
    if result.asil_price:
        result.price_man = result.asil_price
        result.price_text = format_manwon(result.asil_price)
        result.price_basis = "아실 해당 평형 최근 실거래"


def _latest_asil_deal(text: str) -> dict | None:
    start = text.find("[")
    if start < 0:
        return None
    try:
        payload = json.loads(text[start:])
    except json.JSONDecodeError:
        return None
    best = None
    for month_group in payload if isinstance(payload, list) else []:
        for bucket in month_group.get("val") or []:
            year_month = str(bucket.get("yyyymm") or "")
            if not year_month.isdigit():
                continue
            for day_group in bucket.get("val") or []:
                day = day_group.get("day")
                if day in (None, ""):
                    continue
                date = f"{year_month}{int(day):02d}"
                for deal in day_group.get("val") or []:
                    amount = parse_amount(deal.get("money"))
                    if amount is not None and (amount < 1000 or amount > 500000):
                        amount = None
                    candidate = {
                        "date": date,
                        "floor": str(deal.get("floor") or ""),
                        "price": amount,
                    }
                    if best is None or date > best["date"]:
                        best = candidate
    return best


def _parse_molit_xml(content: bytes) -> dict:
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        return {"total": 0, "items": []}
    result_code = ""
    for node in root.iter():
        if node.tag.endswith("resultCode") and node.text:
            result_code = node.text.strip()
            break
    if result_code and result_code not in {"00", "000", "0"}:
        return {"total": 0, "items": []}
    total = 0
    for node in root.iter():
        if node.tag.endswith("totalCount") and node.text and node.text.strip().isdigit():
            total = int(node.text.strip())
            break
    items = []
    for item in root.iter():
        if not item.tag.endswith("item"):
            continue
        row = {}
        for child in list(item):
            tag = child.tag.split("}")[-1]
            row[tag] = (child.text or "").strip()
        name = _first_value(row, "aptNm", "mhouseNm", "offiNm", "houseType", "아파트", "연립다세대", "오피스텔", "주택유형")
        area = _first_value(row, "excluUseAr", "전용면적", "totalFloorAr", "연면적")
        amount = parse_amount(_first_value(row, "dealAmount", "거래금액"))
        if not name or not area or amount is None:
            continue
        year = _first_value(row, "dealYear", "년")
        month = _first_value(row, "dealMonth", "월")
        day = _first_value(row, "dealDay", "일")
        date = ""
        if year and month and day and year.isdigit():
            date = f"{int(year):04d}{int(month):02d}{int(day):02d}"
        items.append(
            {
                "name": name,
                "area": area,
                "price": amount,
                "floor": _first_value(row, "floor", "층"),
                "dong": _first_value(row, "aptDong", "아파트동"),
                "umd": _first_value(row, "umdNm", "법정동"),
                "date": date,
            }
        )
    if not total:
        total = len(items)
    return {"total": total, "items": items}


def legal_dong_aliases(dong: str) -> list[str]:
    text = re.sub(r"\s+", "", dong or "")
    if not text:
        return []
    found = [text]
    numbered = re.fullmatch(r"([가-힣]{2,})\d+동", text)
    if numbered:
        found.append(f"{numbered.group(1)}동")
    return found


def _median_int(values: list[int]) -> int:
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) // 2


def _prefer_similar_area(rows: list[dict], exclusive: str) -> list[dict]:
    if not exclusive or len(rows) < 2:
        return rows
    try:
        target = float(exclusive)
    except (TypeError, ValueError):
        return rows
    close = []
    for row in rows:
        try:
            gap = abs(float(row.get("area") or 0) - target)
        except (TypeError, ValueError):
            continue
        if gap <= 15:
            close.append(row)
    return close or rows


def _umd_in(umd: str, aliases: set[str]) -> bool:
    compact = re.sub(r"\s+", "", umd or "")
    if compact in aliases:
        return True
    return any(alias and compact.startswith(alias) for alias in aliases)


UNIT_HOUSE_PRICE_CAP = 150000


def approximate_kinds(explicit_officetel: bool, individual: bool, house_first: bool) -> list[tuple[str, ...]]:
    if explicit_officetel:
        return [("offi",)]
    if individual:
        return [("rh",), ("sh",)]
    if house_first:
        return [("sh",), ("rh",)]
    return [("apt", "rh"), ("sh",)]


def unit_house_price_ok(price) -> bool:
    try:
        return 0 < int(price) <= UNIT_HOUSE_PRICE_CAP
    except (TypeError, ValueError):
        return False


def _usable_unmatched_summary(summary: dict, parsed: ParsedAddress, individual: bool, kinds: tuple[str, ...] = ()) -> bool:
    basis = summary.get("basis") or ""
    if "유사" in basis:
        return True
    if "시군구" in basis:
        if kinds in {("sh",), ("rh",)} and unit_house_price_ok(summary.get("price")):
            return True
        return bool(parsed.unit_dong) and not individual and not parsed.dong
    return bool(parsed.dong)


def summarize_deals(rows: list[dict], name: str, dong: str, exclusive: str) -> dict | None:
    usable = [row for row in rows if row.get("price")]
    if not usable:
        return None
    pool = usable
    dong_used = False
    if dong:
        aliases = set(legal_dong_aliases(dong))
        matched = [row for row in usable if _umd_in(row.get("umd") or "", aliases)]
        if not matched:
            return None
        pool = matched
        dong_used = True
    name_used = False
    if name:
        named = [row for row in pool if name_similarity(name, row.get("name") or "") >= 0.6]
        if named:
            pool = named
            name_used = True
        elif not dong_used:
            return None
    pool = _prefer_similar_area(pool, exclusive)
    prices = []
    for row in pool:
        try:
            prices.append(int(row["price"]))
        except (TypeError, ValueError):
            continue
    if not prices:
        return None
    count = len(prices)
    if name_used and dong_used:
        basis = f"국토부 {dong} 유사 단지 실거래 {count}건 중앙값(대략)"
    elif name_used:
        basis = f"국토부 유사 단지 실거래 {count}건 중앙값(대략)"
    elif dong_used:
        basis = f"국토부 {dong} 최근 실거래 {count}건 중앙값(대략)"
    else:
        basis = f"국토부 시군구 최근 실거래 {count}건 중앙값(대략, 동·단지 불명)"
    latest = max((str(row.get("date") or "") for row in pool), default="")
    return {"price": _median_int(prices), "count": count, "latest": latest, "basis": basis}


def _first_value(row: dict, *keys: str) -> str:
    for key in keys:
        if row.get(key):
            return str(row[key]).strip()
    return ""


def find_address_column(headers: list) -> int | None:
    ranked = []
    for index, header in enumerate(headers, start=1):
        if header is None:
            continue
        compact = str(header).replace(" ", "")
        if any(token in compact for token in ("이메일", "메일", "홈페이지")):
            continue
        if "주민등록" in compact and "주소" in compact:
            ranked.append((0, index))
        elif "현주소" in compact:
            ranked.append((1, index))
        elif "주소" in compact:
            ranked.append((2, index))
    if not ranked:
        return None
    ranked.sort()
    return ranked[0][1]


def fill_workbook(
    source: Path,
    destination: Path,
    lookup: PriceLookup,
    progress: Callable[[int, int, str], None] | None = None,
    limit: int | None = None,
    cancel: Callable[[], bool] | None = None,
) -> tuple[int, int]:
    workbook = load_workbook(source)
    sheet = None
    column = None
    for candidate in workbook.worksheets:
        headers = [cell.value for cell in next(candidate.iter_rows(min_row=1, max_row=1))]
        found = find_address_column(headers)
        if found:
            sheet = candidate
            column = found
            break
    if sheet is None or column is None:
        workbook.close()
        raise ValueError("첫 행에서 주소 열을 찾지 못했습니다. 열 이름에 '주소'가 있어야 합니다.")

    start_column = sheet.max_column + 1
    for offset, header in enumerate(RESULT_HEADERS):
        cell = sheet.cell(1, start_column + offset, header)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    price_header = RESULT_HEADERS.index("산출가격_표기")
    sheet.cell(1, start_column + price_header).comment = Comment(PRICE_COMMENT, "시세조회")

    rows = []
    for row_index in range(2, sheet.max_row + 1):
        rows.append((row_index, sheet.cell(row_index, column).value))
    if limit is not None:
        rows = rows[:limit]

    total = len(rows)
    done = 0
    found_price = 0
    for row_index, address in rows:
        if cancel and cancel():
            break
        result = lookup.lookup(address)
        for offset, value in enumerate(result.as_row()):
            sheet.cell(row_index, start_column + offset, value)
        done += 1
        if result.price_man:
            found_price += 1
        if progress:
            where = " ".join(
                part
                for part in (parse_address(address).sigungu, result.complex_name, result.price_text)
                if part
            )
            message = f"{done}/{total} {where or result.status}".strip()
            progress(done, total, message)

    sheet.auto_filter.ref = sheet.dimensions
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    lookup.save_cache()
    return done, found_price


def output_path_for(source: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_시세_{stamp}{source.suffix}")


def run_file(source: Path, limit: int | None, progress: Callable[[int, int, str], None] | None, cancel=None) -> Path:
    destination = output_path_for(source)
    cache = app_dir() / "cache" / "시세캐시_v2.json"
    lookup = PriceLookup(cache)
    fill_workbook(source, destination, lookup, progress=progress, limit=limit, cancel=cancel)
    return destination


def run_cli(argv: list[str]) -> int:
    limit = None
    path = None
    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg == "--limit" and index + 1 < len(argv):
            limit = int(argv[index + 1])
            index += 2
            continue
        if not arg.startswith("-") and path is None:
            path = Path(arg)
        index += 1
    if path is None:
        print("사용법: python 시세조회.py 엑셀파일.xlsx [--limit 20]")
        return 2
    if not path.exists():
        print(f"파일을 찾을 수 없습니다: {path}")
        return 2

    def progress(done: int, total: int, message: str) -> None:
        print(message, flush=True)

    destination = run_file(path, limit, progress)
    print(f"저장: {destination}")
    return 0


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("주소 동·호수 평형별 매매가격 조회")
    root.geometry("760x560")
    root.minsize(640, 480)

    path_var = tk.StringVar()
    limit_var = tk.StringVar(value="")
    status_var = tk.StringVar(value="엑셀 파일을 선택한 뒤 조회를 시작하세요.")
    worker: dict[str, threading.Thread | None] = {"thread": None}
    cancel_flag = {"stop": False}
    messages: queue.Queue = queue.Queue()

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)

    ttk.Label(
        frame,
        text="주소 열을 읽어 아파트·빌라 단지를 찾고, 동·호수로 그 집의 면적을 맞춘 뒤 가격을 붙입니다.",
        wraplength=700,
    ).pack(anchor="w")
    ttk.Label(
        frame,
        text="KB 시세, 아실 실거래, 국토교통부 아파트·단독/다가구 실거래(국토부인증키.txt)를 그 평형에 맞춰 사용합니다. 주소만 조회하며 성명과 주민등록번호는 전송하지 않습니다.",
        wraplength=700,
    ).pack(anchor="w", pady=(4, 12))

    file_row = ttk.Frame(frame)
    file_row.pack(fill="x")
    ttk.Entry(file_row, textvariable=path_var).pack(side="left", fill="x", expand=True)

    def choose_file() -> None:
        selected = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx"), ("모든 파일", "*.*")],
        )
        if selected:
            path_var.set(selected)

    ttk.Button(file_row, text="파일 선택", command=choose_file).pack(side="left", padx=(8, 0))

    option_row = ttk.Frame(frame)
    option_row.pack(fill="x", pady=(10, 0))
    ttk.Label(option_row, text="앞에서부터 N건만 (비우면 전체)").pack(side="left")
    ttk.Entry(option_row, width=8, textvariable=limit_var).pack(side="left", padx=(8, 0))

    button_row = ttk.Frame(frame)
    button_row.pack(fill="x", pady=(12, 8))
    start_button = ttk.Button(button_row, text="조회 시작")
    start_button.pack(side="left")
    stop_button = ttk.Button(button_row, text="중지", state="disabled")
    stop_button.pack(side="left", padx=(8, 0))

    progress = ttk.Progressbar(frame, mode="determinate")
    progress.pack(fill="x", pady=(4, 8))
    ttk.Label(frame, textvariable=status_var, wraplength=700).pack(anchor="w")

    log = tk.Text(frame, height=16, wrap="word")
    log.pack(fill="both", expand=True, pady=(8, 0))
    log.configure(state="disabled")

    def append_log(line: str) -> None:
        log.configure(state="normal")
        log.insert("end", line + "\n")
        log.see("end")
        log.configure(state="disabled")

    def set_running(running: bool) -> None:
        start_button.configure(state="disabled" if running else "normal")
        stop_button.configure(state="normal" if running else "disabled")

    def on_start() -> None:
        source = Path(path_var.get().strip())
        if not source.exists():
            messagebox.showwarning("파일 없음", "엑셀 파일을 선택하세요.")
            return
        limit = None
        if limit_var.get().strip():
            try:
                limit = int(limit_var.get().strip())
            except ValueError:
                messagebox.showwarning("입력 오류", "건수는 숫자로 입력하세요.")
                return
        cancel_flag["stop"] = False
        set_running(True)
        status_var.set("조회를 시작합니다. 같은 주소는 캐시를 사용합니다.")
        append_log(f"시작: {source.name}")

        def work() -> None:
            try:
                destination = run_file(
                    source,
                    limit,
                    lambda done, total, message: messages.put(("progress", done, total, message)),
                    cancel=lambda: cancel_flag["stop"],
                )
                messages.put(("done", str(destination)))
            except Exception as exc:
                messages.put(("error", str(exc)))

        thread = threading.Thread(target=work, daemon=True)
        worker["thread"] = thread
        thread.start()

    def on_stop() -> None:
        cancel_flag["stop"] = True
        status_var.set("현재 주소 조회가 끝나면 중지합니다.")

    start_button.configure(command=on_start)
    stop_button.configure(command=on_stop)

    def poll() -> None:
        try:
            while True:
                item = messages.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, done, total, message = item
                    progress.configure(maximum=max(total, 1), value=done)
                    status_var.set(message)
                    append_log(message)
                elif kind == "done":
                    set_running(False)
                    status_var.set(f"저장했습니다: {item[1]}")
                    append_log(f"저장: {item[1]}")
                    messagebox.showinfo("완료", f"결과 파일을 저장했습니다.\n{item[1]}")
                elif kind == "error":
                    set_running(False)
                    status_var.set("오류가 발생했습니다.")
                    append_log(item[1])
                    messagebox.showerror("오류", item[1])
        except queue.Empty:
            pass
        root.after(200, poll)

    poll()
    root.mainloop()


def self_test() -> int:
    samples = [
        (
            "부산광역시 부산광역시 부산진구 초연로11, 206동 1204호",
            {
                "sido": "부산광역시",
                "sigungu": "부산진구",
                "road": "초연로",
                "road_no": "11",
                "names": [],
                "unit_dong": "206",
                "unit_ho": "1204",
                "unit_floor": "12",
            },
        ),
        (
            "경기도 경기도 하남시 덕풍동 신장로 151번길 1305호(덕풍동 대동피렌체 아파트)",
            {
                "sido": "경기도",
                "sigungu": "하남시",
                "dong": "덕풍동",
                "names": ["대동피렌체"],
                "unit_ho": "1305",
                "unit_floor": "13",
            },
        ),
        (
            "경상남도 거제시 수양로 110 아이파크 2차 304동 1501호",
            {
                "sido": "경상남도",
                "sigungu": "거제시",
                "road": "수양로",
                "road_no": "110",
                "names": ["아이파크 2차"],
                "unit_dong": "304",
                "unit_ho": "1501",
                "unit_floor": "15",
            },
        ),
        (
            "경상북도 경상북도 포항시 북구 새천년대로1123번길 14 (창포청구타운)",
            {"sido": "경상북도", "sigungu": "포항시 북구", "road": "새천년대로1123번길", "names": ["창포청구타운"]},
        ),
        (
            "전라북도 전주시 금암2동 1597-24",
            {"sido": "전북특별자치도", "sigungu": "전주시", "dong": "금암2동", "names": []},
        ),
    ]
    failures = []
    for raw, expected in samples:
        parsed = parse_address(raw)
        for key, value in expected.items():
            actual = getattr(parsed, key)
            if actual != value:
                failures.append(f"{raw} | {key}: {actual!r} != {value!r}")
        if not search_queries(parsed) and (parsed.names or (parsed.road and parsed.road_no)):
            failures.append(f"검색어 없음: {raw}")

    assert format_manwon(11750) == "1억 1,750만원"
    assert format_manwon(20000) == "2억원"
    assert format_manwon(950) == "950만원"
    assert name_similarity("창포청구", "창포청구타운") == 1.0
    assert name_similarity("거제2차아이파크2단지", "아이파크") >= 0.72
    assert find_address_column(["성명", "주민등록상 주소", "소속"]) == 2
    assert find_address_column(["본적", "현주소", "주민등록상 주소"]) == 3
    assert find_address_column(["이름", "현주소"]) == 2
    dashed = parse_address("경상북도 포항시 북구 장성동 창포청구타운 103-902")
    if dashed.unit_dong != "103" or dashed.unit_ho != "902" or dashed.unit_floor != "9":
        failures.append(f"동호수 파싱 실패: {dashed.unit_dong} {dashed.unit_ho} {dashed.unit_floor}")
    if dashed.dong != "장성동" or "창포청구타운" not in dashed.names:
        failures.append("동-호 주소의 법정동·단지명 파싱 실패")
    bare = parse_address("전라북도 전주시 금암2동 1597-24")
    if bare.unit_dong or bare.unit_ho:
        failures.append("지번을 동호수로 잘못 읽음")
    mixed = parse_address("전라남도 광주광역시 북구 문흥동")
    if mixed.sido != "광주광역시" or mixed.sigungu != "북구" or mixed.dong != "문흥동":
        failures.append(f"광역시 중복 주소 파싱 실패: {mixed.sido} {mixed.sigungu} {mixed.dong}")
    if not _sido_matches("광주광역시", "전남광주통합특별시 광산구 월계동"):
        failures.append("통합특별시 명칭을 광주 주소와 맞추지 못함")
    if not _sido_matches("전라남도", "전남광주통합특별시 순천시 조례동"):
        failures.append("통합특별시 명칭을 전남 주소와 맞추지 못함")
    road, road_no = _extract_road("강남구 도산대로 66길 30")
    if road != "도산대로66길" or road_no != "30":
        failures.append(f"대로 번길 파싱 실패: {road} {road_no}")
    magellan = parse_address("서울특별시 강남구 테헤란로 87길 17, 마젤란 4층 404호(삼성동)")
    if magellan.names != ["마젤란"] or magellan.unit_ho != "404" or magellan.dong != "삼성동":
        failures.append(f"오피스텔 건물명 파싱 실패: {magellan.names} {magellan.dong} {magellan.unit_ho}")
    renaissance = parse_address("서울특별시 마포구 합정동 강변르네상스 가동 5층 504호")
    if renaissance.names != ["강변르네상스"] or renaissance.unit_ho != "504":
        failures.append(f"가동 건물명 파싱 실패: {renaissance.names} {renaissance.unit_ho}")
    road, road_no = _extract_road("동작구 상도로 346-2 205동 803호")
    if road != "상도로" or road_no != "346-2":
        failures.append(f"도로 부번 파싱 실패: {road} {road_no}")
    sangdo = parse_address("서울특별시 동작구 상도로 346-2 205동 803호")
    central = {
        "type": "apartment",
        "name": "힐스테이트상도센트럴파크",
        "description": "서울특별시 동작구 상도동",
        "_source": {"local1": "서울특별시", "local2": "동작구", "local3": "상도동", "신주소": "서울시 동작구 상도로 346-1"},
    }
    prestige = {
        "type": "apartment",
        "name": "힐스테이트상도프레스티지",
        "description": "서울특별시 동작구 상도동",
        "_source": {"local1": "서울특별시", "local2": "동작구", "local3": "상도동", "신주소": "서울시 동작구 상도로 346-2"},
    }
    if accept_candidate(central, sangdo):
        failures.append("부번이 다른 단지를 통과시킴")
    if not accept_candidate(prestige, sangdo):
        failures.append("같은 부번 단지를 거절함")
    suseo = parse_address("서울특별시 강남구 광평로 47길 17(수서동, 신동아아파트) 707동 707호")
    cheongdam = {
        "type": "apartment",
        "name": "신동아",
        "description": "서울특별시 강남구 청담동",
        "_source": {"local1": "서울특별시", "local2": "강남구", "local3": "청담동", "신주소": "서울시 강남구 학동로105길 30"},
    }
    suseo_item = {
        "type": "apartment",
        "name": "신동아",
        "description": "서울특별시 강남구 수서동",
        "_source": {"local1": "서울특별시", "local2": "강남구", "local3": "수서동", "신주소": "서울시 강남구 광평로47길 17"},
    }
    if candidate_score(suseo_item, suseo) <= candidate_score(cheongdam, suseo):
        failures.append("같은 이름이라도 도로명·법정동이 맞는 단지를 우선하지 않음")
    unit_only = parse_address("서울특별시 강남구 도산대로 66길 30,402호")
    if not looks_like_officetel(unit_only):
        failures.append("호실만 있는 오피스텔을 아파트로 봄")
    named_apt = parse_address("경기도 성남시 분당구 중앙공원로 17 시범한양아파트 317동 902호")
    if looks_like_officetel(named_apt):
        failures.append("아파트 동호를 오피스텔로 봄")
    if legal_dong_aliases("금암2동") != ["금암2동", "금암동"]:
        failures.append("행정동 별칭 변환 실패")
    if not _umd_in("기계면 지가리", set(legal_dong_aliases("기계면"))):
        failures.append("면·리 법정동 비교 실패")
    dong_only = parse_address("전라북도 전주시 금암2동 1597-24")
    road_apt = parse_address("경상남도 통영시 무전길 135 201동 502호")
    unit_only = parse_address("서울특별시 강남구 도산대로 66길 30,402호")
    if not _usable_unmatched_summary({"basis": "국토부 금암동 최근 실거래 4건 중앙값(대략)"}, dong_only, False):
        failures.append("법정동 대략 가격을 버림")
    if not _usable_unmatched_summary({"basis": "국토부 시군구 최근 실거래 10건 중앙값(대략, 동·단지 불명)"}, road_apt, False):
        failures.append("동호가 있는 도로명 주소의 시 단위 대략 가격을 버림")
    if _usable_unmatched_summary({"basis": "국토부 시군구 최근 실거래 10건 중앙값(대략, 동·단지 불명)", "price": 270000}, unit_only, True, ("apt", "rh")):
        failures.append("개별 호실에 구 평균가를 허용함")
    if not _usable_unmatched_summary({"basis": "국토부 시군구 최근 실거래 7건 중앙값(대략, 동·단지 불명)", "price": 24000}, unit_only, True, ("sh",)):
        failures.append("읍면이 없는 주택의 단독 대략가를 버림")
    if _usable_unmatched_summary({"basis": "국토부 시군구 최근 실거래 2건 중앙값(대략, 동·단지 불명)", "price": 960500}, unit_only, True, ("sh",)):
        failures.append("고액 건물 전체 거래를 호실 가격으로 허용함")
    if approximate_kinds(False, True, True) != [("rh",), ("sh",)]:
        failures.append("개별 호실 실거래 순서가 바뀜")
    if unit_house_price_ok(960500) or not unit_house_price_ok(5000):
        failures.append("호실에 넣을 단독·다가구 금액 상한이 잘못됨")
    if _usable_unmatched_summary({"basis": "국토부 시군구 최근 실거래 10건 중앙값(대략, 동·단지 불명)"}, parse_address("서울특별시 강남구 테헤란로 507"), False):
        failures.append("동호 없는 도로명에 구 평균가를 허용함")
    sample_deals = [
        {"name": "가나다", "umd": "금암동", "area": "84", "price": 10000, "date": "20260101"},
        {"name": "라마바", "umd": "금암동", "area": "59", "price": 20000, "date": "20260201"},
        {"name": "가나다", "umd": "다른동", "area": "84", "price": 90000, "date": "20260301"},
    ]
    summary = summarize_deals(sample_deals, "", "금암2동", "")
    if not summary or summary["price"] != 15000 or summary["count"] != 2:
        failures.append(f"동 단위 대략 가격 실패: {summary}")
    named = summarize_deals(sample_deals, "가나다아파트", "금암2동", "84")
    if not named or named["price"] != 10000:
        failures.append(f"단지 유사 대략 가격 실패: {named}")

    road_item = {
        "type": "apartment",
        "name": "연지자이2차",
        "description": "부산광역시 부산진구 연지동",
        "_source": {"local1": "부산광역시", "local2": "부산진구", "local3": "연지동", "신주소": "부산광역시 부산진구 초연로 11"},
    }
    if not accept_candidate(road_item, parse_address(samples[0][0])):
        failures.append("도로명 단지 매칭 실패")
    wrong_city = dict(road_item)
    wrong_city["_source"] = {"local1": "서울특별시", "local2": "강남구", "신주소": "서울특별시 강남구 초연로 11"}
    wrong_city["description"] = "서울특별시 강남구"
    if accept_candidate(wrong_city, parse_address(samples[0][0])):
        failures.append("다른 시군구를 통과시킴")

    if failures:
        print("\n".join(failures))
        return 1
    print("self-test ok")
    return 0


def _ensure_stdio() -> None:
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def main() -> int:
    _ensure_stdio()
    if "--self-test" in sys.argv:
        return self_test()
    if len(sys.argv) > 1:
        return run_cli(sys.argv[1:])
    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
