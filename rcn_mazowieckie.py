    """
Pobiera transakcje sprzedaży lokali mieszkalnych z RCN powiatów całej Polski
przez publiczne usługi WFS (z fallbackiem na WMS GetFeatureInfo gdy WFS nie
odpowiada), wzbogaca o ceny i powierzchnie tam gdzie się da (powiaty na
infrastrukturze epodgik mają darmowe GML transakcji), i zapisuje do pliku
.xlsx z arkuszem na powiat.

Lista powiatów ładowana z powiats.json — wygenerowana z rejestru GUGiK eziudp
(zbiór "Rejestr Cen Nieruchomości"). 370 z 384 powiatów ma WFS lub
domniemany WFS po wzorcu URL. 14 nie wystawiło żadnej usługi (w tym
Warszawa, TERYT 1465) i jest pomijanych.

11 powiatów na infrastrukturze epodgik (np. minski.rciwn.pl) udostępnia pełny
GML transakcji anonimowo — z ceną, powierzchnią, adresem. Pozostałe powiaty
dostarczają tylko metadane (data, lokalizacja, rodzaj nieruchomości) — bez ceny.

Uruchomienie:
    pip install -r requirements.txt
    python rcn_mazowieckie.py --out rcn_pl.xlsx
    python rcn_mazowieckie.py --woj Mazowieckie --out rcn_maz.xlsx
    python rcn_mazowieckie.py --only 1412,1418
    python rcn_mazowieckie.py --skip-details    # pomiń wzbogacanie cenami
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Optional
from xml.etree import ElementTree as ET

import pandas as pd
import requests
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

# ---------------------------------------------------------------------------
# Konfiguracja
# ---------------------------------------------------------------------------

CACHE_DIR = Path(__file__).resolve().parent / ".rcn_cache"
CACHE_DIR.mkdir(exist_ok=True)

DETAIL_WORKERS = 16
WFS_TIMEOUT = 60
DETAIL_TIMEOUT = 20
RETRY_ATTEMPTS = 3

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("rcn")

# Mapowanie TERYT[0:2] -> nazwa województwa
WOJEWODZTWA = {
    "02": "Dolnośląskie", "04": "Kujawsko-Pomorskie", "06": "Lubelskie",
    "08": "Lubuskie", "10": "Łódzkie", "12": "Małopolskie",
    "14": "Mazowieckie", "16": "Opolskie", "18": "Podkarpackie",
    "20": "Podlaskie", "22": "Pomorskie", "24": "Śląskie",
    "26": "Świętokrzyskie", "28": "Warmińsko-Mazurskie",
    "30": "Wielkopolskie", "32": "Zachodniopomorskie",
}


def load_powiats() -> list[dict]:
    """Wczytuje listę powiatów z powiats.json (wygenerowanego z eziudp)."""
    path = Path(__file__).resolve().parent / "powiats.json"
    if not path.exists():
        raise FileNotFoundError(
            f"Brak {path}. Powinien być wygenerowany z rejestru eziudp."
        )
    data = json.loads(path.read_text(encoding="utf-8"))
    for p in data:
        p["woj"] = WOJEWODZTWA.get(p["teryt"][:2], "?")
    return data


POWIATS: list[dict] = load_powiats()

# Kandydaci na nazwę typu obiektu w WFS (różne namespace).
TYPE_NAME_CANDIDATES = ["ms:transakcje", "ewns:transakcje", "transakcje"]

# Mapowanie WFS RODZAJ_NIERUCHOMOSCI (tekstowo) na nasz typ.
KIND_MAP = {
    "nieruchomoscLokalowa": "lokal",
    "nieruchomoscBudynkowa": "budynek",
    "nieruchomoscGruntowaNiezabudowana": "działka niezabudowana",
    "nieruchomoscGruntowaZabudowana": "działka zabudowana",
    "nieruchomoscGruntowa": "działka",
}

# rodzajTransakcji w WFS: "wolnyRynek" to typowa sprzedaż.
SALE_KIND_TEXT = "wolnyRynek"

# Funkcja lokalu w GML: 1 = mieszkalna. Inne kody to garaże/postojowe/użytkowe.
APARTMENT_FUNCTION_CODE = "1"

# Kody przeznaczenia w MPZP/sposobu użytkowania — interpretacja przybliżona
# (pełna tabela jest w słowniku RCN_1.4.XSD; tu mapujemy najpopularniejsze).
SPOSOB_UZYTKOWANIA = {
    "1": "rolne", "2": "leśne", "3": "budowlane", "4": "inwestycyjne",
    "5": "mieszkaniowe", "6": "usługowe", "7": "przemysłowe",
}


# ---------------------------------------------------------------------------
# Model danych
# ---------------------------------------------------------------------------


@dataclass
class Transaction:
    teryt: str = ""
    powiat: str = ""
    wojewodztwo: str = ""
    lokalny_id: str = ""
    data_transakcji: Optional[date] = None
    rodzaj_transakcji: str = ""
    rodzaj_nieruchomosci: str = ""
    x_2180: Optional[float] = None
    y_2180: Optional[float] = None
    detail_link: str = ""

    typ_nieruchomosci: str = ""  # "lokal" / "działka niezabudowana" / "budynek" / ...

    # Pola wzbogacone z detail GML (puste dla powiatów bez detail-fetcha).
    cena_brutto: Optional[float] = None
    powierzchnia_m2: Optional[float] = None
    cena_za_m2: Optional[float] = None

    # Zawartość transakcji — co dokładnie zostało sprzedane.
    zawartosc: str = ""              # np. "lokal mieszkalny 52,93 m²; działka 985 m²"
    liczba_dzialek: int = 0
    liczba_lokali: int = 0
    liczba_budynkow: int = 0
    id_obiektow: str = ""            # identyfikatory katastralne, oddzielone "; "
    pow_dzialek_m2: Optional[float] = None   # suma powierzchni działek
    pow_lokali_m2: Optional[float] = None    # suma powierzchni użytkowej lokali
    pow_budynkow_m2: Optional[float] = None  # suma powierzchni zabudowy budynków
    przeznaczenie_mpzp: str = ""     # np. "MN", "WZ", "U"
    sposob_uzytkowania: str = ""

    # Lokal-specific
    funkcja_lokalu: str = ""
    nr_kondygnacji: str = ""

    # Adres (z pierwszego adresu znalezionego w GML)
    miejscowosc: str = ""
    ulica: str = ""
    numer: str = ""

    rodzaj_rynku: str = ""  # 1 = pierwotny, 2 = wtórny
    data_aktu: Optional[date] = None
    notariusz: str = ""
    oznaczenie_dokumentu: str = ""
    oznaczenie_transakcji: str = ""
    udzial_w_prawie: str = ""

    # Pełny opis tekstowy z GML (RCN_Nieruchomosc.opis), wieloliniowy
    opis: str = ""

    detail_status: str = "skipped"  # skipped | ok | failed | no_link

    def to_row(self) -> dict:
        d = asdict(self)
        d["data_transakcji"] = self.data_transakcji.isoformat() if self.data_transakcji else None
        d["data_aktu"] = self.data_aktu.isoformat() if self.data_aktu else None
        return d


# ---------------------------------------------------------------------------
# WFS — pobieranie metadanych
# ---------------------------------------------------------------------------

WFS_NS = {
    "wfs": "http://www.opengis.net/wfs/2.0",
    "gml": "http://www.opengis.net/gml/3.2",
    "ms": "http://mapserver.gis.umn.edu/mapserver",
    "ewns": "http://xsd.geoportal2.pl/ewns",
}


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


@retry(
    reraise=True,
    stop=stop_after_attempt(RETRY_ATTEMPTS),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((requests.RequestException,)),
)
def _wfs_request(session: requests.Session, url: str, params: dict) -> bytes:
    r = session.get(url, params=params, timeout=WFS_TIMEOUT)
    r.raise_for_status()
    return r.content


def _detect_type_name(session: requests.Session, wfs_url: str) -> Optional[str]:
    """Pobiera GetCapabilities i znajduje pierwszą pasującą nazwę FeatureType."""
    try:
        caps = _wfs_request(session, wfs_url, {"service": "WFS", "request": "GetCapabilities"})
    except Exception as exc:
        log.warning("GetCapabilities nieudany dla %s: %s", wfs_url, exc)
        return None
    try:
        root = ET.fromstring(caps)
    except ET.ParseError as exc:
        log.warning("Niepoprawny XML GetCapabilities dla %s: %s", wfs_url, exc)
        return None

    for el in root.iter():
        if _strip_ns(el.tag) == "Name" and el.text and "transakcje" in el.text.lower():
            return el.text.strip()
    return None


def _parse_pos(text: str) -> tuple[Optional[float], Optional[float]]:
    parts = re.split(r"\s+", text.strip())
    if len(parts) < 2:
        return None, None
    try:
        return float(parts[0]), float(parts[1])
    except ValueError:
        return None, None


def _parse_wfs_member(member_el: ET.Element, powiat: dict) -> Optional[Transaction]:
    # Wnętrze <wfs:member>: jeden element <{ns}transakcje gml:id="...">
    feat = None
    for child in member_el:
        if _strip_ns(child.tag) == "transakcje":
            feat = child
            break
    if feat is None:
        return None

    tx = Transaction(
        teryt=powiat["teryt"],
        powiat=powiat.get("short", powiat.get("organ", "?")),
        wojewodztwo=powiat.get("woj", ""),
    )

    for el in feat:
        tag = _strip_ns(el.tag)
        text = (el.text or "").strip()
        if tag == "LOKALNY_ID":
            tx.lokalny_id = text
        elif tag == "DATA_TRANSAKCJI":
            try:
                tx.data_transakcji = datetime.strptime(text, "%Y-%m-%d").date()
            except ValueError:
                pass
        elif tag == "RODZAJ_TRANSAKCJI":
            tx.rodzaj_transakcji = text
        elif tag == "RODZAJ_NIERUCHOMOSCI":
            tx.rodzaj_nieruchomosci = text
            tx.typ_nieruchomosci = KIND_MAP.get(text, text)
        elif tag == "LINK":
            tx.detail_link = text
        elif tag in ("msGeometry", "geometria"):
            for pt in el.iter():
                if _strip_ns(pt.tag) == "pos":
                    x, y = _parse_pos(pt.text or "")
                    # geoportal2 zwraca często w EPSG:2178 (Y X), epodgik EPSG:2180.
                    # Heurystyka: jeśli pierwszy "X" jest > 1_000_000 to to nie 2180,
                    # to 2178/2179 i osie są (north, east) — odwracamy.
                    if x is not None and y is not None:
                        if x > 1_000_000 or y > 1_000_000:
                            tx.x_2180 = y  # raw X jest north, więc Y po swapie
                            tx.y_2180 = x
                        else:
                            tx.x_2180 = x
                            tx.y_2180 = y
                    break

    return tx


def _powiat_label(p: dict) -> str:
    return f"{p['teryt']} {p.get('short', p.get('organ', '?'))}"


def fetch_powiat_wfs(session: requests.Session, powiat: dict) -> list[Transaction]:
    """Próbuje WFS, zwraca pustą listę gdy brak WFS URL lub błąd."""
    if not powiat.get("wfs"):
        return []

    cache_path = CACHE_DIR / f"wfs_{powiat['teryt']}.xml"
    label = _powiat_label(powiat)
    if cache_path.exists() and cache_path.stat().st_size > 0:
        data = cache_path.read_bytes()
        log.info("[%s] cache WFS (%d B)", label, len(data))
    else:
        type_name = _detect_type_name(session, powiat["wfs"]) or "ms:transakcje"
        log.info("[%s] WFS GetFeature typeNames=%s", label, type_name)
        try:
            data = _wfs_request(
                session,
                powiat["wfs"],
                {
                    "service": "WFS",
                    "version": "2.0.0",
                    "request": "GetFeature",
                    "typeNames": type_name,
                },
            )
        except Exception as exc:
            log.warning("[%s] WFS NIEUDANY: %s", label, exc)
            return []
        cache_path.write_bytes(data)

    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        log.error("[%s] niepoprawny XML: %s", label, exc)
        return []

    out: list[Transaction] = []
    for el in root.iter():
        if _strip_ns(el.tag) == "member":
            tx = _parse_wfs_member(el, powiat)
            if tx:
                out.append(tx)
    log.info("[%s] WFS: %d transakcji", label, len(out))
    return out


# ---------------------------------------------------------------------------
# WMS GetFeatureInfo — fallback gdy WFS nie odpowiada
# ---------------------------------------------------------------------------

# WMS scale denominator transakcji to zwykle 50000. Dla 1024×1024 px tile:
# scale = bbox_width / (1024 * 0.00028) < 50000  →  bbox_width < 14336 m.
# Bezpiecznie biorę 10000 m × 10000 m kafelki.
WMS_TILE_SIZE_M = 10000
WMS_TILE_PX = 1024
WMS_FEATURE_COUNT = 5000

# Heurystyka zasięgu Polski w EPSG:2180.
POLAND_BBOX_2180 = (140000, 130000, 870000, 870000)


def _bbox_for_powiat(p: dict) -> tuple[float, float, float, float]:
    """
    Bez dostępu do bboxu powiatu w EPSG:2180 zwracamy bbox Polski.
    Skutkuje to znacznie większą liczbą kafelków (większość pusta),
    ale jest jedynym uniwersalnym rozwiązaniem dla WMS fallback.
    Powiaty z WFS i tak go nie używają.
    """
    return POLAND_BBOX_2180


@retry(
    reraise=True,
    stop=stop_after_attempt(RETRY_ATTEMPTS),
    wait=wait_exponential(multiplier=1, min=1, max=10),
    retry=retry_if_exception_type((requests.RequestException,)),
)
def _wms_gfi(session: requests.Session, url: str, bbox: tuple[float, float, float, float]) -> bytes:
    params = {
        "SERVICE": "WMS", "VERSION": "1.1.1", "REQUEST": "GetFeatureInfo",
        "LAYERS": "transakcje", "QUERY_LAYERS": "transakcje",
        "SRS": "EPSG:2180",
        "BBOX": ",".join(f"{v:.2f}" for v in bbox),
        "WIDTH": str(WMS_TILE_PX), "HEIGHT": str(WMS_TILE_PX),
        "X": str(WMS_TILE_PX // 2), "Y": str(WMS_TILE_PX // 2),
        "INFO_FORMAT": "application/vnd.ogc.gml",
        "FEATURE_COUNT": str(WMS_FEATURE_COUNT),
    }
    r = session.get(url, params=params, timeout=WFS_TIMEOUT)
    r.raise_for_status()
    return r.content


def fetch_powiat_wms(session: requests.Session, powiat: dict) -> list[Transaction]:
    """Fallback: tile-based WMS GetFeatureInfo. Wolne i niepełne, ale działa."""
    wms_url = powiat.get("wms")
    if not wms_url:
        return []

    label = _powiat_label(powiat)
    log.info("[%s] FALLBACK WMS GetFeatureInfo (kafelki)", label)

    xmin, ymin, xmax, ymax = _bbox_for_powiat(powiat)
    seen: set[str] = set()
    out: list[Transaction] = []
    tiles = []
    x = xmin
    while x < xmax:
        y = ymin
        while y < ymax:
            tiles.append((x, y, min(x + WMS_TILE_SIZE_M, xmax), min(y + WMS_TILE_SIZE_M, ymax)))
            y += WMS_TILE_SIZE_M
        x += WMS_TILE_SIZE_M

    for idx, tile in enumerate(tiles, start=1):
        try:
            data = _wms_gfi(session, wms_url, tile)
        except Exception:
            continue
        if not data or len(data) < 200 or data.startswith(b"<!DOCTYPE") or data.startswith(b"<html"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        for el in root.iter():
            if _strip_ns(el.tag) == "transakcje":
                # Wrap in synthetic member for reuse of parser.
                synthetic = ET.Element("{http://www.opengis.net/wfs/2.0}member")
                synthetic.append(el)
                tx = _parse_wfs_member(synthetic, powiat)
                if not tx:
                    continue
                key = tx.lokalny_id or f"{tx.x_2180}:{tx.y_2180}:{tx.data_transakcji}"
                if key in seen:
                    continue
                seen.add(key)
                out.append(tx)
        if idx % 50 == 0:
            log.info("  [%s] kafelek %d/%d, zebrano %d", label, idx, len(tiles), len(out))

    log.info("[%s] WMS fallback: %d transakcji", label, len(out))
    return out


def fetch_powiat(session: requests.Session, powiat: dict) -> list[Transaction]:
    """Główny fetcher: WFS preferowany, WMS jako fallback."""
    txs = fetch_powiat_wfs(session, powiat)
    if txs:
        return txs
    if powiat.get("wms"):
        return fetch_powiat_wms(session, powiat)
    return []


# ---------------------------------------------------------------------------
# Detail GML — wzbogacanie dla epodgik
# ---------------------------------------------------------------------------

# Lokalny serwis powiatu po dotarciu z LINK (np. https://minski.rciwn.pl/?identyfikatory=...)
# zawiera ukryte linki do getGml.php?p/=... które wydają pełne dane transakcji.
_GETGML_RE = re.compile(r"getGml\.php\?[^\"']+")


@retry(
    reraise=True,
    stop=stop_after_attempt(RETRY_ATTEMPTS),
    wait=wait_exponential(multiplier=1, min=1, max=5),
    retry=retry_if_exception_type((requests.RequestException,)),
)
def _http_get_text(session: requests.Session, url: str) -> str:
    r = session.get(url, timeout=DETAIL_TIMEOUT)
    r.raise_for_status()
    r.encoding = r.encoding or "utf-8"
    return r.text


def _to_f(text: str) -> Optional[float]:
    try:
        return float(text.replace(",", "."))
    except (ValueError, AttributeError):
        return None


def _parse_detail_gml(xml_text: str, tx: Transaction) -> None:
    """
    Parsuje GML transakcji wg schematu RCN_1.4.XSD.
    Struktura:
      RCN_Transakcja (1 sztuka): cenaTransakcjiBrutto, rodzajTransakcji, ...
      RCN_Dokument   (1 sztuka): notariusz, data aktu, oznaczenie
      RCN_Nieruchomosc (1+): cenaNieruchomosciBrutto, pole, opis, → linki
      RCN_Lokal / RCN_Dzialka / RCN_Budynek (0+): szczegóły obiektów
      RCN_Adres (0+): adres
    """
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return

    pow_dzialek: list[float] = []
    pow_lokali: list[float] = []
    pow_budynkow: list[float] = []
    items: list[str] = []
    obj_ids: list[str] = []
    przeznaczenia: list[str] = []
    sposoby_uz: list[str] = []
    opisy: list[str] = []

    for member in root.iter():
        if _strip_ns(member.tag) != "featureMember":
            continue
        # featureMember zawiera jedno z RCN_Transakcja/Dokument/Nieruchomosc/Lokal/Dzialka/Budynek/Adres
        if len(member) == 0:
            continue
        item = member[0]
        item_kind = _strip_ns(item.tag)

        if item_kind == "RCN_Transakcja":
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag == "cenaTransakcjiBrutto" and tx.cena_brutto is None:
                    tx.cena_brutto = _to_f(txt)
                elif tag == "rodzajRynku":
                    tx.rodzaj_rynku = txt
                elif tag == "oznaczenieTransakcji":
                    tx.oznaczenie_transakcji = txt

        elif item_kind == "RCN_Dokument":
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag == "oznaczenieDokumentu":
                    tx.oznaczenie_dokumentu = txt
                elif tag == "dataSporzadzeniaDokumentu":
                    try:
                        tx.data_aktu = datetime.strptime(txt, "%Y-%m-%d").date()
                    except ValueError:
                        pass
                elif tag == "tworcaDokumentu":
                    tx.notariusz = txt

        elif item_kind == "RCN_Nieruchomosc":
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag == "udzialWPrawieDoNieruchomosci" and not tx.udzial_w_prawie:
                    tx.udzial_w_prawie = txt
                elif tag == "opis":
                    opisy.append(txt)

        elif item_kind == "RCN_Lokal":
            tx.liczba_lokali += 1
            local_area = None
            local_id = None
            local_fn = None
            local_floor = None
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag == "powUzytkowaLokalu":
                    local_area = _to_f(txt)
                elif tag == "idLokalu":
                    local_id = txt
                elif tag == "funkcjaLokalu":
                    local_fn = txt
                elif tag == "nrKondygnacji":
                    local_floor = txt
            if local_area:
                pow_lokali.append(local_area)
            if local_id:
                obj_ids.append(local_id)
            if not tx.funkcja_lokalu and local_fn:
                tx.funkcja_lokalu = local_fn
            if not tx.nr_kondygnacji and local_floor:
                tx.nr_kondygnacji = local_floor
            desc = "lokal"
            if local_fn == "1":
                desc = "lokal mieszkalny"
            elif local_fn:
                desc = f"lokal (funkcja {local_fn})"
            if local_area:
                desc += f" {local_area:.2f} m²"
            items.append(desc)

        elif item_kind == "RCN_Dzialka":
            tx.liczba_dzialek += 1
            d_area = None
            d_id = None
            d_mpzp = None
            d_sp = None
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag == "polePowierzchniEwidencyjnej":
                    d_area = _to_f(txt)
                elif tag == "idDzialki":
                    d_id = txt
                elif tag == "przeznaczenieWMPZP":
                    d_mpzp = txt
                elif tag == "sposobUzytkowania":
                    d_sp = txt
            if d_area:
                pow_dzialek.append(d_area)
            if d_id:
                obj_ids.append(d_id)
            if d_mpzp:
                przeznaczenia.append(d_mpzp)
            if d_sp:
                sposoby_uz.append(SPOSOB_UZYTKOWANIA.get(d_sp, d_sp))
            desc = "działka"
            if d_area:
                desc += f" {d_area:.0f} m²"
            if d_mpzp:
                desc += f" ({d_mpzp})"
            items.append(desc)

        elif item_kind == "RCN_Budynek":
            tx.liczba_budynkow += 1
            b_area = None
            b_id = None
            for el in item.iter():
                tag = _strip_ns(el.tag)
                txt = (el.text or "").strip() if el.text else ""
                if not txt:
                    continue
                if tag in ("powZabudowyBudynku", "polePowZabudowyBudynku", "powUzytkowaBudynku"):
                    b_area = _to_f(txt)
                elif tag == "idBudynku":
                    b_id = txt
            if b_area:
                pow_budynkow.append(b_area)
            if b_id:
                obj_ids.append(b_id)
            desc = "budynek"
            if b_area:
                desc += f" {b_area:.0f} m²"
            items.append(desc)

        elif item_kind == "RCN_Adres":
            # Bierzemy pierwszy znaleziony adres jako reprezentatywny.
            if not tx.miejscowosc and not tx.ulica:
                for el in item.iter():
                    tag = _strip_ns(el.tag)
                    txt = (el.text or "").strip() if el.text else ""
                    if not txt:
                        continue
                    if tag == "miejscowosc":
                        tx.miejscowosc = txt
                    elif tag == "ulica":
                        tx.ulica = txt
                    elif tag == "numerPorzadkowy":
                        tx.numer = txt

    # Agregacja
    if pow_dzialek:
        tx.pow_dzialek_m2 = round(sum(pow_dzialek), 2)
    if pow_lokali:
        tx.pow_lokali_m2 = round(sum(pow_lokali), 2)
    if pow_budynkow:
        tx.pow_budynkow_m2 = round(sum(pow_budynkow), 2)

    # Wspólne pole "powierzchnia_m2" — wybieramy najbardziej charakterystyczne:
    # dla lokalu: lokal, dla działki: działka, dla budynku: budynek.
    if tx.pow_lokali_m2:
        tx.powierzchnia_m2 = tx.pow_lokali_m2
    elif tx.pow_dzialek_m2:
        tx.powierzchnia_m2 = tx.pow_dzialek_m2
    elif tx.pow_budynkow_m2:
        tx.powierzchnia_m2 = tx.pow_budynkow_m2

    if items:
        tx.zawartosc = "; ".join(items)
    if obj_ids:
        tx.id_obiektow = "; ".join(obj_ids)
    if przeznaczenia:
        tx.przeznaczenie_mpzp = ", ".join(sorted(set(przeznaczenia)))
    if sposoby_uz:
        tx.sposob_uzytkowania = ", ".join(sorted(set(sposoby_uz)))
    if opisy:
        tx.opis = " | ".join(opisy)

    if tx.cena_brutto and tx.powierzchnia_m2 and tx.powierzchnia_m2 > 0:
        tx.cena_za_m2 = round(tx.cena_brutto / tx.powierzchnia_m2, 2)


def enrich_epodgik_detail(
    session: requests.Session,
    tx: Transaction,
    save_gml: bool = False,
) -> None:
    if not tx.detail_link:
        tx.detail_status = "no_link"
        return

    # Cache HTML/GML lokalnie żeby nie ddosować serwisów przy ponownym uruchomieniu.
    gml_path = CACHE_DIR / "gml" / tx.teryt / f"{tx.lokalny_id}.gml"
    gml_xml: Optional[str] = None
    if gml_path.exists() and gml_path.stat().st_size > 0:
        try:
            gml_xml = gml_path.read_text(encoding="utf-8")
        except Exception:
            gml_xml = None

    if gml_xml is None:
        try:
            html = _http_get_text(session, tx.detail_link)
        except Exception as exc:
            log.debug("detail HTML failed %s: %s", tx.detail_link, exc)
            tx.detail_status = "failed"
            return

        m = _GETGML_RE.search(html)
        if not m:
            tx.detail_status = "no_link"
            return

        host_m = re.match(r"(https?://[^/]+)", tx.detail_link)
        gml_url = (host_m.group(1) if host_m else "") + "/" + m.group(0)

        try:
            gml_xml = _http_get_text(session, gml_url)
        except Exception as exc:
            log.debug("detail GML failed %s: %s", gml_url, exc)
            tx.detail_status = "failed"
            return

        if save_gml:
            gml_path.parent.mkdir(parents=True, exist_ok=True)
            gml_path.write_text(gml_xml, encoding="utf-8")

    _parse_detail_gml(gml_xml, tx)
    tx.detail_status = "ok"


# ---------------------------------------------------------------------------
# Filtry
# ---------------------------------------------------------------------------


def is_relevant(tx: Transaction, types_allowed: set[str], sales_only: bool) -> bool:
    """
    types_allowed — zbiór z {'lokal', 'działka', 'budynek'}. Puste = wszystko.
    sales_only — True ogranicza do wolnyRynek (typowa sprzedaż).
    """
    if types_allowed:
        t = (tx.typ_nieruchomosci or "").lower()
        # Dopasuj heurystycznie po fragmencie ("działka niezabudowana" zawiera "działka").
        if not any(allowed in t for allowed in types_allowed):
            return False
    if sales_only and tx.rodzaj_transakcji:
        if SALE_KIND_TEXT.lower() not in tx.rodzaj_transakcji.lower():
            return False
    return True


def in_last_n_years(d: Optional[date], years: int, today: date) -> bool:
    if d is None:
        return False
    try:
        cutoff = date(today.year - years, today.month, today.day)
    except ValueError:
        cutoff = date(today.year - years, today.month, 28)
    return d >= cutoff


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

SAFE_SHEET_RE = re.compile(r"[\\/*?:\[\]]")


def _safe_sheet(name: str) -> str:
    return SAFE_SHEET_RE.sub("_", name)[:31]


def write_excel(transactions: list[Transaction], out_path: str) -> None:
    rows = [t.to_row() for t in transactions]
    df = pd.DataFrame(rows)

    if df.empty:
        log.warning("Brak transakcji do zapisu — nie tworzę pliku.")
        return

    column_order = [
        "teryt", "wojewodztwo", "powiat",
        "data_transakcji", "typ_nieruchomosci",
        "cena_brutto", "powierzchnia_m2", "cena_za_m2",
        "zawartosc",
        "liczba_dzialek", "liczba_lokali", "liczba_budynkow",
        "pow_dzialek_m2", "pow_lokali_m2", "pow_budynkow_m2",
        "miejscowosc", "ulica", "numer",
        "przeznaczenie_mpzp", "sposob_uzytkowania",
        "funkcja_lokalu", "nr_kondygnacji",
        "rodzaj_transakcji", "rodzaj_nieruchomosci", "rodzaj_rynku",
        "udzial_w_prawie",
        "opis",
        "id_obiektow",
        "oznaczenie_transakcji", "oznaczenie_dokumentu",
        "data_aktu", "notariusz",
        "x_2180", "y_2180",
        "lokalny_id", "detail_link", "detail_status",
    ]
    df = df.reindex(columns=column_order)
    df = df.sort_values(
        by=["wojewodztwo", "powiat", "data_transakcji"],
        ascending=[True, True, False],
        na_position="last",
    )

    log.info("Zapisuję %d transakcji do %s (%d powiatów)",
             len(df), out_path, df["teryt"].nunique())
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # _Podsumowanie per powiat
        summary = (
            df.groupby(["wojewodztwo", "teryt", "powiat"], dropna=False)
            .agg(
                liczba_transakcji=("lokalny_id", "count"),
                z_ceną=("cena_brutto", lambda s: s.notna().sum()),
                średnia_cena=("cena_brutto", "mean"),
                średnia_cena_m2=("cena_za_m2", "mean"),
                mediana_cena_m2=("cena_za_m2", "median"),
                min_cena_m2=("cena_za_m2", "min"),
                max_cena_m2=("cena_za_m2", "max"),
            )
            .round(2)
            .reset_index()
            .sort_values(["wojewodztwo", "liczba_transakcji"], ascending=[True, False])
        )
        summary.to_excel(writer, sheet_name="_Podsumowanie", index=False)

        # _Podsumowanie per województwo
        woj_summary = (
            df.groupby("wojewodztwo", dropna=False)
            .agg(
                liczba_transakcji=("lokalny_id", "count"),
                z_ceną=("cena_brutto", lambda s: s.notna().sum()),
                średnia_cena_m2=("cena_za_m2", "mean"),
                mediana_cena_m2=("cena_za_m2", "median"),
            )
            .round(2)
            .reset_index()
            .sort_values("liczba_transakcji", ascending=False)
        )
        woj_summary.to_excel(writer, sheet_name="_Województwa", index=False)

        # Arkusze per powiat — nazwa: "TERYT short" (unikalne).
        used_names: set[str] = set()
        for (teryt, powiat_name), sub in df.groupby(["teryt", "powiat"], dropna=False):
            base = _safe_sheet(f"{teryt} {powiat_name}")
            name = base
            i = 1
            while name in used_names:
                name = _safe_sheet(f"{base}_{i}")
                i += 1
            used_names.add(name)
            sub.to_excel(writer, sheet_name=name, index=False)

    log.info("Gotowe: %s", out_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default="rcn_polska.xlsx")
    parser.add_argument("--years", type=int, default=3)
    parser.add_argument("--only", default="", help="Lista TERYT po przecinku, np. 1412,1418")
    parser.add_argument(
        "--woj",
        default="",
        help="Filtr województwa: nazwa lub kod 2-cyfrowy (Mazowieckie, 14). "
             "Dla wielu rozdziel przecinkiem.",
    )
    parser.add_argument("--skip-details", action="store_true", help="Nie pobieraj GML szczegółów")
    parser.add_argument(
        "--types",
        default="lokal,działka,budynek",
        help="Typy nieruchomości po przecinku: lokal, działka, budynek. "
             "Pusty string = brak filtra (wszystko).",
    )
    parser.add_argument(
        "--mieszkalne-tylko",
        action="store_true",
        help="Dla lokali zostaw tylko mieszkalne (funkcjaLokalu=1). "
             "Wymaga --include-details (działa po wzbogaceniu).",
    )
    parser.add_argument(
        "--wszystkie-transakcje",
        action="store_true",
        help="Nie filtruj rodzaju transakcji (domyślnie tylko wolnyRynek).",
    )
    parser.add_argument(
        "--save-gml",
        action="store_true",
        help="Zapisuj surowy GML każdej transakcji do .rcn_cache/gml/{teryt}/{id}.gml",
    )
    parser.add_argument("--workers", type=int, default=DETAIL_WORKERS)
    parser.add_argument(
        "--clear-cache",
        action="store_true",
        help="Wyczyść .rcn_cache przed startem",
    )
    args = parser.parse_args()

    if args.clear_cache:
        for f in CACHE_DIR.glob("*"):
            f.unlink()
        log.info("Cache wyczyszczony.")

    today = date.today()
    log.info("Cutoff dat: od %s", date(today.year - args.years, today.month, today.day))

    powiats_to_fetch = POWIATS
    if args.woj:
        wanted_woj: set[str] = set()
        for w in args.woj.split(","):
            w = w.strip()
            if not w:
                continue
            if w.isdigit() and w in WOJEWODZTWA:
                wanted_woj.add(WOJEWODZTWA[w])
            else:
                wanted_woj.add(w)
        powiats_to_fetch = [p for p in powiats_to_fetch if p.get("woj") in wanted_woj]
        log.info("Filtr województw: %s → %d powiatów", wanted_woj, len(powiats_to_fetch))
    if args.only:
        wanted = {x.strip() for x in args.only.split(",") if x.strip()}
        powiats_to_fetch = [p for p in powiats_to_fetch if p["teryt"] in wanted]
        log.info("Filtr TERYT: %s → %d powiatów", wanted, len(powiats_to_fetch))

    if not powiats_to_fetch:
        log.error("Filtry nie wybrały żadnego powiatu.")
        return 1

    session = requests.Session()
    session.headers.update({"User-Agent": "rcn-fetcher/2.0"})

    # 1) WFS / WMS fallback — metadane
    all_tx: list[Transaction] = []
    for idx, p in enumerate(powiats_to_fetch, start=1):
        log.info("=== %d/%d %s [%s] ===",
                 idx, len(powiats_to_fetch), _powiat_label(p), p.get("woj", "?"))
        try:
            all_tx.extend(fetch_powiat(session, p))
        except Exception as exc:
            log.error("[%s] błąd: %s", _powiat_label(p), exc)

    log.info("Łącznie z WFS/WMS: %d transakcji (wszystkie kategorie)", len(all_tx))

    # 2) Filtruj
    types_allowed = {t.strip().lower() for t in args.types.split(",") if t.strip()}
    log.info("Typy nieruchomości: %s | wolnyRynek-only: %s",
             types_allowed or "wszystkie", not args.wszystkie_transakcje)

    filtered = [
        t for t in all_tx
        if is_relevant(t, types_allowed, not args.wszystkie_transakcje)
        and in_last_n_years(t.data_transakcji, args.years, today)
    ]
    log.info("Po filtrach (typ + transakcja + %d lat): %d", args.years, len(filtered))

    # 3) Wzbogać szczegółami (tylko epodgik, gdzie LINK działa)
    if not args.skip_details:
        detail_capable_teryt = {p["teryt"] for p in POWIATS if p.get("detail_capable")}
        enrichable = [t for t in filtered if t.teryt in detail_capable_teryt]
        log.info(
            "Wzbogacam ceny dla %d transakcji (powiaty epodgik), workers=%d",
            len(enrichable), args.workers,
        )
        done = 0
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = {
                pool.submit(enrich_epodgik_detail, session, t, args.save_gml): t
                for t in enrichable
            }
            for fut in as_completed(futures):
                done += 1
                if done % 100 == 0:
                    ok = sum(1 for t in enrichable if t.detail_status == "ok")
                    log.info("  postęp: %d/%d, z ceną: %d", done, len(enrichable), ok)
        ok = sum(1 for t in enrichable if t.detail_status == "ok")
        log.info("Wzbogacanie zakończone: %d/%d z ceną", ok, len(enrichable))

        # Opcjonalnie zostaw tylko lokale mieszkalne (funkcjaLokalu==1).
        # Działkom/budynkom (które nie mają tego pola) nie ruszamy.
        if args.mieszkalne_tylko:
            before = len(filtered)
            filtered = [
                t for t in filtered
                if "lokal" not in (t.typ_nieruchomosci or "").lower()
                or (not t.funkcja_lokalu) or t.funkcja_lokalu == APARTMENT_FUNCTION_CODE
            ]
            if before != len(filtered):
                log.info("Odfiltrowano %d niemieszkalnych lokali (garaże, postojowe).",
                         before - len(filtered))

    # 4) Excel
    write_excel(filtered, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
