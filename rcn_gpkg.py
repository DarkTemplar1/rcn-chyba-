"""
Wczytuje plik(i) GeoPackage z RCN (Rejestr Cen Nieruchomości) w standardowym
schemacie GUGiK (3 tabele: transakcje_lokale, transakcje_dzialki,
transakcje_budynki) i zapisuje wszystkie transakcje do pojedynczego pliku
.xlsx z arkuszami pogrupowanymi po dzielnicy (dla Warszawy) lub gminie.

Uruchomienie:
    pip install -r requirements.txt

    # jeden plik (auto-detekcja: jeśli to Warszawa → arkusz na dzielnicę)
    python rcn_gpkg.py "C:/Users/skalk/Downloads/1465_transakcje_ceny.gpkg/1465_transakcje_ceny.gpkg"

    # wiele plików łączone w jeden Excel
    python rcn_gpkg.py 1465_*.gpkg 1418_*.gpkg --out wszystko.xlsx

    # tylko mieszkania, ostatnie 3 lata, sprzedaż na wolnym rynku
    python rcn_gpkg.py plik.gpkg --types lokal --years 3 --mieszkalne-tylko

    # bez filtrów (cała historia)
    python rcn_gpkg.py plik.gpkg --years 0 --wszystkie-transakcje
"""

from __future__ import annotations

import argparse
import logging
import re
import sqlite3
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd

# ---------------------------------------------------------------------------
# Konfiguracja
# ---------------------------------------------------------------------------

# TERYT pierwszych 6 znaków identyfikatora obiektu → dzielnica Warszawy.
# Ustalone empirycznie z bazy 1465_transakcje_ceny.gpkg (porównaniem adresów).
WARSAW_DISTRICTS = {
    "146502": "Bemowo",
    "146503": "Białołęka",
    "146504": "Bielany",
    "146505": "Mokotów",
    "146506": "Ochota",
    "146507": "Praga-Południe",
    "146508": "Praga-Północ",
    "146509": "Rembertów",
    "146510": "Śródmieście",
    "146511": "Targówek",
    "146512": "Ursus",
    "146513": "Ursynów",
    "146514": "Wawer",
    "146515": "Wesoła",
    "146516": "Wilanów",
    "146517": "Włochy",
    "146518": "Wola",
    "146519": "Żoliborz",
}

# Mapowanie TERYT[0:4] (powiat) → nazwa, dla pliku wielu powiatów.
# Można rozszerzyć — tu tylko niezbędne minimum + Warszawa.
KNOWN_POWIATS = {
    "1465": "M. st. Warszawa",
}

# Mapowanie pierwszych 2 cyfr TERYT → województwo.
WOJEWODZTWA = {
    "02": "Dolnośląskie", "04": "Kujawsko-Pomorskie", "06": "Lubelskie",
    "08": "Lubuskie", "10": "Łódzkie", "12": "Małopolskie",
    "14": "Mazowieckie", "16": "Opolskie", "18": "Podkarpackie",
    "20": "Podlaskie", "22": "Pomorskie", "24": "Śląskie",
    "26": "Świętokrzyskie", "28": "Warmińsko-Mazurskie",
    "30": "Wielkopolskie", "32": "Zachodniopomorskie",
}

# Definicje 3 tabel GPKG: jak je zunifikować do wspólnych kolumn.
# Mapowanie: kolumna_w_GPKG → kolumna_w_wyniku.
TABLE_DEFS = [
    {
        "table": "transakcje_lokale",
        "typ": "lokal",
        "id_field": "lok_id_lokalu",
        "area_field": "lok_pow_uzyt",
        "item_price_field": "lok_cena_brutto",
        "address_field": "lok_adres",
        "type_specific": {
            "lok_funkcja": "funkcja_lokalu",
            "lok_liczba_izb": "liczba_izb",
            "lok_nr_kond": "nr_kondygnacji",
            "lok_nr_lokalu": "nr_lokalu",
            "lok_pow_przyn": "pow_przynalezna_m2",
        },
    },
    {
        "table": "transakcje_dzialki",
        "typ": "działka",
        "id_field": "dzi_id_dzialki",
        "area_field": "dzi_pow_ewid",
        "item_price_field": "dzi_cena_brutto",
        "address_field": "dzi_adres",
        "type_specific": {
            "dzi_nr_dzialki": "nr_dzialki",
            "dzi_przezn_wmpzp": "przeznaczenie_mpzp",
            "dzi_sposob_uzyt": "sposob_uzytkowania",
        },
    },
    {
        "table": "transakcje_budynki",
        "typ": "budynek",
        "id_field": "bud_id_budynku",
        "area_field": "bud_pow_uzyt",
        "item_price_field": "bud_cena_brutto",
        "address_field": "bud_adres",
        "type_specific": {
            "bud_nr_budynku": "nr_budynku",
            "bud_rodzaj": "rodzaj_budynku",
        },
    },
]

# Kolumny wspólne dla każdej tabeli — pobierane bez zmiany nazwy.
COMMON_COLS = [
    "fid", "teryt",
    "tran_lokalny_id_iip", "tran_rodzaj_trans", "tran_rodzaj_rynku",
    "tran_sprzedajacy", "tran_kupujacy",
    "tran_cena_brutto", "tran_vat",
    "dok_data",
    "nier_rodzaj", "nier_prawo", "nier_udzial",
    "nier_pow_gruntu", "nier_cena_brutto", "nier_vat",
]

# Format adresu w RCN: "MSC:Warszawa;UL:ulica X;NR_PORZ:10".
ADR_RE = re.compile(r"(MSC|UL|NR_PORZ|MSCS|GMI|POW|WOJ):([^;]*)")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("rcn-gpkg")


# ---------------------------------------------------------------------------
# Wczytywanie
# ---------------------------------------------------------------------------


def _table_exists(con: sqlite3.Connection, name: str) -> bool:
    r = con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone()
    return r is not None


def _build_select(td: dict, con: sqlite3.Connection) -> str:
    """Zbuduj SQL pobierający wspólne kolumny + specyficzne, w jednolitym aliasing."""
    # Sprawdź które kolumny faktycznie są w tabeli (różne wersje schematu).
    existing = {
        c[1] for c in con.execute(f'PRAGMA table_info("{td["table"]}")').fetchall()
    }
    parts: list[str] = []
    for c in COMMON_COLS:
        if c in existing:
            parts.append(f'"{c}"')
        else:
            parts.append(f'NULL AS "{c}"')
    parts.append(f'"{td["id_field"]}" AS id_obiektu')
    parts.append(f'"{td["area_field"]}" AS powierzchnia_m2')
    parts.append(f'"{td["item_price_field"]}" AS cena_obiektu_brutto')
    parts.append(f'"{td["address_field"]}" AS adres_raw')
    parts.append(f"'{td['typ']}' AS typ")
    for src, dst in td["type_specific"].items():
        if src in existing:
            parts.append(f'"{src}" AS {dst}')
        else:
            parts.append(f'NULL AS {dst}')
    return f'SELECT {", ".join(parts)} FROM "{td["table"]}"'


def load_gpkg(path: Path) -> pd.DataFrame:
    log.info("Czytam: %s", path)
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    try:
        frames: list[pd.DataFrame] = []
        for td in TABLE_DEFS:
            if not _table_exists(con, td["table"]):
                log.warning("  brak tabeli %s — pomijam", td["table"])
                continue
            sql = _build_select(td, con)
            df = pd.read_sql_query(sql, con)
            log.info("  %s: %d wierszy", td["table"], len(df))
            frames.append(df)
        if not frames:
            return pd.DataFrame()
        return pd.concat(frames, ignore_index=True)
    finally:
        con.close()


# ---------------------------------------------------------------------------
# Wzbogacanie i filtry
# ---------------------------------------------------------------------------


def parse_adres(text) -> dict[str, str]:
    if not isinstance(text, str) or not text:
        return {}
    out = {}
    for k, v in ADR_RE.findall(text):
        v = v.strip()
        if not v:
            continue
        if k == "MSC":
            out["miejscowosc"] = v
        elif k == "MSCS":
            out["miejscowosc_szczegolowa"] = v
        elif k == "UL":
            # Usuwamy prefiks "ulica" / "aleja" / itp.
            out["ulica"] = re.sub(r"^(ulica|aleja|aleje|plac|skwer|rondo|park|bulwar)\s+", "", v, flags=re.I)
        elif k == "NR_PORZ":
            out["numer"] = v
        elif k == "GMI":
            out["gmina"] = v
        elif k == "POW":
            out["powiat_adres"] = v
    return out


def derive_powiat_name(teryt: str) -> str:
    teryt = teryt or ""
    if teryt in KNOWN_POWIATS:
        return KNOWN_POWIATS[teryt]
    return f"powiat {teryt}"


def derive_dzielnica(id_obiektu: Optional[str], teryt: str) -> Optional[str]:
    """Dla Warszawy: dzielnica po prefiksie 146502..146519."""
    if not id_obiektu or teryt != "1465":
        return None
    return WARSAW_DISTRICTS.get(id_obiektu[:6])


def derive_wojewodztwo(teryt: str) -> str:
    return WOJEWODZTWA.get((teryt or "")[:2], "")


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    # Adres
    parsed = df["adres_raw"].map(parse_adres).tolist()
    addr_df = pd.DataFrame(parsed, index=df.index)
    df = pd.concat([df, addr_df], axis=1)

    # Województwo / powiat / dzielnica
    df["wojewodztwo"] = df["teryt"].map(derive_wojewodztwo)
    df["powiat"] = df["teryt"].map(derive_powiat_name)
    df["dzielnica"] = df.apply(
        lambda r: derive_dzielnica(r.get("id_obiektu"), r.get("teryt", "")), axis=1
    )

    # Data i cena za m²
    df["dok_data"] = pd.to_datetime(df["dok_data"], errors="coerce", utc=True)
    df["data_transakcji"] = df["dok_data"].dt.date

    df["cena_za_m2"] = (
        df["cena_obiektu_brutto"].fillna(df["nier_cena_brutto"]).fillna(df["tran_cena_brutto"])
        / df["powierzchnia_m2"]
    ).round(2)
    df.loc[df["powierzchnia_m2"].fillna(0) <= 0, "cena_za_m2"] = None

    return df


def filter_df(
    df: pd.DataFrame,
    *,
    years: int,
    types_allowed: set[str],
    sales_only: bool,
    mieszkalne_only: bool,
) -> pd.DataFrame:
    if df.empty:
        return df
    before = len(df)

    if years > 0:
        today = date.today()
        try:
            cutoff = date(today.year - years, today.month, today.day)
        except ValueError:
            cutoff = date(today.year - years, today.month, 28)
        df = df[df["data_transakcji"].notna() & (df["data_transakcji"] >= cutoff)]
        log.info("Filtr dat (>=%s): %d → %d", cutoff, before, len(df))
        before = len(df)

    if types_allowed:
        df = df[df["typ"].str.lower().isin(types_allowed)]
        log.info("Filtr typu %s: %d → %d", types_allowed, before, len(df))
        before = len(df)

    if sales_only:
        df = df[df["tran_rodzaj_trans"].astype(str).str.lower().str.contains("wolnyrynek")]
        log.info("Filtr wolnyRynek: %d → %d", before, len(df))
        before = len(df)

    if mieszkalne_only:
        mask = (df["typ"] != "lokal") | (
            df["funkcja_lokalu"].astype(str).str.lower().eq("mieszkalna")
        )
        df = df[mask]
        log.info("Filtr funkcja_lokalu=mieszkalna: %d → %d", before, len(df))

    return df


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

SHEET_RE = re.compile(r"[\\/*?:\[\]]")


def _safe_sheet(name: str, used: set[str]) -> str:
    base = SHEET_RE.sub("_", name or "Bez_nazwy")[:31] or "Arkusz"
    out = base
    i = 1
    while out in used:
        suffix = f"_{i}"
        out = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(out)
    return out


OUTPUT_COLS = [
    "wojewodztwo", "powiat", "dzielnica",
    "data_transakcji", "typ",
    "cena_obiektu_brutto", "powierzchnia_m2", "cena_za_m2",
    "tran_cena_brutto", "nier_cena_brutto",
    "miejscowosc", "ulica", "numer",
    "tran_rodzaj_trans", "tran_rodzaj_rynku",
    "tran_sprzedajacy", "tran_kupujacy",
    "nier_rodzaj", "nier_prawo", "nier_udzial", "nier_pow_gruntu",
    # lokal-specific
    "funkcja_lokalu", "liczba_izb", "nr_kondygnacji",
    "nr_lokalu", "pow_przynalezna_m2",
    # działka-specific
    "nr_dzialki", "przeznaczenie_mpzp", "sposob_uzytkowania",
    # budynek-specific
    "nr_budynku", "rodzaj_budynku",
    # identyfikatory / źródło
    "id_obiektu", "tran_lokalny_id_iip",
    "teryt", "adres_raw",
]


def write_excel(df: pd.DataFrame, out_path: Path, group_by: str) -> None:
    if df.empty:
        log.warning("Brak transakcji — nie zapisuję.")
        return

    df = df.reindex(columns=[c for c in OUTPUT_COLS if c in df.columns])
    df = df.sort_values(
        by=["wojewodztwo", "powiat", "dzielnica", "data_transakcji"],
        ascending=[True, True, True, False],
        na_position="last",
    )

    n_obiektow = df["teryt"].nunique()
    log.info("Zapisuję %d wierszy do %s (powiatów: %d)", len(df), out_path, n_obiektow)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Sumary
        summary_cols = ["wojewodztwo", "powiat"]
        if df["dzielnica"].notna().any():
            summary_cols.append("dzielnica")
        summary = (
            df.groupby(summary_cols + ["typ"], dropna=False)
            .agg(
                liczba=("tran_lokalny_id_iip", "count"),
                średnia_cena=("cena_obiektu_brutto", "mean"),
                średnia_cena_m2=("cena_za_m2", "mean"),
                mediana_cena_m2=("cena_za_m2", "median"),
                min_cena_m2=("cena_za_m2", "min"),
                max_cena_m2=("cena_za_m2", "max"),
            )
            .round(2)
            .reset_index()
        )
        summary.to_excel(writer, sheet_name="_Podsumowanie", index=False)

        # Arkusze
        used: set[str] = set()
        used.add("_Podsumowanie")

        if group_by == "dzielnica" and df["dzielnica"].notna().any():
            for d in sorted(df["dzielnica"].dropna().unique()):
                sub = df[df["dzielnica"] == d]
                sub.to_excel(writer, sheet_name=_safe_sheet(d, used), index=False)
            others = df[df["dzielnica"].isna()]
            if not others.empty:
                others.to_excel(writer, sheet_name=_safe_sheet("Pozostałe", used), index=False)
        elif group_by == "typ":
            for typ in sorted(df["typ"].dropna().unique()):
                sub = df[df["typ"] == typ]
                sub.to_excel(writer, sheet_name=_safe_sheet(typ, used), index=False)
        elif group_by == "powiat":
            for (teryt, powiat), sub in df.groupby(["teryt", "powiat"], dropna=False):
                sub.to_excel(writer, sheet_name=_safe_sheet(f"{teryt} {powiat}", used), index=False)
        else:  # all-in-one
            df.to_excel(writer, sheet_name="Transakcje", index=False)

    log.info("Gotowe: %s", out_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        help="Ścieżki do plików .gpkg (lub wzorce shellowe).",
    )
    parser.add_argument("--out", default="rcn.xlsx", help="Plik wynikowy (domyślnie rcn.xlsx)")
    parser.add_argument(
        "--years", type=int, default=3,
        help="Ile lat wstecz. 0 = bez filtra dat (cała historia). Domyślnie 3.",
    )
    parser.add_argument(
        "--types", default="lokal,działka,budynek",
        help="Typy nieruchomości po przecinku. Pusty string = wszystko.",
    )
    parser.add_argument(
        "--mieszkalne-tylko", action="store_true",
        help="Dla lokali zostaw tylko funkcja_lokalu=mieszkalna.",
    )
    parser.add_argument(
        "--wszystkie-transakcje", action="store_true",
        help="Nie filtruj rodzaju transakcji (domyślnie tylko wolnyRynek).",
    )
    parser.add_argument(
        "--group-by", default="auto",
        choices=["auto", "dzielnica", "powiat", "typ", "all-in-one"],
        help="Sposób podziału na arkusze. auto = dzielnica jeśli Warszawa, "
             "inaczej powiat (lub typ gdy 1 powiat).",
    )
    args = parser.parse_args()

    # Rozwiń wzorce / zwaliduj ścieżki
    paths: list[Path] = []
    for inp in args.inputs:
        p = Path(inp)
        if "*" in inp or "?" in inp:
            matches = list(Path().glob(inp))
            if not matches:
                log.warning("Wzorzec '%s' nic nie pasuje", inp)
            paths.extend(matches)
        elif not p.exists():
            log.error("Nie ma takiego pliku: %s", p)
            return 1
        else:
            paths.append(p)

    if not paths:
        log.error("Brak plików wejściowych.")
        return 1

    # Wczytaj wszystkie
    frames = [load_gpkg(p) for p in paths]
    df = pd.concat([f for f in frames if not f.empty], ignore_index=True)
    log.info("Łącznie wczytano %d wierszy z %d plików", len(df), len(paths))

    # Wzbogać + filtruj
    df = enrich(df)
    types_allowed = {t.strip().lower() for t in args.types.split(",") if t.strip()}
    df = filter_df(
        df,
        years=args.years,
        types_allowed=types_allowed,
        sales_only=not args.wszystkie_transakcje,
        mieszkalne_only=args.mieszkalne_tylko,
    )

    if df.empty:
        log.error("Po filtrach brak wierszy.")
        return 1

    # Wybór grupowania
    group_by = args.group_by
    if group_by == "auto":
        if df["dzielnica"].notna().any():
            group_by = "dzielnica"
        elif df["teryt"].nunique() > 1:
            group_by = "powiat"
        else:
            group_by = "typ"
        log.info("Grupowanie arkuszy: auto → %s", group_by)

    write_excel(df, Path(args.out), group_by)
    return 0


if __name__ == "__main__":
    sys.exit(main())
