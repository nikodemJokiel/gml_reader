#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GML Reader for RCiWN (Rejestr Cen i Wartosci Nieruchomosci).

Parses GML files containing real estate transaction data from the Polish
RCiWN registry. Resolves relational xlink:href references between
transactions, documents, properties, plots, buildings, premises, and addresses.

Generates an Excel file with 3 sheets:
  - Dzialki: transactions involving only plots (no buildings/premises)
  - Budynki: transactions with buildings (but no premises)
  - Lokale:  transactions with premises (even if buildings present too)

Optionally reads TERC.xml (auto-detected) to resolve TERYT codes to names.

Required libraries: pandas, openpyxl
Required libraries: pandas, openpyxl, dbfread
Usage: python gml_reader.py <file.gml>
       Output file name matches the input file name with .xlsx extension.
       TERC.xml is auto-detected next to the GML file or in the working directory.
"""

import sys
import re
import os
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd
import dbfread


# ============================================================================
# Constants & mappings
# ============================================================================

# Ownership shares considered "full" (no remark generated)
FULL_SHARES = {"", "1/1", "1", "100%"}

# Buyer/seller type codes -> human-readable labels
PARTY_TYPE_MAP = {
    "1": "Skarb Panstwa",
    "2": "jednostka samorzadu terytorialnego",
    "3": "osoba fizyczna",
    "4": "osoba prawna inna niz Skarb Panstwa i jednostka samorzadu terytorialnego",
    "5": "Spoldzielnia mieszkaniowa",
    "6": "Inne",
}

# Transaction type codes -> human-readable labels
TRANSACTION_TYPE_MAP = {
    "1": "wolny rynek",
    "2": "sprzedaz bezprzetargowa nieruchomosci Skarbu Panstwa lub jednostek samorzadu terytorialnego",
    "3": "sprzedaz przetargowa nieruchomosci Skarbu Panstwa lub jednostek samorzadu terytorialnego",
    "4": "sprzedaz w postepowaniu egzekucyjnym",
    "5": "sprzedaz na cel publiczny",
    "6": "sprzedaz z zastosowaniem bonifikaty",
}

# Market type codes
MARKET_TYPE_MAP = {
    "1": "Rynek pierwotny",
    "2": "Rynek wtorny",
}

# Property type codes (rodzajNieruchomosci)
PROPERTY_TYPE_MAP = {
    "1": "Nieruchomosc gruntowa niezabudowana",
    "2": "Nieruchomosc gruntowa zabudowana",
    "3": "Nieruchomosc budynkowa",
    "4": "Nieruchomosc lokalowa",
}

# Property right codes (rodzajPrawaDoNieruchomosci)
PROPERTY_RIGHT_MAP = {
    "1": "Wlasnosc nieruchomosci gruntowej",
    "2": "Uzytkowanie wieczyste",
    "3": "Wlasnosc lokalu z prawem do nieruchomosci wspolnej",
}

# Land use codes (sposobUzytkowania)
LAND_USE_MAP = {
    "1": "Grunty rolne",
    "2": "Grunty lesne",
    "3": "Grunty zabudowane i zurbanizowane",
    "4": "Tereny komunikacyjne",
    "5": "Tereny inne",
    "6": "Grunty pod wodami",
    "7": "Nieuzytki",
}

# Building type codes (rodzajBudynku) — PKOB / KST
BUILDING_TYPE_MAP = {
    "101": "Budynki przemyslowe",
    "102": "Budynki transportu i lacznosci",
    "103": "Budynki handlowo-uslugowe",
    "104": "Zbiorniki, silosy i budynki magazynowe",
    "105": "Budynki biurowe",
    "106": "Budynki szpitali i zakladow opieki medycznej",
    "107": "Budynki oswiaty, nauki, kultury i sportowe",
    "108": "Budynki produkcyjne, uslugowe i gospodarcze dla rolnictwa",
    "109": "Pozostale budynki niemieszkalne",
    "110": "Budynki mieszkalne",
    "111": "Budynki mieszkalne jednorodzinne",
    "112": "Budynki o dwoch lub wiecej mieszkaniach",
    "113": "Budynki zbiorowego zamieszkania",
    "121": "Hotele i budynki zakwaterowania turystycznego",
    "122": "Budynki biurowe",
    "123": "Budynki handlowo-uslugowe",
    "124": "Budynki transportu i lacznosci",
    "125": "Budynki przemyslowe i magazynowe",
    "126": "Budynki kultury, nauki i oswiaty",
    "127": "Budynki opieki zdrowotnej",
    "128": "Budynki kultury fizycznej",
    "130": "Budynki zbiorowego zamieszkania (pozostale)",
    "141": "Budynki gospodarcze",
    "142": "Garaże",
    # 4-digit PKOB classes often seen in data
    "1110": "Budynki mieszkalne jednorodzinne",
    "1121": "Budynki o dwoch mieszkaniach",
    "1122": "Budynki o trzech i wiecej mieszkaniach",
    "1130": "Budynki zbiorowego zamieszkania",
    "1211": "Hotele",
    "1220": "Budynki biurowe",
    "1230": "Budynki handlowo-uslugowe",
    "1241": "Budynki lacznosci, dworce i terminale",
    "1242": "Garaze",
    "1251": "Budynki przemyslowe",
    "1252": "Zbiorniki, silosy i budynki magazynowe",
    "1261": "Budynki kultury i rozrywki",
    "1262": "Muzea i biblioteki",
    "1263": "Budynki szkol i instytucji badawczych",
    "1264": "Budynki szpitali i zakladow opieki medycznej",
    "1265": "Budynki kultury fizycznej",
    "1271": "Budynki gospodarstw rolnych",
    "1272": "Budynki kultu religijnego",
    "1273": "Budynki zabytkowe",
    "1274": "Pozostale budynki niemieszkalne",
}

# Premises function codes (funkcjaLokalu)
PREMISES_FUNCTION_MAP = {
    "1": "Mieszkalna",
    "2": "Handlowo-uslugowa",
    "5": "Garaz",
    "6": "Inna niemieszkalna",
}

# Semicolon separator for multi-value fields
SEP = "; "


# ============================================================================
# TERC dictionary builder
# ============================================================================

def build_terc_dict(terc_path: str) -> dict:
    """
    Parse TERC.xml and build a lookup: WWPPGGR (7 chars) -> "Name (type)".

    GML plot ID '301909_5.0008.203/26' encodes WWPPGG='301909', RODZ='5',
    so the lookup key is '3019095'.
    Only municipality-level rows (all 4 codes present) are included.
    """
    terc = {}
    try:
        tree = ET.parse(terc_path)
        root = tree.getroot()

        for row in root.iter("row"):
            woj = (row.findtext("WOJ") or "").strip()
            pow_ = (row.findtext("POW") or "").strip()
            gmi = (row.findtext("GMI") or "").strip()
            rodz = (row.findtext("RODZ") or "").strip()
            nazwa = (row.findtext("NAZWA") or "").strip()
            nazwa_dod = (row.findtext("NAZWA_DOD") or "").strip()

            if not (woj and pow_):
                continue
                
            if not (gmi and rodz):
                # For cities on powiat rights, TERC often leaves GMI and RODZ empty, 
                # but plots use 01 and 1 respectively.
                if "na prawach powiatu" in nazwa_dod:
                    gmi = "01"
                    rodz = "1"
                else:
                    continue

            key = f"{woj}{pow_}{gmi}{rodz}"
            label = f"{nazwa} ({nazwa_dod})" if nazwa_dod else nazwa
            terc[key] = label

    except ET.ParseError as e:
        print(f"[WARN] Could not parse TERC.xml: {e}")
    except FileNotFoundError:
        print(f"[WARN] TERC.xml not found at: {terc_path}")

    return terc


def build_precinct_dict(dbf_path: str) -> dict:
    """
    Load DBF file and map JPT_KOD_JE to JPT_NAZWA_
    """
    precinct_dict = {}
    if not os.path.exists(dbf_path):
        print(f"[!] Warning: {dbf_path} not found. Precinct names will be empty.")
        return precinct_dict

    try:
        db = dbfread.DBF(dbf_path, encoding='utf-8')
        for row in db:
            code = str(row.get('JPT_KOD_JE', '')).strip()
            name = str(row.get('JPT_NAZWA_', '')).strip()
            if code and name:
                precinct_dict[code] = name
    except Exception as e:
        print(f"[!] Error reading {dbf_path}: {e}")

    return precinct_dict


def resolve_municipality(plot_id: str, terc_dict: dict) -> str:
    """
    Resolve TERYT code from plot ID to municipality name via TERC dict.
    Plot ID '301909_5.0008.203/26' -> key '3019095' -> "Wysoka (obszar wiejski)"
    Falls back to raw 6-digit code if no TERC match.
    """
    match = re.match(r"^(\d{6})_(\d)", plot_id)
    if not match:
        return extract_municipality_code(plot_id)

    wwppgg = match.group(1)
    rodz = match.group(2)
    key = f"{wwppgg}{rodz}"

    if terc_dict and key in terc_dict:
        return terc_dict[key]

    return wwppgg


# ============================================================================
# Helper functions
# ============================================================================

def detect_namespaces(gml_path: str) -> dict:
    """
    Dynamically detect XML namespaces from the GML file.
    Looks for: rcn/rcwn (RCiWN data), gml, xlink.
    """
    raw_namespaces = {}
    for event, elem in ET.iterparse(gml_path, events=["start-ns"]):
        prefix, uri = elem
        raw_namespaces[prefix] = uri

    ns = {}

    for prefix, uri in raw_namespaces.items():
        uri_lower = uri.lower()
        if "rcn" in uri_lower or "rcwn" in uri_lower or "rciwn" in uri_lower or "cenwartosci" in uri_lower:
            ns["rcn"] = uri
        elif "opengis.net/gml" in uri_lower:
            ns["gml"] = uri
        elif "xlink" in uri_lower:
            ns["xlink"] = uri

    # Fallback by prefix name
    if "rcn" not in ns:
        for prefix, uri in raw_namespaces.items():
            if prefix in ("rcn", "rcwn"):
                ns["rcn"] = uri
                break

    if "gml" not in ns:
        for prefix, uri in raw_namespaces.items():
            if prefix == "gml":
                ns["gml"] = uri
                break

    if "xlink" not in ns:
        for prefix, uri in raw_namespaces.items():
            if prefix == "xlink":
                ns["xlink"] = uri
                break

    ns.setdefault("gml", "http://www.opengis.net/gml/3.2")
    ns.setdefault("xlink", "http://www.w3.org/1999/xlink")

    return ns


def get_text(element: ET.Element, tag: str, ns: dict) -> str:
    """Extract text content from a child element. Returns empty string if missing."""
    child = element.find(tag, ns)
    if child is not None and child.text:
        return child.text.strip()
    return ""


def get_href(element: ET.Element, tag: str, ns: dict) -> str:
    """Extract xlink:href from a child element, stripping leading '#'."""
    child = element.find(tag, ns)
    if child is not None:
        href = child.get(f"{{{ns['xlink']}}}href", "")
        return href.lstrip("#")
    return ""


def get_all_hrefs(element: ET.Element, tag: str, ns: dict) -> list:
    """Extract all xlink:href values from matching child elements."""
    hrefs = []
    for child in element.findall(tag, ns):
        href = child.get(f"{{{ns['xlink']}}}href", "")
        if href:
            hrefs.append(href.lstrip("#"))
    return hrefs


def extract_municipality_code(plot_id: str) -> str:
    """Extract 6-digit TERYT code: '301909_5.0008.203/26' -> '301909'."""
    match = re.match(r"^(\d{6})", plot_id)
    return match.group(1) if match else ""


def extract_plot_number(plot_id: str) -> str:
    """Extract plot number: '301909_5.0008.203/26' -> '203/26' (after last dot)."""
    if "." in plot_id:
        return plot_id.rsplit(".", maxsplit=1)[-1]
    return plot_id


def hectares_to_sqm(value_ha: str) -> int | None:
    """Convert hectares to square meters. Returns rounded int or None."""
    try:
        return round(float(value_ha) * 10000)
    except (ValueError, TypeError):
        return None


def safe_float(value: str) -> float | None:
    """Parse string to float, return None on failure."""
    try:
        return float(value.replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError, AttributeError):
        return None


def join_unique(values: list[str], separator: str = SEP) -> str:
    """Join non-empty values, preserving order but removing duplicates."""
    seen = set()
    result = []
    for v in values:
        if v and v not in seen:
            seen.add(v)
            result.append(v)
    return separator.join(result)


def format_address(addr_data: dict) -> str:
    """Format address dict into a single string: 'locality, street number'."""
    parts = []
    locality = addr_data.get("locality", "")
    street = addr_data.get("street", "")
    number = addr_data.get("number", "")

    if locality:
        parts.append(locality)
    if street:
        addr_str = street
        if number:
            addr_str += f" {number}"
        parts.append(addr_str)
    elif number:
        parts.append(number)

    return ", ".join(parts)


# ============================================================================
# Building reference lookup dictionaries
# ============================================================================

def build_document_dict(root: ET.Element, ns: dict) -> dict:
    """
    Documents: gml:id -> dict with date and extended metadata.
    Fields: date, oznaczenie, tworca, sygnatura, opis, rodzaj, numer_zmiany.
    """
    documents = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Dokument"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Dokument"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue

            date = get_text(elem, f"{{{ns['rcn']}}}dataSporzadzeniaDokumentu", ns)
            if not date:
                date = get_text(elem, f"{{{ns['rcn']}}}dataDokumentu", ns)

            documents[gml_id] = {
                "date": date,
                "oznaczenie": get_text(elem, f"{{{ns['rcn']}}}oznaczenieDokumentu", ns),
                "tworca": get_text(elem, f"{{{ns['rcn']}}}tworcaDokumentu", ns),
            }

    return documents


def build_address_dict(root: ET.Element, ns: dict) -> dict:
    """Addresses: gml:id -> dict with locality, street, number."""
    addresses = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Adres"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Adres"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue
            addresses[gml_id] = {
                "locality": get_text(elem, f"{{{ns['rcn']}}}miejscowosc", ns),
                "street": get_text(elem, f"{{{ns['rcn']}}}ulica", ns),
                "number": get_text(elem, f"{{{ns['rcn']}}}numerPorzadkowy", ns),
            }

    return addresses


def build_plot_dict(root: ET.Element, ns: dict) -> dict:
    """Plots: gml:id -> dict with plot data + address reference."""
    plots = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Dzialka"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Dzialka"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue

            plot_id = get_text(elem, f"{{{ns['rcn']}}}idDzialki", ns)

            area_elem = elem.find(f"{{{ns['rcn']}}}polePowierzchniEwidencyjnej", ns)
            area_ha = ""
            if area_elem is not None and area_elem.text:
                area_ha = area_elem.text.strip()

            plots[gml_id] = {
                "plot_id": plot_id,
                "area_ha": area_ha,
                "designation": get_text(elem, f"{{{ns['rcn']}}}przeznaczenieWMPZP", ns),
                "land_use": get_text(elem, f"{{{ns['rcn']}}}sposobUzytkowania", ns),
                "additional_info": get_text(elem, f"{{{ns['rcn']}}}dodatkoweInformacje", ns),
                "address_href": get_href(elem, f"{{{ns['rcn']}}}adresDzialki", ns),
            }

    return plots


def build_property_dict(root: ET.Element, ns: dict) -> dict:
    """
    Properties: gml:id -> dict with share, price, type, right, and
    lists of plot/building/premises refs.
    """
    properties = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Nieruchomosc"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Nieruchomosc"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue

            properties[gml_id] = {
                "share": get_text(elem, f"{{{ns['rcn']}}}udzialWPrawieDoNieruchomosci", ns),
                "property_price_gross": get_text(elem, f"{{{ns['rcn']}}}cenaNieruchomosciBrutto", ns),
                "property_type": get_text(elem, f"{{{ns['rcn']}}}rodzajNieruchomosci", ns),
                "property_right": get_text(elem, f"{{{ns['rcn']}}}rodzajPrawaDoNieruchomosci", ns),
                "plot_hrefs": get_all_hrefs(elem, f"{{{ns['rcn']}}}dzialka", ns),
                "building_hrefs": get_all_hrefs(elem, f"{{{ns['rcn']}}}budynek", ns),
                "premises_hrefs": get_all_hrefs(elem, f"{{{ns['rcn']}}}lokal", ns),
            }

    return properties


def build_building_dict(root: ET.Element, ns: dict) -> dict:
    """
    Buildings: gml:id -> dict with building ID, type, usable area,
    address ref, and additional info.
    """
    buildings = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Budynek"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Budynek"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue

            area_elem = elem.find(f"{{{ns['rcn']}}}powierzchniaUzytkowaBudynku", ns)
            area_m2 = ""
            if area_elem is not None and area_elem.text:
                area_m2 = area_elem.text.strip()

            buildings[gml_id] = {
                "building_id": get_text(elem, f"{{{ns['rcn']}}}idBudynku", ns),
                "building_type": get_text(elem, f"{{{ns['rcn']}}}rodzajBudynku", ns),
                "area_m2": area_m2,
                "additional_info": get_text(elem, f"{{{ns['rcn']}}}dodatkoweInformacje", ns),
                "address_href": get_href(elem, f"{{{ns['rcn']}}}adresBudynku", ns),
            }

    return buildings


def build_premises_dict(root: ET.Element, ns: dict) -> dict:
    """
    Premises (lokale): gml:id -> dict with premises ID, function,
    usable area, ancillary area, rooms, floor, address ref, additional info.
    """
    premises = {}
    target_tag = f"{{{ns['rcn']}}}RCN_Lokal"

    for elem in root.iter():
        if elem.tag == target_tag or elem.tag.endswith("}RCN_Lokal"):
            gml_id = elem.get(f"{{{ns['gml']}}}id", "")
            if not gml_id:
                continue

            # Check for inline address (Poznań format)
            inline_address = {}
            addr_container = elem.find(f"{{{ns['rcn']}}}adresBudynkuZLokalem", ns)
            if addr_container is not None:
                inline_addr_elem = addr_container.find(f"{{{ns['rcn']}}}RCN_Adres", ns)
                if inline_addr_elem is None:
                    # Try without namespace (some files)
                    for child in addr_container:
                        if child.tag.endswith("}RCN_Adres") or child.tag == "RCN_Adres":
                            inline_addr_elem = child
                            break
                if inline_addr_elem is not None:
                    inline_address = {
                        "locality": get_text(inline_addr_elem, f"{{{ns['rcn']}}}miejscowosc", ns),
                        "street": get_text(inline_addr_elem, f"{{{ns['rcn']}}}ulica", ns),
                        "number": get_text(inline_addr_elem, f"{{{ns['rcn']}}}numerPorzadkowy", ns),
                    }

            premises[gml_id] = {
                "premises_id": get_text(elem, f"{{{ns['rcn']}}}idLokalu", ns),
                "function": get_text(elem, f"{{{ns['rcn']}}}funkcjaLokalu", ns),
                "area_m2": get_text(elem, f"{{{ns['rcn']}}}powUzytkowaLokalu", ns),
                "ancillary_area_m2": get_text(elem, f"{{{ns['rcn']}}}powUzytkowaPomieszczenPrzynal", ns),
                "room_count": get_text(elem, f"{{{ns['rcn']}}}liczbaIzb", ns),
                "floor": get_text(elem, f"{{{ns['rcn']}}}nrKondygnacji", ns),
                "additional_info": get_text(elem, f"{{{ns['rcn']}}}dodatkoweInformacje", ns),
                "address_href": get_href(elem, f"{{{ns['rcn']}}}adresBudynkuZLokalem", ns),
                "inline_address": inline_address,
                "premises_price": get_text(elem, f"{{{ns['rcn']}}}cenaLokaluBrutto", ns),
                "premises_vat": get_text(elem, f"{{{ns['rcn']}}}kwotaPodatkuVAT", ns),
            }

    return premises


# ============================================================================
# Transaction processing — common data extraction helpers
# ============================================================================

def extract_common_transaction_data(transaction, ns: dict, doc_dict: dict) -> dict:
    """Extract fields common to all 3 sheet types from a transaction element."""
    transaction_id = get_text(transaction, f"{{{ns['rcn']}}}oznaczenieTransakcji", ns)
    transaction_type = get_text(transaction, f"{{{ns['rcn']}}}rodzajTransakcji", ns)
    buyer_type = get_text(transaction, f"{{{ns['rcn']}}}stronaKupujaca", ns)
    seller_type = get_text(transaction, f"{{{ns['rcn']}}}stronaSprzedajaca", ns)
    transaction_price = get_text(transaction, f"{{{ns['rcn']}}}cenaTransakcjiBrutto", ns)
    market_type = get_text(transaction, f"{{{ns['rcn']}}}rodzajRynku", ns)

    # VAT -- try several possible tag names
    vat = ""
    for vat_tag_name in ("stawkaVAT", "VAT", "vat", "podatekVAT", "kwotaVAT", "wartoscVAT"):
        vat_elem = transaction.find(f"{{{ns['rcn']}}}{vat_tag_name}", ns)
        if vat_elem is not None and vat_elem.text:
            vat = vat_elem.text.strip()
            break

    # Reference to document -> transaction date + document metadata
    doc_href = get_href(transaction, f"{{{ns['rcn']}}}podstawaPrawna", ns)
    doc_data = doc_dict.get(doc_href, {})

    return {
        "ID transakcji": transaction_id,
        "Skrocone ID": transaction_id[:9] if transaction_id else "",
        "Data transakcji": doc_data.get("date", ""),
        "Rodzaj transakcji": TRANSACTION_TYPE_MAP.get(transaction_type, transaction_type),
        "Rodzaj rynku": MARKET_TYPE_MAP.get(market_type, market_type),
        "Strona kupujaca": PARTY_TYPE_MAP.get(buyer_type, buyer_type),
        "Strona sprzedajaca": PARTY_TYPE_MAP.get(seller_type, seller_type),
        "Cena [PLN]": transaction_price,
        "VAT": vat,
        "Oznaczenie dokumentu": doc_data.get("oznaczenie", ""),
        "Tworca dokumentu": doc_data.get("tworca", ""),
    }


def collect_plot_data(property_hrefs: list, prop_dict: dict, plot_dict: dict,
                      addr_dict: dict, terc_dict: dict, precinct_dict: dict) -> dict:
    """
    Aggregate plot data across all properties in a transaction.
    Returns a dict with aggregated plot fields + share remarks + property metadata.
    """
    all_plot_numbers = []
    all_plot_full_ids = []
    all_localities = []
    all_streets = []
    all_municipalities = []
    all_precinct_names = []
    all_designations = []
    all_land_uses = []
    all_additional_info = []
    all_share_remarks = []
    all_property_types = []
    all_property_rights = []
    all_property_prices = []
    total_area_sqm = 0
    area_valid = False
    fallback_price_sum = 0
    fallback_area_sum = 0

    for prop_href in property_hrefs:
        prop_data = prop_dict.get(prop_href, {})
        share = prop_data.get("share", "")
        prop_price = prop_data.get("property_price_gross", "")

        # Accumulate fallback price
        parsed_price = safe_float(prop_price)
        if parsed_price is not None:
            fallback_price_sum += parsed_price

        # Property metadata
        ptype = prop_data.get("property_type", "")
        if ptype:
            all_property_types.append(PROPERTY_TYPE_MAP.get(ptype, ptype))
        pright = prop_data.get("property_right", "")
        if pright:
            all_property_rights.append(PROPERTY_RIGHT_MAP.get(pright, pright))

        # Check for fractional ownership
        is_fractional = share and share not in FULL_SHARES

        # Always record the share value for proper ordering
        all_share_remarks.append(share if share else "1/1")

        # Collect property-level price
        all_property_prices.append(prop_price if prop_price else "")

        # Resolve all plots for this property
        for pl_href in prop_data.get("plot_hrefs", []):
            pl_data = plot_dict.get(pl_href, {})
            plot_id = pl_data.get("plot_id", "")

            if plot_id:
                all_plot_numbers.append(extract_plot_number(plot_id))
                all_plot_full_ids.append(plot_id)
                all_municipalities.append(resolve_municipality(plot_id, terc_dict))
                # Precinct resolution
                parts = plot_id.split('.')
                if len(parts) >= 2:
                    precinct_code = f"{parts[0]}.{parts[1]}"
                    all_precinct_names.append(precinct_dict.get(precinct_code, ""))
                else:
                    all_precinct_names.append("")
            else:
                all_precinct_names.append("")

            # Resolve address
            address_href = pl_data.get("address_href", "")
            addr_data = addr_dict.get(address_href, {})
            locality = addr_data.get("locality", "")
            street = addr_data.get("street", "")
            if locality:
                all_localities.append(locality)
            if street:
                all_streets.append(street)

            # Accumulate area ONLY if full share
            area_ha = pl_data.get("area_ha", "")
            area_sqm = hectares_to_sqm(area_ha)
            if area_sqm is not None:
                fallback_area_sum += area_sqm
                if not is_fractional:
                    total_area_sqm += area_sqm
                    area_valid = True

            # Collect descriptive fields
            land_use_raw = pl_data.get("land_use", "")
            all_land_uses.append(LAND_USE_MAP.get(land_use_raw, land_use_raw))
            all_designations.append(pl_data.get("designation", ""))
            all_additional_info.append(pl_data.get("additional_info", ""))

    return {
        "plot_numbers": all_plot_numbers,
        "plot_full_ids": all_plot_full_ids,
        "localities": all_localities,
        "streets": all_streets,
        "municipalities": all_municipalities,
        "precinct_names": all_precinct_names,
        "designations": all_designations,
        "land_uses": all_land_uses,
        "additional_info": all_additional_info,
        "share_remarks": all_share_remarks,
        "property_types": all_property_types,
        "property_rights": all_property_rights,
        "property_prices": all_property_prices,
        "total_area_sqm": total_area_sqm if area_valid else None,
        "fallback_price_sum": fallback_price_sum,
        "fallback_area_sum": fallback_area_sum,
    }


def collect_building_data(property_hrefs: list, prop_dict: dict,
                          building_dict: dict, addr_dict: dict) -> dict:
    """
    Aggregate building data across all properties in a transaction.
    Returns lists of building fields + share remarks for building properties.
    """
    all_building_ids = []
    all_building_types = []
    all_building_areas = []
    all_building_addresses = []
    all_building_streets = []
    all_building_info = []
    all_building_share_remarks = []
    all_building_property_prices = []

    for prop_href in property_hrefs:
        prop_data = prop_dict.get(prop_href, {})
        share = prop_data.get("share", "")

        for bld_href in prop_data.get("building_hrefs", []):
            bld_data = building_dict.get(bld_href, {})
            bld_id = bld_data.get("building_id", "")
            bld_type_raw = bld_data.get("building_type", "")
            bld_area = bld_data.get("area_m2", "")
            bld_info = bld_data.get("additional_info", "")

            # Resolve building address
            bld_addr_href = bld_data.get("address_href", "")
            bld_addr = format_address(addr_dict.get(bld_addr_href, {}))

            all_building_ids.append(bld_id if bld_id else bld_href)
            all_building_types.append(BUILDING_TYPE_MAP.get(bld_type_raw, bld_type_raw))
            # Extract street from building address
            bld_addr_data = addr_dict.get(bld_addr_href, {})
            bld_street = bld_addr_data.get("street", "")

            all_building_areas.append(bld_area)
            all_building_addresses.append(bld_addr)
            all_building_streets.append(bld_street)
            all_building_info.append(bld_info)

            # Always record the share value for proper ordering
            all_building_share_remarks.append(share if share else "1/1")
            all_building_property_prices.append(prop_data.get("property_price_gross", ""))

    return {
        "building_ids": all_building_ids,
        "building_types": all_building_types,
        "building_areas": all_building_areas,
        "building_addresses": all_building_addresses,
        "building_streets": all_building_streets,
        "building_info": all_building_info,
        "building_share_remarks": all_building_share_remarks,
        "building_property_prices": all_building_property_prices,
    }


def collect_premises_data(property_hrefs: list, prop_dict: dict,
                          premises_dict: dict, addr_dict: dict) -> dict:
    """
    Aggregate premises (lokal) data across all properties in a transaction.
    Returns lists of premises fields + share remarks.
    """
    all_premises_ids = []
    all_premises_functions = []
    all_premises_areas = []
    all_premises_ancillary_areas = []
    all_premises_rooms = []
    all_premises_floors = []
    all_premises_addresses = []
    all_premises_streets = []
    all_premises_info = []
    all_premises_share_remarks = []
    all_property_prices = []
    all_premises_prices = []
    all_premises_vat = []
    total_premises_area = 0.0

    for prop_href in property_hrefs:
        prop_data = prop_dict.get(prop_href, {})
        share = prop_data.get("share", "")
        is_full_share = (not share) or (share in FULL_SHARES)

        # Property-level price
        prop_price = prop_data.get("property_price_gross", "")

        for prem_href in prop_data.get("premises_hrefs", []):
            prem_data = premises_dict.get(prem_href, {})
            prem_id = prem_data.get("premises_id", "")
            func_raw = prem_data.get("function", "")
            prem_area = prem_data.get("area_m2", "")

            # Only sum area for full-share properties (excludes fractional garages etc.)
            if prem_area and is_full_share:
                try:
                    total_premises_area += float(prem_area.replace(',', '.'))
                except ValueError:
                    pass

            prem_anc_area = prem_data.get("ancillary_area_m2", "")
            prem_rooms = prem_data.get("room_count", "")
            prem_floor = prem_data.get("floor", "")
            prem_info = prem_data.get("additional_info", "")

            # Resolve premises address — try addr_dict first, then inline (Poznań)
            prem_addr_href = prem_data.get("address_href", "")
            prem_addr_data = addr_dict.get(prem_addr_href, {})
            if not prem_addr_data:
                prem_addr_data = prem_data.get("inline_address", {})
            prem_addr = format_address(prem_addr_data)
            prem_street = prem_addr_data.get("street", "")

            all_premises_ids.append(prem_id if prem_id else prem_href)
            all_premises_functions.append(PREMISES_FUNCTION_MAP.get(func_raw, func_raw))
            all_premises_areas.append(prem_area)
            all_premises_ancillary_areas.append(prem_anc_area)
            all_premises_rooms.append(prem_rooms)
            all_premises_floors.append(prem_floor)
            all_premises_addresses.append(prem_addr)
            all_premises_streets.append(prem_street)
            all_premises_info.append(prem_info)
            all_property_prices.append(prop_price)
            all_premises_prices.append(prem_data.get("premises_price", ""))
            all_premises_vat.append(prem_data.get("premises_vat", ""))

            # Always record the share value for proper ordering
            all_premises_share_remarks.append(share if share else "1/1")

    return {
        "premises_ids": all_premises_ids,
        "premises_functions": all_premises_functions,
        "premises_areas": all_premises_areas,
        "premises_ancillary_areas": all_premises_ancillary_areas,
        "premises_rooms": all_premises_rooms,
        "premises_floors": all_premises_floors,
        "premises_addresses": all_premises_addresses,
        "premises_streets": all_premises_streets,
        "premises_info": all_premises_info,
        "premises_share_remarks": all_premises_share_remarks,
        "property_prices": all_property_prices,
        "premises_prices": all_premises_prices,
        "premises_vat": all_premises_vat,
        "total_premises_area": total_premises_area,
    }


# ============================================================================
# Main transaction processing (returns 3 separate row lists)
# ============================================================================

def process_transactions(root: ET.Element, ns: dict,
                         doc_dict: dict, addr_dict: dict,
                         plot_dict: dict, prop_dict: dict,
                         building_dict: dict, premises_dict: dict,
                         terc_dict: dict, precinct_dict: dict) -> tuple[list, list, list]:
    """
    Iterate over all transactions and sort them into 3 lists:
      - plot_rows:     transactions with only plots (no buildings/premises)
      - building_rows: transactions with buildings (no premises)
      - premises_rows: transactions with premises (even if buildings present too)
    """
    plot_rows = []
    building_rows = []
    premises_rows = []

    target_tag = f"{{{ns['rcn']}}}RCN_Transakcja"
    global_lp = 1

    for transaction in root.iter():
        if transaction.tag != target_tag and not transaction.tag.endswith("}RCN_Transakcja"):
            continue

        # References to properties
        property_hrefs = get_all_hrefs(transaction, f"{{{ns['rcn']}}}nieruchomosc", ns)
        if not property_hrefs:
            single_href = get_href(transaction, f"{{{ns['rcn']}}}nieruchomosc", ns)
            if single_href:
                property_hrefs = [single_href]

        # Determine transaction category
        has_building = False
        has_premises = False
        for prop_href in property_hrefs:
            prop_data = prop_dict.get(prop_href, {})
            if prop_data.get("building_hrefs"):
                has_building = True
            if prop_data.get("premises_hrefs"):
                has_premises = True

        # --- Common data ---
        common = extract_common_transaction_data(transaction, ns, doc_dict)
        common["Lp."] = global_lp
        global_lp += 1

        # --- Plot data (always needed in some form) ---
        pdata = collect_plot_data(
            property_hrefs, prop_dict, plot_dict, addr_dict, terc_dict, precinct_dict
        )

        # Determine final price
        price_str = common["Cena [PLN]"]
        if not price_str and pdata["fallback_price_sum"] > 0:
            price_str = str(int(pdata["fallback_price_sum"]))
        common["Cena [PLN]"] = price_str

        # --- Route to the correct sheet ---
        if has_premises:
            # PREMISES sheet (lokal)
            premdata = collect_premises_data(
                property_hrefs, prop_dict, premises_dict, addr_dict
            )

            cena_m2 = ""
            if common["Cena [PLN]"] and premdata["total_premises_area"] > 0:
                try:
                    c = float(str(common["Cena [PLN]"]).replace(',', '.'))
                    cena_m2 = round(c / premdata["total_premises_area"], 2)
                except ValueError:
                    pass

            # Use premises VAT as fallback if transaction-level VAT is empty
            vat_value = common.get("VAT", "")
            if not vat_value:
                prem_vats = [v for v in premdata["premises_vat"] if v]
                if prem_vats:
                    vat_value = SEP.join(prem_vats)

            gmina = join_unique(pdata["municipalities"])
            obreb = join_unique(pdata["precinct_names"])
            miejscowosc = join_unique(pdata["localities"])
            ulica = join_unique(pdata["streets"])
            
            # If gmina or obreb are empty (e.g. Poznań format without plots), extract from ID lokalu
            if not gmina or not obreb:
                inline_gminas = []
                inline_obreby = []
                for prem_id in premdata["premises_ids"]:
                    if prem_id:
                        if not gmina:
                            inline_gminas.append(resolve_municipality(prem_id, terc_dict))
                        if not obreb:
                            parts = prem_id.split('.')
                            if len(parts) >= 2:
                                precinct_code = f"{parts[0]}.{parts[1]}"
                                inline_obreby.append(precinct_dict.get(precinct_code, ""))
                
                if not gmina:
                    gmina = join_unique(inline_gminas)
                if not obreb:
                    obreb = join_unique(inline_obreby)

            # For Poznań: if plot-based locality/street are empty, use inline premises addresses
            if not miejscowosc:
                prem_addrs = [a for a in premdata["premises_addresses"] if a]
                inline_localities = []
                for addr_str in prem_addrs:
                    # Extract locality from formatted address "LOCALITY, street num"
                    parts = addr_str.split(",", 1)
                    if parts:
                        inline_localities.append(parts[0].strip())
                miejscowosc = join_unique(inline_localities)
            if not ulica:
                ulica = join_unique(premdata["premises_streets"])

            row = {
                **common,
                "VAT": vat_value,
                "Rodzaj nieruchomosci": join_unique(pdata["property_types"]),
                "Rodzaj prawa": join_unique(pdata["property_rights"]),
                "Gmina": gmina,
                "Obręb": obreb,
                "Miejscowosc": miejscowosc,
                "Ulica": ulica,
                "ID lokalu": SEP.join(premdata["premises_ids"]),
                "Funkcja lokalu": SEP.join(premdata["premises_functions"]),
                "Pow. uzytkowa lokalu [m2]": SEP.join(premdata["premises_areas"]),
                "Cena za m2 [PLN]": cena_m2,
                "Pow. pom. przynal. [m2]": SEP.join(premdata["premises_ancillary_areas"]),
                "Liczba izb": SEP.join(premdata["premises_rooms"]),
                "Nr kondygnacji": SEP.join(premdata["premises_floors"]),
                "Adres lokalu": SEP.join(premdata["premises_addresses"]),
                "Ulica lokalu": join_unique(premdata["premises_streets"]),
                "Info dodatkowe (lokal)": join_unique(premdata["premises_info"]),
                "Uwagi udzialy (lokal)": SEP.join(premdata["premises_share_remarks"]),
                "Ceny nieruchomosci": SEP.join(premdata["property_prices"]),
                "ID dzialki": SEP.join(pdata["plot_full_ids"]),
            }
            premises_rows.append(row)

        elif has_building:
            # BUILDING sheet (budynek)
            bdata = collect_building_data(
                property_hrefs, prop_dict, building_dict, addr_dict
            )
            row = {
                **common,
                "Rodzaj nieruchomosci": join_unique(pdata["property_types"]),
                "Rodzaj prawa": join_unique(pdata["property_rights"]),
                "Gmina": join_unique(pdata["municipalities"]),
                "Obręb": join_unique(pdata["precinct_names"]),
                "Miejscowosc": join_unique(pdata["localities"]),
                "Ulica": join_unique(pdata["streets"]),
                "Numer dzialki": SEP.join(pdata["plot_numbers"]),
                "Pow. dzialki [m2]": pdata["total_area_sqm"],
                "Przeznaczenie (MPZP)": join_unique(pdata["designations"]),
                "Sposob uzytkowania": join_unique(pdata["land_uses"]),
                "Info dodatkowe (dzialka)": join_unique(pdata["additional_info"]),
                "Uwagi udzialy (dzialka)": SEP.join(pdata["share_remarks"]),
                "ID budynku": SEP.join(bdata["building_ids"]),
                "Rodzaj budynku": join_unique(bdata["building_types"]),
                "Pow. uzytkowa budynku [m2]": SEP.join(bdata["building_areas"]),
                "Adres budynku": SEP.join(bdata["building_addresses"]),
                "Ulica budynku": join_unique(bdata["building_streets"]),
                "Info dodatkowe (budynek)": join_unique(bdata["building_info"]),
                "Uwagi udzialy (budynek)": SEP.join(bdata["building_share_remarks"]),
                "Ceny nieruchomosci": SEP.join(bdata["building_property_prices"]),
                "ID dzialki": SEP.join(pdata["plot_full_ids"]),
            }
            building_rows.append(row)

        else:
            # PLOT-only sheet (dzialka)
            row = {
                **common,
                "Rodzaj nieruchomosci": join_unique(pdata["property_types"]),
                "Rodzaj prawa": join_unique(pdata["property_rights"]),
                "Gmina": join_unique(pdata["municipalities"]),
                "Obręb": join_unique(pdata["precinct_names"]),
                "Miejscowosc": join_unique(pdata["localities"]),
                "Ulica": join_unique(pdata["streets"]),
                "Numer dzialki": SEP.join(pdata["plot_numbers"]),
                "Powierzchnia [m2]": pdata["total_area_sqm"],
                "Przeznaczenie (MPZP)": join_unique(pdata["designations"]),
                "Sposob uzytkowania": join_unique(pdata["land_uses"]),
                "Info dodatkowe (dzialka)": join_unique(pdata["additional_info"]),
                "Uwagi (Udzialy)": SEP.join(pdata["share_remarks"]),
                "Ceny nieruchomosci": SEP.join(pdata["property_prices"]),
                "ID dzialki": SEP.join(pdata["plot_full_ids"]),
            }
            plot_rows.append(row)

    return plot_rows, building_rows, premises_rows


# ============================================================================
# Excel export
# ============================================================================

# Column orders per sheet type
PLOT_COLUMNS = [
    "Lp.",
    "Skrocone ID",
    "ID transakcji",
    "Data transakcji",
    "Rodzaj transakcji",
    "Rodzaj rynku",
    "Strona kupujaca",
    "Strona sprzedajaca",
    "Rodzaj nieruchomosci",
    "Rodzaj prawa",
    "Gmina",
    "Obręb",
    "Miejscowosc",
    "Ulica",
    "Numer dzialki",
    "Powierzchnia [m2]",
    "Cena [PLN]",
    "Cena za m2 [PLN]",
    "Przeznaczenie (MPZP)",
    "Sposob uzytkowania",
    "Info dodatkowe (dzialka)",
    "Uwagi (Udzialy)",
    "Ceny nieruchomosci",
    "VAT",
    "Oznaczenie dokumentu",
    "Tworca dokumentu",
    "ID dzialki",
]

BUILDING_COLUMNS = [
    "Lp.",
    "Skrocone ID",
    "ID transakcji",
    "Data transakcji",
    "Rodzaj transakcji",
    "Rodzaj rynku",
    "Strona kupujaca",
    "Strona sprzedajaca",
    "Rodzaj nieruchomosci",
    "Rodzaj prawa",
    "Gmina",
    "Obręb",
    "Miejscowosc",
    "Ulica",
    "Numer dzialki",
    "Pow. dzialki [m2]",
    "Cena [PLN]",
    "Przeznaczenie (MPZP)",
    "Sposob uzytkowania",
    "Info dodatkowe (dzialka)",
    "Uwagi udzialy (dzialka)",
    "ID budynku",
    "Rodzaj budynku",
    "Pow. uzytkowa budynku [m2]",
    "Adres budynku",
    "Ulica budynku",
    "Info dodatkowe (budynek)",
    "Uwagi udzialy (budynek)",
    "Ceny nieruchomosci",
    "VAT",
    "Oznaczenie dokumentu",
    "Tworca dokumentu",
    "ID dzialki",
]

PREMISES_COLUMNS = [
    "Lp.",
    "Skrocone ID",
    "ID transakcji",
    "Data transakcji",
    "Rodzaj transakcji",
    "Rodzaj rynku",
    "Strona kupujaca",
    "Strona sprzedajaca",
    "Rodzaj nieruchomosci",
    "Rodzaj prawa",
    "Gmina",
    "Obręb",
    "Miejscowosc",
    "Ulica",
    "Cena [PLN]",
    "ID lokalu",
    "Funkcja lokalu",
    "Pow. uzytkowa lokalu [m2]",
    "Cena za m2 [PLN]",
    "Pow. pom. przynal. [m2]",
    "Liczba izb",
    "Nr kondygnacji",
    "Adres lokalu",
    "Ulica lokalu",
    "Info dodatkowe (lokal)",
    "Uwagi udzialy (lokal)",
    "Ceny nieruchomosci",
    "VAT",
    "Oznaczenie dokumentu",
    "Tworca dokumentu",
    "ID dzialki",
]


def prepare_sheet_df(rows: list[dict], column_order: list[str],
                     compute_price_per_sqm: bool = False,
                     area_col: str = "Powierzchnia [m2]") -> pd.DataFrame:
    """
    Build a DataFrame from rows, add Lp., optionally compute price/m2,
    and reorder columns.
    """
    if not rows:
        return pd.DataFrame(columns=column_order)

    df = pd.DataFrame(rows)

    # Type conversions for area if present
    if area_col in df.columns:
        df[area_col] = pd.to_numeric(df[area_col], errors="coerce")

    # Type conversions for price
    if "Cena [PLN]" in df.columns:
        df["Cena [PLN]"] = (
            df["Cena [PLN]"]
            .astype(str)
            .str.replace(r"\s+", "", regex=True)
            .str.replace(",", ".")
        )
        df["Cena [PLN]"] = pd.to_numeric(df["Cena [PLN]"], errors="coerce")

    # Compute price per m2 (only for plot sheet)
    if compute_price_per_sqm and area_col in df.columns:
        df["Cena za m2 [PLN]"] = df.apply(
            lambda row: round(row["Cena [PLN]"] / row[area_col], 2)
            if pd.notna(row.get("Cena [PLN]")) and pd.notna(row.get(area_col))
            and row.get("Cena [PLN]", 0) > 0 and row.get(area_col, 0) > 0
            else None,
            axis=1,
        )

    # Reorder columns (only those that exist)
    existing_columns = [c for c in column_order if c in df.columns]
    df = df[existing_columns]

    return df


def auto_fit_columns(worksheet, df: pd.DataFrame) -> None:
    """Set column widths based on content length."""
    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = max(
            len(str(col_name)),
            df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
        )
        adjusted_width = min(max_len + 3, 50)
        col_letter = worksheet.cell(row=1, column=col_idx).column_letter
        worksheet.column_dimensions[col_letter].width = adjusted_width


def export_to_excel(plot_rows: list, building_rows: list, premises_rows: list,
                    output_path: str) -> None:
    """
    Create an Excel file with 3 sheets: Dzialki, Budynki, Lokale.
    Each sheet has its own column layout and independent Lp. numbering.
    """
    df_plots = prepare_sheet_df(plot_rows, PLOT_COLUMNS,
                                compute_price_per_sqm=True,
                                area_col="Powierzchnia [m2]")
    df_buildings = prepare_sheet_df(building_rows, BUILDING_COLUMNS)
    df_premises = prepare_sheet_df(premises_rows, PREMISES_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: Dzialki
        df_plots.to_excel(writer, index=False, sheet_name="Dzialki")
        auto_fit_columns(writer.sheets["Dzialki"], df_plots)

        # Sheet 2: Budynki
        df_buildings.to_excel(writer, index=False, sheet_name="Budynki")
        auto_fit_columns(writer.sheets["Budynki"], df_buildings)

        # Sheet 3: Lokale
        df_premises.to_excel(writer, index=False, sheet_name="Lokale")
        auto_fit_columns(writer.sheets["Lokale"], df_premises)

    print(f"[OK] Saved to: {output_path}")
    print(f"     Dzialki:  {len(df_plots)} rows")
    print(f"     Budynki:  {len(df_buildings)} rows")
    print(f"     Lokale:   {len(df_premises)} rows")


# ============================================================================
# Entry point
# ============================================================================

def main():
    # --- Argument handling ---
    # Usage: python gml_reader.py <file.gml>
    # Output file name = GML file name with .xlsx extension
    if len(sys.argv) < 2:
        print("Usage: python gml_reader.py <file.gml>")
        print("Example: python gml_reader.py dane.gml")
        print("         Output: dane.xlsx")
        print("         TERC.xml is auto-detected in the same directory.")
        sys.exit(1)

    gml_path = sys.argv[1]

    # Output name = same as input with .xlsx extension
    output_name = Path(gml_path).stem + ".xlsx"

    # Auto-detect TERC.xml next to the GML file or in the current directory
    terc_path = None
    candidates = [
        Path(gml_path).parent / "TERC.xml",
        Path("TERC.xml"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            terc_path = str(candidate)
            break

    if not Path(gml_path).is_file():
        print(f"[ERROR] File not found: {gml_path}")
        sys.exit(1)

    print(f"[>] Parsing file: {gml_path}")

    # --- Step 0: Load TERC dictionary (optional) ---
    terc_dict = {}
    if terc_path:
        print(f"   Loading TERC dictionary from: {terc_path}")
        terc_dict = build_terc_dict(terc_path)
        print(f"     -> Loaded {len(terc_dict)} municipality entries")
    else:
        print("   [INFO] TERC.xml not found -- using raw TERYT codes")

    print("   Loading Precinct dictionary from: Obręby.dbf")
    precinct_dict = build_precinct_dict("Obręby.dbf")
    print(f"     -> Loaded {len(precinct_dict)} precinct entries")

    # --- Step 1: Detect XML namespaces ---
    print("   Detecting XML namespaces...")
    ns = detect_namespaces(gml_path)

    if "rcn" not in ns:
        print("[ERROR] Could not find an RCiWN namespace (rcn/rcwn) in the file.")
        print(f"   Detected namespaces: {ns}")
        sys.exit(1)

    print(f"   Namespaces: rcn={ns.get('rcn', '?')}")

    # --- Step 2: Parse XML tree ---
    print("   Loading XML tree...")
    tree = ET.parse(gml_path)
    root = tree.getroot()

    # --- Step 3: Build reference lookup dictionaries ---
    print("   Building document dictionary...")
    doc_dict = build_document_dict(root, ns)
    print(f"     -> Found {len(doc_dict)} documents")

    print("   Building address dictionary...")
    addr_dict = build_address_dict(root, ns)
    print(f"     -> Found {len(addr_dict)} addresses")

    print("   Building plot dictionary...")
    plot_dict = build_plot_dict(root, ns)
    print(f"     -> Found {len(plot_dict)} plots")

    print("   Building property dictionary...")
    prop_dict = build_property_dict(root, ns)
    print(f"     -> Found {len(prop_dict)} properties")

    print("   Building building dictionary...")
    bld_dict = build_building_dict(root, ns)
    print(f"     -> Found {len(bld_dict)} buildings")

    print("   Building premises dictionary...")
    prem_dict = build_premises_dict(root, ns)
    print(f"     -> Found {len(prem_dict)} premises")

    # --- Step 4: Process transactions (split into 3 categories) ---
    print("   Processing transactions...")
    plot_rows, building_rows, premises_rows = process_transactions(
        root, ns, doc_dict, addr_dict, plot_dict, prop_dict,
        bld_dict, prem_dict, terc_dict, precinct_dict
    )
    total = len(plot_rows) + len(building_rows) + len(premises_rows)
    print(f"     -> {total} transactions total")
    print(f"        Dzialki:  {len(plot_rows)}")
    print(f"        Budynki:  {len(building_rows)}")
    print(f"        Lokale:   {len(premises_rows)}")

    # --- Step 5: Export to Excel ---
    print(f"   Exporting to: {output_name}")
    export_to_excel(plot_rows, building_rows, premises_rows, output_name)

    print("[DONE] Finished.")


if __name__ == "__main__":
    main()
