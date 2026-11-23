from __future__ import annotations

import re
import unicodedata
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from collections import defaultdict
from copy import deepcopy
from functools import lru_cache
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, redirect, render_template, request, url_for
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data" / "abastecimiento"
RESERVOIR_FILE = DATA_DIR / "BD_ EMBALSE PALO VERDE.xlsx"
OFFERS_FILE = DATA_DIR / "BD_ OFERTAS.xlsx"
RESERVOIR_LEVELS_FILE = DATA_DIR / "BD_NIVELES DE RESERVORIOS 1.xlsx"
WATER_QUALITY_FILE = DATA_DIR / "BD_ Formato de calidad de agua de tratamiento.xlsx"

MONTH_NAMES = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
MONTH_INDEX = {name.casefold(): number for number, name in enumerate(MONTH_NAMES, start=1)}
AREAS = [
    {"slug": "abastecimiento", "name": "Abastecimiento", "description": "Tratamiento de agua y continuidad del suministro para la operación del fundo.", "detail": "Gestión del tratamiento de agua, disponibilidad y calidad para la operación agrícola.", "icon": "water", "status": "Reporte disponible"},
    {"slug": "monitoreo", "name": "Monitoreo", "description": "Seguimiento de indicadores y condiciones operativas del fundo.", "detail": "Espacio para integrar indicadores, alertas y seguimiento operativo.", "icon": "monitor", "status": "En preparación"},
    {"slug": "electromecanica", "name": "Electromecánica", "description": "Estado de equipos, mantenimiento y servicios electromecánicos.", "detail": "Espacio para integrar mantenimiento, disponibilidad de equipos y servicios.", "icon": "gear", "status": "En preparación"},
    {"slug": "proyectos", "name": "Proyectos", "description": "Avance, hitos e inversión de los proyectos del fundo.", "detail": "Espacio para integrar avances, hitos, responsables e inversión de proyectos.", "icon": "projects", "status": "Reporte externo"},
    {"slug": "infraestructura", "name": "Infraestructura", "description": "Condición y necesidades de la infraestructura de soporte.", "detail": "Espacio para integrar estado, intervenciones y necesidades de infraestructura.", "icon": "building", "status": "En preparación"},
]


class ReportSourceError(RuntimeError):
    """Raised when an Excel source cannot provide the requested report."""


def _number(value: Any) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    number = float(value)
    return number if number == number and abs(number) != float("inf") else None


def _weeknum_monday(value: date) -> int:
    january_first = date(value.year, 1, 1)
    return ((value - january_first).days + january_first.weekday()) // 7 + 1

def _headers(row: tuple[Any, ...]) -> dict[str, int]:
    return {" ".join(str(value).strip().casefold().split()): index for index, value in enumerate(row) if value is not None}


def _open_sheet(path: Path, sheet_name: str):
    if not path.is_file():
        raise ReportSourceError(f"No se encontró el archivo fuente: {path.name}")
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as error:
        raise ReportSourceError(f"No se pudo abrir {path.name}.") from error
    if sheet_name not in workbook.sheetnames:
        workbook.close()
        raise ReportSourceError(f"No se encontró la hoja {sheet_name} en {path.name}.")
    return workbook, workbook[sheet_name]


def _excel_table_last_row(path: Path, table_name: str) -> int:
    try:
        with ZipFile(path) as archive:
            for member in archive.namelist():
                if not member.startswith("xl/tables/table") or not member.endswith(".xml"):
                    continue
                table = ET.fromstring(archive.read(member))
                if table.attrib.get("name") == table_name:
                    reference = table.attrib.get("ref")
                    if not reference:
                        break
                    return range_boundaries(reference)[3]
    except Exception as error:
        raise ReportSourceError(f"No se pudo leer el rango de la tabla {table_name} en {path.name}.") from error
    raise ReportSourceError(f"No se encontro la tabla {table_name} en {path.name}.")


def _load_reservoir_matrix(requested_year: int | None = None) -> dict[str, Any]:
    workbook, sheet = _open_sheet(RESERVOIR_FILE, "H2OMOS")
    try:
        columns = _headers(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        required = ("week", "year", "date", "flow transferred", "inflow", "reservoir capacity (hm3)")
        missing = [name for name in required if name not in columns]
        if missing:
            raise ReportSourceError("Faltan columnas requeridas en H2OMOS: " + ", ".join(missing))
        daily: dict[date, dict[str, Any]] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            raw_date = row[columns["date"]]
            if isinstance(raw_date, datetime):
                record_date = raw_date.date()
            elif isinstance(raw_date, date):
                record_date = raw_date
            else:
                continue
            raw_year = _number(row[columns["year"]])
            record_year = int(raw_year) if raw_year is not None else record_date.year
            transferred = _number(row[columns["flow transferred"]])
            inflow = _number(row[columns["inflow"]])
            capacity = _number(row[columns["reservoir capacity (hm3)"]])
            raw_week = _number(row[columns["week"]])
            week_number = int(raw_week) if raw_week is not None else _weeknum_monday(record_date)
            reserve_capacity = {2026: 8.15, 2025: 11.4, 2024: 14.5, 2023: 14.5, 2022: 14.5, 2021: 17.8, 2020: 17.8, 2019: 19.5, 2018: 19.5}.get(record_year)
            capacity_percent = capacity / reserve_capacity if capacity is not None and reserve_capacity else None
            if transferred is None or transferred < 0:
                flow_band = ""
            elif transferred < 10.5:
                flow_band = "flow-low"
            elif transferred < 15:
                flow_band = "flow-license"
            elif transferred > 15 and transferred < 100:
                flow_band = "flow-free"
            else:
                flow_band = ""
            has_data = any(value is not None for value in (transferred, inflow, capacity))
            daily[record_date] = {
                "day": record_date.day,
                "year": record_year,
                "week": week_number,
                "has_data": has_data,
                "flow_transferred": transferred,
                "flow_band": flow_band,
                "inflow": inflow,
                "capacity": capacity,
                "capacity_percent": capacity_percent,
            }
    except ReportSourceError:
        raise
    except Exception as error:
        raise ReportSourceError("No se pudo leer la matriz diaria de H2OMOS.") from error
    finally:
        workbook.close()

    years = sorted({record["year"] for record in daily.values()})
    if not years:
        raise ReportSourceError("H2OMOS no contiene filas con fecha.")
    selected_year = requested_year if requested_year in years else years[-1]
    matrix_months = []
    for month_number in range(1, 13):
        rows = [daily[item] for item in sorted(daily) if daily[item]["year"] == selected_year and item.month == month_number]
        if rows:
            matrix_months.append({"number": month_number, "name": MONTH_NAMES[month_number - 1], "days": rows})
    observed_dates = [item for item, record in daily.items() if record["year"] == selected_year and record["has_data"]]
    return {
        "year": selected_year,
        "years": years,
        "months": matrix_months,
        "row_count": len(observed_dates),
        "latest_date": max(observed_dates).strftime("%d/%m/%Y") if observed_dates else "—",
        "source_name": RESERVOIR_FILE.name,
        "sheet_name": "H2OMOS",
    }

@lru_cache(maxsize=18)
def _cached_reservoir_matrix(requested_year: int | None, modified_ns: int, file_size: int) -> dict[str, Any]:
    return _load_reservoir_matrix(requested_year)


def load_reservoir_matrix(requested_year: int | None = None) -> dict[str, Any]:
    if not RESERVOIR_FILE.is_file():
        raise ReportSourceError(f"No se encontró el archivo fuente: {RESERVOIR_FILE.name}")
    stat = RESERVOIR_FILE.stat()
    return deepcopy(_cached_reservoir_matrix(requested_year, stat.st_mtime_ns, stat.st_size))

def _load_license_usage_matrix(requested_year: int | None = None) -> dict[str, Any]:
    workbook, sheet = _open_sheet(OFFERS_FILE, "EXE lice")
    try:
        columns = _headers(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        required = ("mes", "fundo", "licencia (m3)", "ejecucion de licencia (%)")
        missing = [name for name in required if name not in columns]
        if missing:
            raise ReportSourceError("Faltan columnas requeridas en EXE lice: " + ", ".join(missing))
        funds: set[str] = set()
        grouped: dict[tuple[int, float], dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
        for row in sheet.iter_rows(min_row=2, values_only=True):
            month_name = str(row[columns["mes"]] or "").strip()
            fund = str(row[columns["fundo"]] or "").strip()
            license_value = _number(row[columns["licencia (m3)"]])
            execution = _number(row[columns["ejecucion de licencia (%)"]])
            if license_value is None or not month_name or not fund:
                continue
            month_number = int(_number(row[columns["n° mes"]]) or 0) if "n° mes" in columns else 0
            if not 1 <= month_number <= 12:
                month_number = MONTH_INDEX.get(month_name.casefold(), 0)
            if not 1 <= month_number <= 12:
                continue
            funds.add(fund)
            key = (month_number, round(license_value, 2))
            if execution is not None:
                grouped[key][fund].append(execution)
    except ReportSourceError:
        raise
    except Exception as error:
        raise ReportSourceError("No se pudo leer la matriz WATER LICENSE USAGE de EXE lice.") from error
    finally:
        workbook.close()

    if not funds:
        raise ReportSourceError("EXE lice no contiene datos de licencia por mes y fundo.")
    sorted_funds = sorted(funds)
    months = []
    for month_number in range(1, 13):
        entries = []
        for key in sorted((key for key in grouped if key[0] == month_number), key=lambda item: item[1]):
            averages = {}
            for fund in sorted_funds:
                values = grouped[key].get(fund, [])
                averages[fund] = sum(values) / len(values) if values else None
            entries.append({"license": key[1], "averages": averages})
        if entries:
            months.append({"number": month_number, "name": MONTH_NAMES[month_number - 1], "groups": entries})
    return {
        "funds": sorted_funds,
        "months": months,
        "row_count": sum(len(month["groups"]) for month in months),
        "source_name": OFFERS_FILE.name,
        "sheet_name": "EXE lice",
    }

@lru_cache(maxsize=18)
def _cached_license_usage_matrix(requested_year: int | None, modified_ns: int, file_size: int) -> dict[str, Any]:
    return _load_license_usage_matrix(requested_year)


def load_license_usage_matrix(requested_year: int | None = None) -> dict[str, Any]:
    if not OFFERS_FILE.is_file():
        raise ReportSourceError(f"No se encontró el archivo fuente: {OFFERS_FILE.name}")
    stat = OFFERS_FILE.stat()
    return deepcopy(_cached_license_usage_matrix(requested_year, stat.st_mtime_ns, stat.st_size))

def _load_reservoir_levels_data() -> dict[str, Any]:
    last_row = _excel_table_last_row(RESERVOIR_LEVELS_FILE, "Tabla4")
    workbook, sheet = _open_sheet(RESERVOIR_LEVELS_FILE, "DATA")
    try:
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, max_col=11, values_only=True), ())
        columns = {}
        for index, value in enumerate(header_row):
            if value is None:
                continue
            key = unicodedata.normalize("NFKD", str(value).strip()).encode("ascii", "ignore").decode("ascii").casefold()
            columns[" ".join(key.split())] = index
        required = ("fecha", "semana", "ano", "fundo", "reservorio", "capacidad total (m3)", "volumen (m3)")
        missing = [name for name in required if name not in columns]
        if missing:
            raise ReportSourceError("Faltan columnas requeridas en DATA: " + ", ".join(missing))
        grouped: dict[tuple[int, int, date, str, str], list[Any]] = {}
        funds: set[str] = set()
        for row in sheet.iter_rows(min_row=2, max_row=last_row, max_col=11, values_only=True):
            raw_date = row[columns["fecha"]]
            if isinstance(raw_date, datetime):
                record_date = raw_date.date()
            elif isinstance(raw_date, date):
                record_date = raw_date
            else:
                continue
            raw_year = _number(row[columns["ano"]])
            record_year = int(raw_year) if raw_year is not None else record_date.year
            raw_week = _number(row[columns["semana"]])
            record_week = int(raw_week) if raw_week is not None else _weeknum_monday(record_date)
            fund = str(row[columns["fundo"]] or "").strip()
            reservoir = str(row[columns["reservorio"]] or "").strip()
            if not fund or not reservoir:
                continue
            capacity = _number(row[columns["capacidad total (m3)"]])
            volume = _number(row[columns["volumen (m3)"]])
            key = (record_year, record_week, record_date, fund, reservoir)
            values = grouped.setdefault(key, [0.0, 0.0, False, False])
            if capacity is not None:
                values[0] += capacity
                values[2] = True
            if volume is not None:
                values[1] += volume
                values[3] = True
            funds.add(fund)
    except ReportSourceError:
        raise
    except Exception as error:
        raise ReportSourceError("No se pudo leer la tabla DATA de niveles de reservorios.") from error
    finally:
        workbook.close()
    if not grouped:
        raise ReportSourceError("DATA no contiene registros de reservorios con fecha.")
    return {"groups": grouped, "funds": sorted(funds)}


@lru_cache(maxsize=4)
def _cached_reservoir_levels_data(modified_ns: int, file_size: int) -> dict[str, Any]:
    return _load_reservoir_levels_data()


def load_reservoir_levels_matrix(
    selected_year: int,
    selected_month: int | None,
    selected_day: int | None,
    selected_week: int | None,
) -> dict[str, Any]:
    if not RESERVOIR_LEVELS_FILE.is_file():
        raise ReportSourceError(f"No se encontro el archivo fuente: {RESERVOIR_LEVELS_FILE.name}")
    stat = RESERVOIR_LEVELS_FILE.stat()
    source = _cached_reservoir_levels_data(stat.st_mtime_ns, stat.st_size)
    matching_dates = [
        record_date
        for year, record_week, record_date, fund, reservoir in source["groups"]
        if year == selected_year
        and (selected_month is None or record_date.month == selected_month)
        and (selected_day is None or record_date.day == selected_day)
        and (selected_week is None or record_week == selected_week)
    ]
    latest_date = max(matching_dates) if matching_dates else None
    totals: dict[str, dict[str, list[Any]]] = defaultdict(dict)
    for (year, record_week, record_date, fund, reservoir), values in source["groups"].items():
        if latest_date is None or year != selected_year or record_date != latest_date:
            continue
        if selected_week is not None and record_week != selected_week:
            continue
        totals[reservoir].setdefault(fund, [0.0, 0.0, False, False])
        aggregate = totals[reservoir][fund]
        aggregate[0] += values[0]
        aggregate[1] += values[1]
        aggregate[2] = aggregate[2] or values[2]
        aggregate[3] = aggregate[3] or values[3]
    rows = []
    for reservoir in sorted(totals, reverse=True):
        fund_values = {}
        for fund in source["funds"]:
            values = totals[reservoir].get(fund)
            if values is None:
                fund_values[fund] = {"capacity": None, "volume": None, "level": 0.0}
                continue
            capacity = values[0] if values[2] else None
            volume = values[1] if values[3] else None
            level = volume / capacity if volume is not None and capacity not in (None, 0) else 0.0
            fund_values[fund] = {"capacity": capacity, "volume": volume, "level": level}
        rows.append({"reservoir": reservoir, "funds": fund_values, "is_total": False})
    if rows:
        accumulated = {}
        for fund in source["funds"]:
            fund_rows = [row["funds"][fund] for row in rows]
            capacities = [item["capacity"] for item in fund_rows if item["capacity"] is not None]
            volumes = [item["volume"] for item in fund_rows if item["volume"] is not None]
            capacity = sum(capacities) if capacities else None
            volume = sum(volumes) if volumes else None
            level = volume / capacity if volume is not None and capacity not in (None, 0) else 0.0
            accumulated[fund] = {"capacity": capacity, "volume": volume, "level": level}
        rows.append({"reservoir": "ACCUMULATED", "funds": accumulated, "is_total": True})
    return {"funds": source["funds"], "rows": rows}


def _load_water_quality_data() -> dict[str, Any]:
    last_row = _excel_table_last_row(WATER_QUALITY_FILE, "DATA1")
    workbook, sheet = _open_sheet(WATER_QUALITY_FILE, "DATA")
    try:
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, max_col=39, values_only=True), ())
        columns = {}
        for index, value in enumerate(header_row):
            if value is None:
                continue
            key = unicodedata.normalize("NFKD", str(value).strip()).encode("ascii", "ignore").decode("ascii").casefold()
            columns[" ".join(key.split())] = index
        required = ("fecha", "semana", "ano", "reservorio", "ntu hidrante", "ntu df", "ph df", "ce df")
        missing = [name for name in required if name not in columns]
        if missing:
            raise ReportSourceError("Faltan columnas requeridas en calidad de agua: " + ", ".join(missing))
        measure_columns = {
            "hydrante_ntu": columns["ntu hidrante"],
            "leaked_ntu": columns["ntu df"],
            "leaked_ph": columns["ph df"],
            "leaked_ce": columns["ce df"],
        }
        grouped: dict[tuple[int, int, date, int], dict[str, list[float]]] = {}
        for row in sheet.iter_rows(min_row=2, max_row=last_row, max_col=39, values_only=True):
            raw_date = row[columns["fecha"]]
            if isinstance(raw_date, datetime):
                record_date = raw_date.date()
            elif isinstance(raw_date, date):
                record_date = raw_date
            else:
                continue
            raw_year = _number(row[columns["ano"]])
            record_year = int(raw_year) if raw_year is not None else record_date.year
            raw_week = _number(row[columns["semana"]])
            record_week = int(raw_week) if raw_week is not None else _weeknum_monday(record_date)
            raw_block = _number(row[columns["reservorio"]])
            if raw_block is None or int(raw_block) not in (1, 2, 3):
                continue
            block_number = int(raw_block)
            key = (record_year, record_week, record_date, block_number)
            values = grouped.setdefault(key, {name: [0.0, 0] for name in measure_columns})
            for name, column in measure_columns.items():
                value = _number(row[column])
                if value is not None and value != 0:
                    values[name][0] += value
                    values[name][1] += 1
    except ReportSourceError:
        raise
    except Exception as error:
        raise ReportSourceError("No se pudo leer la tabla DATA de calidad de agua.") from error
    finally:
        workbook.close()
    if not grouped:
        raise ReportSourceError("DATA de calidad no contiene registros para los bloques 1, 2 o 3.")
    return {"groups": grouped, "blocks": (1, 2, 3), "measures": ("hydrante_ntu", "leaked_ntu", "leaked_ph", "leaked_ce")}


@lru_cache(maxsize=4)
def _cached_water_quality_data(modified_ns: int, file_size: int) -> dict[str, Any]:
    return _load_water_quality_data()


def load_water_quality_matrix(
    selected_year: int,
    selected_month: int | None,
    selected_day: int | None,
    selected_week: int | None,
) -> dict[str, Any]:
    if not WATER_QUALITY_FILE.is_file():
        raise ReportSourceError(f"No se encontro el archivo fuente: {WATER_QUALITY_FILE.name}")
    stat = WATER_QUALITY_FILE.stat()
    source = _cached_water_quality_data(stat.st_mtime_ns, stat.st_size)
    matching_dates = [
        record_date
        for year, record_week, record_date, block in source["groups"]
        if year == selected_year
        and (selected_month is None or record_date.month == selected_month)
        and (selected_day is None or record_date.day == selected_day)
        and (selected_week is None or record_week == selected_week)
    ]
    latest_date = max(matching_dates) if matching_dates else None
    totals = {name: [0.0, 0] for name in source["measures"]}
    block_totals = {block: {name: [0.0, 0] for name in source["measures"]} for block in source["blocks"]}
    if latest_date is not None:
        for (year, record_week, record_date, block), values in source["groups"].items():
            if year != selected_year or record_date != latest_date:
                continue
            if selected_week is not None and record_week != selected_week:
                continue
            for name in source["measures"]:
                block_totals[block][name][0] += values[name][0]
                block_totals[block][name][1] += values[name][1]
                totals[name][0] += values[name][0]
                totals[name][1] += values[name][1]
    rows = []
    if latest_date is not None:
        for block in source["blocks"]:
            averages = {
                name: values[0] / values[1] if values[1] else None
                for name, values in block_totals[block].items()
            }
            rows.append({"block": block, "averages": averages, "is_total": False})
        averages = {name: values[0] / values[1] if values[1] else None for name, values in totals.items()}
        rows.append({"block": "AVERAGE", "averages": averages, "is_total": True})
    return {"rows": rows, "as_of_date": latest_date.strftime("%d/%m/%Y") if latest_date else None}


def load_abastecimiento_report(requested_year: int | None = None, requested_month: int | None = None, requested_day: int | None = None, requested_week: int | None = None) -> dict[str, Any]:
    reservoir = load_reservoir_matrix(requested_year)
    license_usage = load_license_usage_matrix(0)
    selected_year = reservoir["year"]
    month_numbers = [month["number"] for month in reservoir["months"]]
    if not month_numbers:
        raise ReportSourceError(f"No hay datos disponibles para {selected_year}.")
    today = date.today()
    default_month = today.month if today.month in month_numbers else month_numbers[-1]
    if requested_month == 0:
        selected_month = None
    elif requested_month is None:
        selected_month = default_month
    else:
        selected_month = requested_month if requested_month in month_numbers else default_month
    month_options = [{"number": number, "name": MONTH_NAMES[number - 1]} for number in month_numbers]
    if selected_month is not None:
        reservoir["months"] = [month for month in reservoir["months"] if month["number"] == selected_month]

    current_week = _weeknum_monday(today)
    week_options = sorted({row["week"] for month in reservoir["months"] for row in month["days"]})
    default_week = current_week if current_week in week_options else (week_options[-1] if week_options else None)
    if requested_week == 0:
        selected_week = None
    elif requested_week is None:
        selected_week = default_week
    else:
        selected_week = requested_week if requested_week in week_options else default_week
    if selected_week is not None:
        filtered_months = []
        for month in reservoir["months"]:
            days = [row for row in month["days"] if row["week"] == selected_week]
            if days:
                filtered_months.append({**month, "days": days})
        reservoir["months"] = filtered_months
    else:
        filtered_months = []
        for month in reservoir["months"]:
            days = [row for row in month["days"] if row["has_data"]]
            if days:
                filtered_months.append({**month, "days": days})
        reservoir["months"] = filtered_months

    day_options = sorted({row["day"] for month in reservoir["months"] for row in month["days"]})
    selected_day = requested_day if requested_day in day_options else None
    if selected_day is not None:
        reservoir["months"] = [
            {**month, "days": [row for row in month["days"] if row["day"] == selected_day]}
            for month in reservoir["months"]
            if any(row["day"] == selected_day for row in month["days"])
        ]
    reservoir["row_count"] = sum(len(month["days"]) for month in reservoir["months"])
    reservoir_levels = load_reservoir_levels_matrix(selected_year, selected_month, selected_day, selected_week)
    water_quality = load_water_quality_matrix(selected_year, selected_month, selected_day, selected_week)
    return {
        "year": selected_year,
        "years": reservoir["years"],
        "months": month_options,
        "selected_month": selected_month,
        "selected_day": selected_day,
        "day_options": day_options,
        "selected_week": selected_week,
        "week_options": week_options,
        "week_number": current_week,
        "reservoir": reservoir,
        "license_usage": license_usage,
        "reservoir_levels": reservoir_levels,
        "water_quality": water_quality,
    }

@app.get("/")
def home():
    return render_template("index.html", areas=AREAS, year=date.today().year)


@app.get("/api/areas")
def areas_api():
    return jsonify({"areas": AREAS})


@app.get("/abastecimiento")
def abastecimiento_report():
    selected_year = request.args.get("anio", type=int) or date.today().year
    month_arg = request.args.get("mes")
    selected_month = 0 if month_arg == "" else request.args.get("mes", type=int)
    day_arg = request.args.get("dia")
    selected_day = 0 if day_arg == "" else request.args.get("dia", type=int)
    week_arg = request.args.get("semana")
    selected_week = 0 if week_arg == "" else request.args.get("semana", type=int)
    try:
        report = load_abastecimiento_report(selected_year, selected_month, selected_day, selected_week)
    except ReportSourceError as error:
        return render_template("reports/abastecimiento.html", report=None, error=str(error)), 503
    return render_template("reports/abastecimiento.html", report=report, error=None)


@app.get("/api/abastecimiento")
def abastecimiento_api():
    selected_year = request.args.get("anio", type=int) or date.today().year
    month_arg = request.args.get("mes")
    selected_month = 0 if month_arg == "" else request.args.get("mes", type=int)
    day_arg = request.args.get("dia")
    selected_day = 0 if day_arg == "" else request.args.get("dia", type=int)
    week_arg = request.args.get("semana")
    selected_week = 0 if week_arg == "" else request.args.get("semana", type=int)
    try:
        report = load_abastecimiento_report(selected_year, selected_month, selected_day, selected_week)
    except ReportSourceError as error:
        return jsonify({"error": str(error)}), 503
    return jsonify(report)

@app.get("/abastecimiento/otro-reporte")
def abastecimiento_extra_report():
    return render_template("reports/coming_soon.html")


@app.get("/area/<slug>")
def area_detail(slug: str):
    area = next((item for item in AREAS if item["slug"] == slug), None)
    if area is None:
        abort(404)
    if slug == "abastecimiento":
        return redirect(url_for("abastecimiento_report"))
    return render_template("area.html", area=area)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)










