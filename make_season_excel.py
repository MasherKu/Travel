# -*- coding: utf-8 -*-
"""beach_destinations.json -> beach_season_by_month.xlsx
Таблица «месяц -> где пляжный сезон» с тремя категориями и настройкой печати.
Запуск: python make_season_excel.py (из папки travel)
"""
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

BASE = Path(__file__).resolve().parent
SRC = BASE / "beach_destinations.json"
OUT = BASE / "beach_season_by_month.xlsx"

MONTHS = ["january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"]
MONTH_RU = {"january": "Январь", "february": "Февраль", "march": "Март", "april": "Апрель",
            "may": "Май", "june": "Июнь", "july": "Июль", "august": "Август",
            "september": "Сентябрь", "october": "Октябрь", "november": "Ноябрь", "december": "Декабрь"}

MIN_DAY, MIN_WATER, MAX_RAIN = 25.0, 24.0, 80.0   # категория «Тоже неплохо»
RAIN_WET, WIND_STRONG = 150.0, 25.0               # категория «Дожди» (осадки/волны)


def visa_text(x):
    if x["visa_free_for_russians"]:
        d = x["visa_free_duration_days"]
        return d if d is not None else "не требуется"
    return "виза нужна"


def collect():
    data = json.loads(SRC.read_text(encoding="utf-8"))
    by_month = {m: {"ideal": [], "ok": [], "wet": []} for m in MONTHS}
    for x in data["destinations"]:
        md = x["seasonality"]["monthly_data"]
        best = {b.lower() for b in x["seasonality"]["best_months"]}
        for m in MONTHS:
            v = md[m]
            rec = {"country": x["country"], "city": x["city"], "visa": visa_text(x),
                   "visa_note": x.get("visa_note", ""),
                   "direct": "да" if x["direct_flights_from_moscow"] else "нет",
                   "direct_note": x.get("direct_flights_note", ""),
                   "day": v["air_temp"]["day"], "water": v["water_temp_c"],
                   "rain": v["rain_mm"], "wind": v["wind_speed_kmh"]}
            in_best = m in best
            ok_enough = rec["day"] >= MIN_DAY and rec["water"] >= MIN_WATER and rec["rain"] <= MAX_RAIN
            if not (in_best or ok_enough):
                continue
            if rec["rain"] >= RAIN_WET or rec["wind"] >= WIND_STRONG:
                by_month[m]["wet"].append(rec)
            elif in_best:
                by_month[m]["ideal"].append(rec)
            else:
                by_month[m]["ok"].append(rec)
    for m in MONTHS:
        for k in ("ideal", "ok", "wet"):
            by_month[m][k].sort(key=lambda r: (r["country"], r["city"]))
    return data, by_month


HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
IDEAL_FILL = PatternFill("solid", fgColor="C6EFCE")   # зелёный
OK_FILL = PatternFill("solid", fgColor="FFEB9C")      # жёлтый
WET_FILL = PatternFill("solid", fgColor="B4C6E7")     # голубой (дожди/волны)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
WRAP = Alignment(wrap_text=True, vertical="top")


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, THIN
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)


def setup_print(ws, orientation, fit_height, title, titles_rows=None):
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = 9  # A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    m = ws.page_margins
    m.left, m.right, m.top, m.bottom = 0.4, 0.4, 0.6, 0.6
    m.header, m.footer = 0.3, 0.3
    ws.print_options.horizontalCentered = True
    ws.print_options.gridLines = False
    if titles_rows:
        ws.print_title_rows = titles_rows
    ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.oddHeader.center.text = title
    ws.oddHeader.center.size = 9
    ws.oddFooter.center.text = "стр. &P из &N"
    ws.oddFooter.center.size = 8


def main():
    data, by_month = collect()

    wb = Workbook()

    # Лист 1: построчно месяц-направление
    ws1 = wb.active
    ws1.title = "Сезон по месяцам"
    cols1 = ["Месяц", "Категория", "Страна", "Город", "Безвиз для россиян (дней)", "Прямой рейс из Москвы",
             "Воздух днём °C", "Вода °C", "Осадки мм", "Ветер км/ч", "Примечание (виза)", "Примечание (рейсы)"]
    ws1.append(cols1)
    style_header(ws1, len(cols1))
    r = 2
    for m in MONTHS:
        for cat, fill, lst in (("Идеально", IDEAL_FILL, by_month[m]["ideal"]),
                               ("Тоже неплохо", OK_FILL, by_month[m]["ok"]),
                               ("Дожди", WET_FILL, by_month[m]["wet"])):
            for rec in lst:
                vals = [MONTH_RU[m], cat, rec["country"], rec["city"], rec["visa"], rec["direct"],
                        rec["day"], rec["water"], rec["rain"], rec["wind"],
                        rec["visa_note"], rec["direct_note"]]
                for c, v in enumerate(vals, 1):
                    ws1.cell(row=r, column=c, value=v)
                if isinstance(rec["visa"], int):
                    ws1.cell(row=r, column=5).number_format = "0"
                for c in range(1, len(cols1) + 1):
                    cell = ws1.cell(row=r, column=c)
                    cell.fill, cell.border = fill, THIN
                r += 1
    for i, w in enumerate([11, 13, 16, 24, 24, 22, 14, 10, 11, 11, 42, 42], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:{get_column_letter(len(cols1))}{r - 1}"

    # Лист 2: сводка по месяцам
    ws2 = wb.create_sheet("Сводка по месяцам")
    cols2 = ["Месяц", "Идеально (шт)", "Идеально — направления", "Тоже неплохо (шт)",
             "Тоже неплохо — направления", "Дожди (шт)", "Дожди — направления"]
    ws2.append(cols2)
    style_header(ws2, len(cols2))
    r = 2
    names = lambda lst: "; ".join(f"{x['country']}, {x['city']}" for x in lst)
    for m in MONTHS:
        ideal, ok, wet = by_month[m]["ideal"], by_month[m]["ok"], by_month[m]["wet"]
        for c, v in enumerate([MONTH_RU[m], len(ideal), names(ideal), len(ok), names(ok),
                               len(wet), names(wet)], 1):
            ws2.cell(row=r, column=c, value=v)
        for c in range(1, len(cols2) + 1):
            cell = ws2.cell(row=r, column=c)
            cell.border, cell.alignment = THIN, WRAP
        ws2.row_dimensions[r].height = min(409, max(30, 13 * (1 + max(len(ideal), len(ok), len(wet)) // 4)))
        r += 1
    for i, w in enumerate([10, 13, 85, 13, 85, 11, 55], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    # Лист 3: правила
    ws3 = wb.create_sheet("Правила и источник")
    rules = [
        ["Файл-источник", "beach_destinations.json — 200 направлений (generated 2026-10-15)"],
        ["Категория «Идеально»", "Месяц входит в seasonality.best_months (4–6 месяцев) И осадки < 150 мм И ветер < 25 км/ч"],
        ["Категория «Дожди»", f"Сезонный месяц (best_months или тёплый), но осадки ≥ {RAIN_WET:.0f} мм ИЛИ ветер ≥ {WIND_STRONG:.0f} км/ч (волны)"],
        ["Категория «Тоже неплохо»", f"Не в best_months, но: воздух днём ≥ {MIN_DAY:.0f}°C, вода ≥ {MIN_WATER:.0f}°C, осадки ≤ {MAX_RAIN:.0f} мм"],
        ["Порог 150 мм", "Справочно: 0–30 мм — сухо, до 80 — приемлемо, 150+ — выраженный дождливый период даже для тропиков"],
        ["Порог ветер 25 км/ч", "Заметная волна на море; в датасете максимум 24–25 км/ч (Красное море, кайтспоты — сентябрь/май)"],
        ["Безвиз (дней)", "visa_free_for_russians + visa_free_duration_days; «виза нужна» = false/null; «не требуется» = null при true"],
        ["Прямой рейс", "direct_flights_from_moscow из исходного файла (статус на октябрь 2026)"],
        ["Примечания", "Столбцы «Примечание» — поля visa_note и direct_flights_note из исходного файла"],
        ["Дисклеймер источника", data.get("metadata", {}).get("disclaimer", "")],
    ]
    for i, (k, v) in enumerate(rules, 1):
        ws3.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws3.cell(row=i, column=2, value=v).alignment = WRAP
    ws3.column_dimensions["A"].width = 26
    ws3.column_dimensions["B"].width = 110

    setup_print(ws1, "landscape", 0, "Пляжный сезон по месяцам — 200 направлений", "1:1")
    setup_print(ws2, "landscape", 0, "Сводка: пляжный сезон по месяцам")
    setup_print(ws3, "portrait", 1, "Правила формирования таблицы")

    if OUT.exists():
        OUT.unlink()
    wb.save(OUT)

    ti = sum(len(by_month[m]["ideal"]) for m in MONTHS)
    to = sum(len(by_month[m]["ok"]) for m in MONTHS)
    tw = sum(len(by_month[m]["wet"]) for m in MONTHS)
    print(f"OK: {ti} идеально, {to} тоже неплохо, {tw} дожди -> {OUT.name}")


if __name__ == "__main__":
    main()
