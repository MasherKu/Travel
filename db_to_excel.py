"""SQLite -> Excel: beach_destinations.db -> beach_destinations.xlsx
Запуск: python db_to_excel.py (из папки travel)
"""
import sqlite3
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "beach_destinations.db"
XLSX_PATH = BASE / "beach_destinations.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")

def write_sheet(ws, headers, rows):
    ws.append(headers)
    for c in ws[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for r in rows:
        ws.append(list(r))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    widths = {}
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row):
            widths[i] = max(widths.get(i, 0), len(str(v)) if v is not None else 0)
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i + 1)].width = min(max(w + 2, 12), 60)

def main():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    dest_headers = ["id", "country", "city", "visa_free", "visa_days",
                    "visa_note", "direct_flight", "direct_note",
                    "climate_profile", "best_months"]
    dest_rows = cur.execute("""SELECT id, country, city, visa_free_for_russians,
        visa_free_duration_days, visa_note, direct_flights_from_moscow,
        direct_flights_note, climate_profile, best_months
        FROM destinations ORDER BY country, city""").fetchall()
    mon_headers = ["destination_id", "country", "city", "month", "month_num",
                   "rain_mm", "wind_kmh", "water_c",
                   "air_morning", "air_day", "air_evening", "air_night"]
    mon_rows = cur.execute("""SELECT m.destination_id, d.country, d.city, m.month, m.month_num,
        m.rain_mm, m.wind_speed_kmh, m.water_temp_c,
        m.air_morning, m.air_day, m.air_evening, m.air_night
        FROM monthly_data m JOIN destinations d ON d.id = m.destination_id
        ORDER BY d.country, d.city, m.month_num""").fetchall()
    meta_rows = cur.execute("SELECT key, value FROM metadata").fetchall()
    con.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Destinations"
    write_sheet(ws1, dest_headers, dest_rows)
    ws2 = wb.create_sheet("Monthly")
    write_sheet(ws2, mon_headers, mon_rows)
    ws3 = wb.create_sheet("Metadata")
    write_sheet(ws3, ["key", "value"], meta_rows)

    if XLSX_PATH.exists():
        XLSX_PATH.unlink()
    wb.save(XLSX_PATH)
    print(f"OK: {len(dest_rows)} направлений, {len(mon_rows)} строк -> {XLSX_PATH.name}")

if __name__ == "__main__":
    main()
