"""JSON -> SQLite: beach_destinations.json -> beach_destinations.db
Запуск: python json_to_database.py (работает из папки travel)
"""
import json
import sqlite3
from pathlib import Path

BASE = Path(__file__).resolve().parent
JSON_PATH = BASE / "beach_destinations.json"
DB_PATH = BASE / "beach_destinations.db"

MONTH_ORDER = ["january","february","march","april","may","june","july",
               "august","september","october","november","december"]

def main():
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    destinations = data["destinations"]
    metadata = data.get("metadata", {})

    if DB_PATH.exists():
        DB_PATH.unlink()  # пересоздаём базу

    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    cur.execute("""CREATE TABLE destinations(
        id INTEGER PRIMARY KEY,
        country TEXT NOT NULL,
        city TEXT NOT NULL,
        visa_free_for_russians INTEGER NOT NULL,
        visa_free_duration_days INTEGER NOT NULL,
        visa_note TEXT,
        direct_flights_from_moscow INTEGER NOT NULL,
        direct_flights_note TEXT,
        climate_profile TEXT,
        best_months TEXT
    )""")
    cur.execute("""CREATE TABLE monthly_data(
        id INTEGER PRIMARY KEY,
        destination_id INTEGER NOT NULL REFERENCES destinations(id),
        month TEXT NOT NULL,
        month_num INTEGER NOT NULL,
        rain_mm REAL,
        wind_speed_kmh REAL,
        water_temp_c REAL,
        air_morning REAL,
        air_day REAL,
        air_evening REAL,
        air_night REAL
    )""")
    cur.execute("CREATE INDEX idx_month_dest ON monthly_data(destination_id, month_num)")
    cur.execute("CREATE INDEX idx_dest_country ON destinations(country, city)")
    cur.execute("""CREATE TABLE metadata(
        key TEXT PRIMARY KEY,
        value TEXT
    )""")

    for k, v in metadata.items():
        cur.execute("INSERT INTO metadata(key,value) VALUES(?,?)",
                    (k, json.dumps(v, ensure_ascii=False) if isinstance(v,(list,dict)) else str(v)))

    for i, d in enumerate(destinations, start=1):
        cur.execute("""INSERT INTO destinations
            (id,country,city,visa_free_for_russians,visa_free_duration_days,
             visa_note,direct_flights_from_moscow,direct_flights_note,
             climate_profile,best_months)
            VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (i, d["country"], d["city"],
             int(bool(d["visa_free_for_russians"])), int(d["visa_free_duration_days"]),
             d.get("visa_note",""), int(bool(d["direct_flights_from_moscow"])),
             d.get("direct_flights_note",""), d.get("climate_profile",""),
             ",".join(d["seasonality"]["best_months"])))
        md = d["seasonality"]["monthly_data"]
        for n, m in enumerate(MONTH_ORDER, start=1):
            r = md[m]
            a = r["air_temp"]
            cur.execute("""INSERT INTO monthly_data
                (destination_id,month,month_num,rain_mm,wind_speed_kmh,water_temp_c,
                 air_morning,air_day,air_evening,air_night)
                VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (i, m, n, r["rain_mm"], r["wind_speed_kmh"], r["water_temp_c"],
                 a["morning"], a["day"], a["evening"], a["night"]))

    con.commit()
    nd = cur.execute("SELECT COUNT(*) FROM destinations").fetchone()[0]
    nm = cur.execute("SELECT COUNT(*) FROM monthly_data").fetchone()[0]
    print(f"OK: {nd} destinations, {nm} monthly rows -> {DB_PATH.name}")
    # пример запроса
    row = cur.execute("""SELECT country, city FROM destinations
                         WHERE visa_free_for_russians=1 AND direct_flights_from_moscow=1
                         LIMIT 5""").fetchall()
    print("Пример (безвиз + прямой рейс):", row)
    con.close()

if __name__ == "__main__":
    main()
