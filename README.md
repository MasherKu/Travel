# Пляжные направления: сезонность, визы, перелёты

Датасет из **200 пляжных направлений** со среднемесячными климатическими нормами,
визовыми правилами для граждан России (актуальность — октябрь 2026) и наличием
прямых перелётов из Москвы. Плюс готовый конвейер: JSON → SQLite → Excel.

## Состав репозитория

| Файл | Что это |
|---|---|
| `beach_destinations.json` | Основной датасет (источник истины) |
| `json_to_database.py` | JSON → SQLite (`beach_destinations.db`) |
| `db_to_excel.py` | SQLite → Excel (`beach_destinations.xlsx`, 3 листа: Destinations / Monthly / Metadata) |
| `make_season_excel.py` | JSON → `beach_season_by_month.xlsx` — таблица «месяц → где пляжный сезон», готовая к печати |
| `requirements.txt` | Зависимости (openpyxl) |

Файлы `*.db`, `*.xlsx`, `*.ods` — генерируемые артефакты, в репозиторий не включаются
(см. `.gitignore`), пересоздаются скриптами.

## Схема данных

Каждая запись в `destinations[]`:

```
country, city                     — страна, курорт
visa_free_for_russians            — true/false (для граждан РФ)
visa_free_duration_days           — срок безвиза в днях (null = виза нужна / не требуется)
visa_note                         — примечание по визе
direct_flights_from_moscow        — есть ли прямые рейсы из Москвы
direct_flights_note               — примечание по рейсам
climate_profile                   — климатический профиль направления
seasonality.best_months           — лучшие месяцы для пляжного отдыха (4–6)
seasonality.monthly_data          — 12 месяцев (january…december):
    rain_mm                       — среднемесячные осадки, мм
    wind_speed_kmh                — средняя скорость ветра, км/ч
    water_temp_c                  — температура воды, °C
    air_temp.morning/day/evening/night — воздух утром/днём/вечером/ночью, °C
```

## Использование

```bash
pip install -r requirements.txt

python json_to_database.py    # пересоздаёт beach_destinations.db
python db_to_excel.py         # пересоздаёт beach_destinations.xlsx
python make_season_excel.py   # пересоздаёт beach_season_by_month.xlsx
```

## Категории в «Сезон по месяцам»

- **Идеально** — месяц из `best_months` направления, осадки < 150 мм и ветер < 25 км/ч.
- **Тоже неплохо** — не в `best_months`, но: воздух днём ≥ 25 °C, вода ≥ 24 °C, осадки ≤ 80 мм.
- **Дожди** — сезонный месяц, но осадки ≥ 150 мм или ветер ≥ 25 км/ч (заметная волна).

Итог по датасету: 922 строки «Идеально», 179 «Тоже неплохо», 49 «Дожди».

Лист «Сезон по месяцам» свёрстан для печати: альбомный A4, вписывание по ширине,
повтор шапки на каждой странице, колонтитулы с номерами страниц.

## Источники и дисклеймер

Климат — среднемноголетние ориентировочные нормы из открытых климатологических
данных (climatestotravel, weather2travel, weather-and-climate, meteoblue,
seatemperature и др.), не прогноз погоды. Визы и прямые рейсы — статус на
октябрь 2026 (АТОР, Консульский департамент МИД РФ, расписания авиакомпаний),
проверяйте перед покупкой билетов.
