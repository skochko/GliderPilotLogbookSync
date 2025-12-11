from datetime import datetime, time, date
from typing import Union


def normalize_flight_time(value) -> str:
    # Если это datetime — берём time
    if isinstance(value, datetime):
        return value.strftime("%H:%M")

    # Если это time — форматируем сразу
    if isinstance(value, time):
        return value.strftime("%H:%M")

    # Если это строка — пытаемся распарсить
    if isinstance(value, str):
        # Частые форматы времени
        time_formats = [
            "%H:%M",
            "%H:%M:%S",
            "%I:%M %p",  # 3:52 PM
            "%Y-%m-%d %H:%M:%S",  # Excel-like full datetime
            "%Y-%m-%d %H:%M",
        ]

        for fmt in time_formats:
            try:
                return datetime.strptime(value, fmt).strftime("%H:%M")
            except ValueError:
                pass

        # Последняя попытка: пусть datetime сам попробует (ISO)
        try:
            return datetime.fromisoformat(value).strftime("%H:%M")
        except:
            pass

    return value


def normalize_flight_date(value) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, str):
        # Попробуем несколько частых форматов
        formats = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d.%m.%Y",
            "%m/%d/%Y",
            "%d-%m-%Y",
            "%Y.%m.%d",
            "%d %b %Y",
            "%d %B %Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass

    # Если ничего не подошло — последнее спасение
    try:
        return str(datetime.fromisoformat(value).date())
    except:
        pass

    return value


def normalize_date(dt: Union[str, datetime]):
    if isinstance(dt, str):
        dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
    return dt.strftime("%Y-%m-%d")


def normalize_time(dt: str):
    """
    1899-12-30 14:21:00
    """
    dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
    return dt.strftime("%H:%M")
