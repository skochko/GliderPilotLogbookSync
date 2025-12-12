from datetime import datetime, time, date
from typing import Union
import pandas as pd


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


def normalize_time(dt):
    """
    Converts Access/Excel datetime/Timestamp/string to HH:MM
    """
    if dt is None or pd.isna(dt):
        return None

    # If it's already datetime or Timestamp
    if isinstance(dt, (datetime, pd.Timestamp)):
        return dt.strftime("%H:%M")

    # If it's a string → try to parse
    if isinstance(dt, str):
        # List of acceptable formats
        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%d.%m.%Y %H:%M",
            "%H:%M",
        ]
        for fmt in formats:
            try:
                parsed = datetime.strptime(dt, fmt)
                return parsed.strftime("%H:%M")
            except ValueError:
                pass
        raise ValueError(f"Unrecognized time string: {dt}")

    # Unexpected type
    raise TypeError(f"normalize_time(): unsupported type {type(dt)} with value {dt}")