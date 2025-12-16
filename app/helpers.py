from datetime import datetime, time, date
from typing import List, Union
import pandas as pd


DEFAULT_DATE_FORMAT = "%Y-%m-%d"

DATE_FORMATS = [
    # ISO
    "%Y-%m-%d",  # 2025-11-08
    "%Y.%m.%d",  # 2025.11.08
    "%Y/%m/%d",  # 2025/11/08
    # European / UK numeric
    "%d.%m.%Y",  # 11.08.2025
    "%d.%m.%y",  # 08.11.25
    "%m/%d/%Y",  # 11/08/2025
    "%d-%m-%y",  # 08-11-25
    "%d-%m-%Y",  # 08-11-2025
    "%d/%m/%y",  # 08/11/25
    "%d/%m/%Y",  # 08/11/2025
    # Short month name
    "%d %b %Y",  # 08 Nov 2025
    "%d-%b-%Y",  # 08-Nov-2025
    "%d/%b/%Y",  # 08/Nov/2025
    "%d %b %y",  # 08 Nov 25
    "%d-%b-%y",  # 08-Nov-25
    "%d/%b/%y",  # 08/Nov/25
    # Full month name
    "%d %B %Y",  # 08 November 2025
    "%d-%B-%Y",  # 08-November-2025
    "%d/%B/%Y",  # 08/November/2025
    "%d %B %y",  # 08 November 25
    "%d-%B-%y",  # 08-November-25
    "%d/%B/%y",  # 08/November/25
    # US
    "%m-%d-%Y",   # 11-08-2025
    "%m-%d-%y",   # 11-08-25
    "%b %d %Y",   # Nov 08 2025
    "%B %d %Y",   # November 08 2025
    # US numeric
    "%m-%d-%Y",
    "%m-%d-%y",
    "%m/%d/%Y",
    "%m/%d/%y",
    # US Full month name
    "%B %d %Y",
    "%B %d %y",
]


def get_date_format(list_values: List[str]) -> str:
    for value in list_values:
        if isinstance(value, str):
            for fmt in DATE_FORMATS:
                try:
                    datetime.strptime(value, fmt).strftime("%Y-%m-%d")
                    return fmt
                except ValueError:
                    pass
    return DEFAULT_DATE_FORMAT


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
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
    try:
        return str(datetime.fromisoformat(value).date())
    except:
        pass
    return value


def normalize_date(dt: Union[str, datetime], date_format: str):
    if isinstance(dt, str):
        dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
    return dt.strftime(date_format)


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