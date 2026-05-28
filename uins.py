from datetime import datetime
import re


SPECIAL_PREFIX_MARKS = ("п", "g")


def _prefix_for(value, year):
    stripped = value.strip().lower()
    return f"12{year}20020" if stripped.startswith(SPECIAL_PREFIX_MARKS) else f"11{year}20020"


def _numeric_part(value):
    match = re.search(r"\d+", value)
    if not match:
        raise ValueError(f"Не найден номер УИН: {value}")
    return int(match.group(0))


def format_uins(value):
    if not value:
        return []

    year = datetime.now().year - 2000
    result = []
    parts = [part.strip() for part in re.split(r"[,;\n]+", value) if part.strip()]

    for part in parts:
        if "-" in part:
            start_raw, end_raw = [item.strip() for item in part.split("-", 1)]
            prefix = _prefix_for(start_raw, year)
            start = _numeric_part(start_raw)
            end = _numeric_part(end_raw)
            if end < start:
                start, end = end, start
            width = max(len(str(start)), len(str(end)))
            for number in range(start, end + 1):
                result.append(prefix + str(number).zfill(width))
            continue

        prefix = _prefix_for(part, year)
        result.append(prefix + str(_numeric_part(part)))

    return result
