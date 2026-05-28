from datetime import datetime, timedelta


MONTH_NAMES = {
    1: "января",
    2: "февраля",
    3: "марта",
    4: "апреля",
    5: "мая",
    6: "июня",
    7: "июля",
    8: "августа",
    9: "сентября",
    10: "октября",
    11: "ноября",
    12: "декабря",
}


def format_ru_date(date_value, quoted=False, suffix=" г."):
    day = f"{date_value.day:02d}"
    if quoted:
        return f"«{day}» {MONTH_NAMES[date_value.month]} {date_value.year}{suffix}"
    return f"{day} {MONTH_NAMES[date_value.month]} {date_value.year}{suffix}"


def next_work_date(east=False, now=None):
    now = now or datetime.now()
    is_friday = now.isoweekday() == 5
    if is_friday and east:
        return now + timedelta(days=4)
    if east and now.isoweekday() == 4:
        return now + timedelta(days=4)
    if is_friday:
        return now + timedelta(days=3)
    if east:
        return now + timedelta(days=2)
    return now + timedelta(days=1)
