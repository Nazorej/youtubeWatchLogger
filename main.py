import os
from datetime import datetime, date, time, timedelta

from openpyxl import Workbook, load_workbook

from excelFormatter import applyFullStyle
from ytService import fetchVideoData


FILE_PATH = os.path.join(os.getcwd(), "youtube.xlsx")

DATE_FMT = "dd.mm.yy"     # как дата отображается в Excel
DUR_FMT = "[h]:mm:ss"     # формат длительности: часы не обнуляются после 24


def asDate(value) -> date | None:
    """Значение ячейки -> datetime.date (строки старого файла тоже понимаем)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        return datetime.strptime(value, "%d.%m.%y").date()
    return None


def asSeconds(value) -> int | None:
    """Значение ячейки -> секунды (timedelta / time / datetime-эпоха / строка)."""
    if value is None:
        return None
    if isinstance(value, timedelta):
        return int(value.total_seconds())
    if isinstance(value, time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, datetime):
        # длительности > 24 ч openpyxl читает как datetime от 1899-12-30
        return int((value - datetime(1899, 12, 30)).total_seconds())
    if isinstance(value, str):
        h, m, s = map(int, value.split(":"))
        return h * 3600 + m * 60 + s
    return None


def openOrCreateWorkbook():
    if os.path.exists(FILE_PATH):
        return load_workbook(FILE_PATH)

    wb = Workbook()

    ws = wb.active
    ws.title = "Просмотры"
    ws.append(["Дата", "Время", "Автор", "ID", "Название"])

    wb.create_sheet("Статистика по дням")

    return wb


def isDuplicate(ws, videoId: str) -> bool:
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[3] == videoId:
            return True
    return False


def appendVideo(ws, data: dict):
    row = ws.max_row + 1
    ws.append([
        data["date"],
        data["duration"],
        data["author"],
        data["video_id"],
        data["title"]
    ])
    # без явного формата длительность отобразится числом дней, а дата — числом
    ws.cell(row=row, column=1).number_format = DATE_FMT
    ws.cell(row=row, column=2).number_format = DUR_FMT


def updateStatsSheet(wb):
    ws = wb["Просмотры"]
    statsWs = wb["Статистика по дням"]

    stats: dict[date, int] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        d = asDate(row[0])
        seconds = asSeconds(row[1])

        if not d or seconds is None:
            continue

        stats[d] = stats.get(d, 0) + seconds

    # очищаем, но оставляем заголовок
    statsWs.delete_rows(1, statsWs.max_row)
    statsWs.append(["Дата", "Просмотрено (HH:MM:SS)"])

    r = 2
    for d in sorted(stats):   # date-объекты сортируются сами, strptime не нужен
        statsWs.cell(row=r, column=1, value=d).number_format = DATE_FMT
        statsWs.cell(row=r, column=2, value=timedelta(seconds=stats[d])).number_format = DUR_FMT
        r += 1


def processUrl(url: str):
    data = fetchVideoData(url)
    if not data:
        print("❌ Не удалось получить данные")
        return

    wb = openOrCreateWorkbook()
    ws = wb["Просмотры"]

    if isDuplicate(ws, data["video_id"]):
        print("⚠ Уже есть:", data["title"])
        return

    appendVideo(ws, data)
    updateStatsSheet(wb)

    # применяем стиль
    applyFullStyle(ws)
    applyFullStyle(wb["Статистика по дням"])

    wb.save(FILE_PATH)

    print("✔ Добавлено:", data["title"])


if __name__ == "__main__":
    processUrl("https://youtu.be/WOvzyVwxN9M")
