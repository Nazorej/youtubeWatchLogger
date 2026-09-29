import os
from datetime import datetime, timedelta

import yt_dlp

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COOKIES_FILE = os.path.join(BASE_DIR, "cookies.txt")

# Можно читать куки прямо из браузера вместо файла:
# "vivaldi", "chrome", "edge", "firefox" и т.д.
# Тогда браузер при запуске должен быть ПОЛНОСТЬЮ закрыт.
# Пустая строка = использовать cookies.txt из папки проекта (браузер может быть открыт).
COOKIES_BROWSER = ""


def fetchVideoData(url: str):
    ydl_opts = {
        "quiet": True,            # без баннеров и прогресса в консоль
        "no_warnings": True,      # предупреждения (PO Token и т.п.) не критичны для метаданных
        "noplaylist": True,       # ссылка с list= не должна тянуть весь плейлист
        "skip_download": True,    # нужны только метаданные, без скачивания видео
        "remote_components": ["ejs:github"],
    }

    if COOKIES_BROWSER:
        ydl_opts["cookiesfrombrowser"] = (COOKIES_BROWSER,)
    elif os.path.exists(COOKIES_FILE):
        ydl_opts["cookiefile"] = COOKIES_FILE

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
    except Exception as e:
        print("Ошибка yt-dlp")
        print(f"Детали ошибки:\n{type(e).__name__}: {e}")
        return None

    if not info:
        return None

    if "entries" in info:
        info = next(iter(info["entries"] or []), None)
        if not info:
            return None

    duration = info.get("duration")
    if not duration:
        return None

    return {
        # Настоящие date/timedelta объекты, а не строки —
        # openpyxl запишет их как реальные значения даты/времени в Excel.
        "date": datetime.now().date(),
        "duration": timedelta(seconds=duration),
        "author": info.get("uploader", ""),
        "video_id": info.get("id", ""),
        "title": info.get("title", "")
    }
