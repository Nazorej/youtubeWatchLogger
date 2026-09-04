import subprocess
import json
from datetime import datetime, timedelta


def fetchVideoData(url: str):
    # Добавили --cookies-from-browser (используем firefox, так как он проще отдает куки без закрытия браузера)
    # Если используете Chrome/Edge, замените "firefox" на "chrome" (при этом браузер должен быть закрыт)
    result = subprocess.run(
        ["yt-dlp", "-J", "--remote-components", "ejs:github", url],
        capture_output=True
    )

    if result.returncode != 0:
        print("Ошибка yt-dlp")
        error_msg = result.stderr.decode("utf-8", errors="replace")
        print(f"Детали ошибки:\n{error_msg}")
        return None

    stdout_text = result.stdout.decode("utf-8", errors="replace")

    try:
        info = json.loads(stdout_text)
    except json.JSONDecodeError:
        print("Ошибка JSON")
        return None

    if "entries" in info:
        info = info["entries"][0]

    duration = info.get("duration")
    if not duration:
        return None

    return {
        # Возвращаем настоящие date/timedelta объекты, а не строки —
        # тогда openpyxl запишет их как реальные значения даты/времени в Excel.
        "date": datetime.now().date(),
        "duration": timedelta(seconds=duration),
        "author": info.get("uploader", ""),
        "video_id": info.get("id", ""),
        "title": info.get("title", "")
    }
