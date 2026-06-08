"""
Лаунчер РискНавигатор.

Запускает Streamlit-приложение (app.py) в текущем процессе и открывает браузер.
Подходит как для запуска из исходников, так и для сборки автономного .exe
(PyInstaller). В отличие от запуска через subprocess, этот способ корректно
работает внутри «замороженного» приложения.
"""
import os
import sys
import socket


def resolve_path(rel_path: str) -> str:
    """Путь к ресурсу как из исходников, так и из распакованного .exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


def get_free_port(default: int = 8501) -> int:
    """Свободный TCP-порт (чтобы не конфликтовать с занятым 8501)."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("", 0))
            return s.getsockname()[1]
    except OSError:
        return default


def silence_first_run_prompt():
    """
    Streamlit при первом запуске спрашивает e-mail в консоли и блокирует старт.
    Создаём пустые учётные данные, чтобы запрос не появлялся (важно для .exe).
    """
    cred_dir = os.path.join(os.path.expanduser("~"), ".streamlit")
    cred_file = os.path.join(cred_dir, "credentials.toml")
    try:
        os.makedirs(cred_dir, exist_ok=True)
        if not os.path.exists(cred_file):
            with open(cred_file, "w", encoding="utf-8") as f:
                f.write('[general]\nemail = ""\n')
    except OSError:
        pass


def main():
    silence_first_run_prompt()
    app_path = resolve_path("app.py")

    # В «замороженном» режиме переходим в каталог рядом с .exe,
    # чтобы база данных (risk_management.db) создавалась в доступном месте,
    # а не во временной папке распаковки.
    if getattr(sys, "frozen", False):
        run_dir = os.path.dirname(sys.executable)
        try:
            os.chdir(run_dir)
        except OSError:
            pass

    port = get_free_port()

    sys.argv = [
        "streamlit", "run", app_path,
        f"--server.port={port}",
        "--server.headless=false",          # false -> браузер откроется сам
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
        "--server.fileWatcherType=none",
        # Тема задаётся флагами, чтобы оформление работало и без config.toml
        "--theme.base=light",
        "--theme.primaryColor=#2F6FED",
        "--theme.backgroundColor=#F4F6FB",
        "--theme.secondaryBackgroundColor=#FFFFFF",
        "--theme.textColor=#16243A",
    ]

    print("=" * 52)
    print("   РискНавигатор")
    print(f"   Приложение запускается на http://localhost:{port}")
    print("   Для остановки закройте это окно или нажмите Ctrl+C")
    print("=" * 52)

    from streamlit.web import cli as stcli
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
