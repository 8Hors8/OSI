import subprocess
import main

# 1. Сначала генерируем version.txt
subprocess.run(["python", "creating_fail_version.py"], check=True)

# 2. Затем вызываем pyinstaller
subprocess.run([
    "pyinstaller",
    "--onefile",
    "--noconsole",
    "--version-file=version.txt",  # Убедись, что путь правильный!
    "--icon=лого.ico",                    # ОБЯЗАТЕЛЬНАЯ запятая здесь
    f"--name=Помощник ОСИ v{main.version}",                         # Исправлено с --n NAME=ОСИ
    "main.py"
], check=True)
