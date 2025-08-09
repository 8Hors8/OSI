import subprocess
import main

# 1. Сначала генерируем version.txt
subprocess.run(["python", "creating_fail_version.py"], check=True)

# 2. Затем вызываем pyinstaller
subprocess.run([
    "pyinstaller",
    "--onefile",
    "--noconsole",
    "--version-file=version.txt",                                   # указать путь к фалу
    "--icon=лого.ico",                                              # указать иконку
    f"--name=Помощник ОСИ v{main.version}",
    "main.py"
], check=True)
