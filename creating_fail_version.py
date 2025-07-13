import os
import main

version_str = main.version.replace(',', '.')  # '1, 0, 0, 1' -> '1.0.0.1'

version_info = f"""# UTF-8 encoding is required
VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=({main.version}),  # Версия файла
        prodvers=({main.version}),  # Версия продукта
        mask=0x3f,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0)
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    '040904b0',
                    [
                        StringStruct('CompanyName', 'OSI'),
                        StringStruct('FileDescription', 'Программа для быстрого заполнения ведомости'),
                        StringStruct('FileVersion', '{version_str}'),
                        StringStruct('InternalName', 'OSI'),
                        StringStruct('LegalCopyright', '© 2025 OSI'),
                        StringStruct('OriginalFilename', 'OSI.exe'),
                        StringStruct('ProductName', 'OSI'),
                        StringStruct('ProductVersion', '{version_str}')
                    ]
                )
            ]
        ),
        VarFileInfo([VarStruct('Translation', [1049, 1200])])
    ]
)
"""

with open("version.txt", "w", encoding="utf-8") as file:
    file.write(version_info)

print("Файл version.txt успешно создан!")
