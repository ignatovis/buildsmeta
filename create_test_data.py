#!/usr/bin/env python3
"""
Генерация тестовых Excel-файлов для проверки checker.py.

Создаёт:
  test_classifier.xlsx  — 15 позиций, 3 раздела
  test_estimate.xlsx    — 7 строк (3 ОК, 3 ошибки, 1 уточнение)

Запуск:
  python create_test_data.py
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def header_style(cell):
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")
    cell.alignment = Alignment(horizontal="center", wrap_text=True)


def create_classifier(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Классификатор"

    headers = ["Раздел", "Статья", "Позиция", "Наименование работы", "Ед. изм."]
    ws.append(headers)
    for cell in ws[1]:
        header_style(cell)
    ws.row_dimensions[1].height = 36

    data = [
        ("1",   None,  None,    "Земляные работы",                              None),
        (None, "1.1",  None,    "Разработка грунта",                            None),
        (None,  None, "1.1.1",  "Разработка грунта экскаватором",               "м³"),
        (None,  None, "1.1.2",  "Ручная разработка грунта",                     "м³"),
        (None, "1.2",  None,    "Обратная засыпка",                             None),
        (None,  None, "1.2.1",  "Обратная засыпка с уплотнением",               "м³"),
        ("2",   None,  None,    "Бетонные и железобетонные работы",             None),
        (None, "2.1",  None,    "Устройство фундаментов",                       None),
        (None,  None, "2.1.1",  "Монолитный ленточный фундамент",               "м³"),
        (None,  None, "2.1.2",  "Свайный фундамент",                            "шт"),
        (None, "2.2",  None,    "Устройство перекрытий",                        None),
        (None,  None, "2.2.1",  "Монолитное перекрытие",                        "м²"),
        ("3",   None,  None,    "Металлоконструкции",                           None),
        (None, "3.1",  None,    "Монтаж несущих конструкций",                   None),
        (None,  None, "3.1.1",  "Монтаж металлоконструкций каркаса здания",     "т"),
    ]

    for row in data:
        ws.append(row)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 46
    ws.column_dimensions["E"].width = 12

    wb.save(path)
    print(f"  Классификатор : {path}  ({len(data)} строк)")


def create_estimate(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Смета"

    headers = [
        "Кодификатор работы", "Номер", "Наименование работы", "Единицы измерения",
        "Объём", "Ставка за материал", "Ставка за работу", "Итог ставка+работа",
        "Общая сумма за материал", "Общая сумма за работу", "Итог общий",
        "Примечание Генерального подрядчика", "Примечание заказчика",
    ]
    ws.append(headers)
    for cell in ws[1]:
        header_style(cell)
    ws.row_dimensions[1].height = 50

    data = [
        ("1.1.1", 1, "Разработка грунта экскаватором",         "м³",   100, 500,  1200, 1700,  50000,  120000,  170000, None, None),
        ("1.1.2", 2, "Ручная разработка грунта",               "тонн",  50,   0,   800,  800,      0,   40000,   40000, None, None),
        ("1.2.1", 3, "Обратная засыпка",                       "м³",    80,   0,   600,  600,      0,   48000,   48000, None, None),
        ("2.1.1", 4, "Монолитный ленточный фундамент",         "м³",   200, 3000, 2000, 5000, 600000,  400000, 1000000, None, None),
        ("2.2.1", 5, "Монолитное перекрытие",                  "м²",   500,  800, 1500, 2300, 400000,  750000, 1150000, None, None),
        ("3.9.9", 6, "Кровельные работы (неизвестный код)",    "м²",   300,  200,  500,  700,  60000,  150000,  210000, None, None),
        ("2.1",   7, "Устройство фундаментов",                 "м³",   150, 2000, 1800, 3800, 300000,  270000,  570000, None, None),
    ]

    for row in data:
        ws.append(row)

    col_widths = [22, 8, 42, 18, 10, 18, 16, 18, 22, 20, 14, 28, 22]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w

    wb.save(path)
    print(f"  Смета         : {path}  ({len(data)} строк)")


if __name__ == "__main__":
    clf_path = os.path.join(SCRIPT_DIR, "test_classifier.xlsx")
    est_path = os.path.join(SCRIPT_DIR, "test_estimate.xlsx")

    print("Создание тестовых файлов...")
    create_classifier(clf_path)
    create_estimate(est_path)
    print("Готово.")
    print()
    print("Ожидаемые результаты при проверке:")
    print("  Строка 2 (1.1.1) : ОК")
    print("  Строка 3 (1.1.2) : ОШИБКА  — Ед. изм. не соответствует (тонн вместо м3)")
    print("  Строка 4 (1.2.1) : ОШИБКА  — Наименование не соответствует")
    print("  Строка 5 (2.1.1) : ОК")
    print("  Строка 6 (2.2.1) : ОК")
    print("  Строка 7 (3.9.9) : ОШИБКА  — Код не найден")
    print("  Строка 8 (2.1)   : УТОЧНЕНИЕ — Требуется уточнение детализации")
