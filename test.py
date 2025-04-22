import openpyxl
from docx import Document
import re

# Открываем .docx
doc = Document(r"C:\Users\User\Documents\диплом\Тестовый.docx")

# Создаем новую Excel-книгу
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Данные из DOCX"

# Заголовки
ws.append(["Наименование", "Идентификатор", "Номер 16р слова", "Начальный участок выдачи", "Конечный участок выдачи", "Тип", "Задача/Блок", "Количество переменных"])

# Берем первую таблицу
table = doc.tables[0]

# Массивы
words = []
used_words = []
rows_to_process = []

number_word = 51  # Начинаем отсчет слов

# --- Первый проход: собираем строки из Word и сохраняем их в Excel без слов ---
for i, row in enumerate(table.rows):
    if i == 0:
        continue

    cells = [cell.text.strip() for cell in row.cells]
    if len(cells) < 5:
        continue

    name = cells[0]
    id = cells[1]

    numbers = re.findall(r"\d+", cells[3])
    if len(numbers) == 3:
        start_num = numbers[1]
        end_num = numbers[2]
    else:
        start_num = numbers[0]
        end_num = numbers[1]

    count_match = re.findall(r"^(\d+)", cells[-2])
    count = int(count_match[0]) if count_match else 1

    tip_match = re.findall(r"/(\d+)", cells[-2])
    tip = tip_match[0] if tip_match else ""

    if tip:
        tip = tip[0]
    else:
        tip = ""

    block = cells[-1]

    if tip == "1":
        words_to_add = 1
    elif tip == "2":
        words_to_add = 2
    elif tip == "3":
        words_to_add = 1
    elif tip == "4":
        words_to_add = 2
    elif tip == "5":
        words_to_add = 2
    elif tip == "6":
        words_to_add = 4
    elif tip == "7":
        words_to_add = 1
    else:
        words_to_add = 1

    # Сохраняем строку для второго прохода
    excel_row = ws.max_row + 1
    rows_to_process.append({
        "name": name,
        "id": id,
        "start_num": start_num,
        "end_num": end_num,
        "tip": tip,
        "block": block,
        "count": count,
        "words_to_add": words_to_add,
        "excel_row": excel_row
    })

    # Записываем строку в Excel (без слова пока)
    ws.append([name, id, "", start_num, end_num, tip, block, count])

# --- Второй проход: выдача слов ---

# Сначала приоритетные переменные
priority_rows = []
other_rows = []

for row in rows_to_process:
    start = int(row["start_num"])
    end = int(row["end_num"])
    if 100 <= start <= 120 and 235 <= end <= 255:
        priority_rows.append(row)
    else:
        other_rows.append(row)

other_rows.sort(key=lambda r: int(r["start_num"]))
all_rows_sorted = priority_rows + other_rows

# Выдаём слова
for row in all_rows_sorted:
    name = row["name"]
    id = row["id"]
    start_num = row["start_num"]
    end_num = row["end_num"]
    tip = row["tip"]
    block = row["block"]
    count = row["count"]
    words_to_add = row["words_to_add"]
    excel_row = row["excel_row"]

    first_word = number_word

    for _ in range(words_to_add * count):
        words.append({
            "номер": str(number_word),
            "начальный": start_num,
            "конечный": end_num
        })
        used_words.append({
            "номер": str(number_word),
            "начальный": start_num,
            "конечный": end_num
        })
        number_word += 1

    last_word = number_word - 1
    word_range = f"{first_word}" if first_word == last_word else f"{first_word}-{last_word}"

    # Обновляем строку в Excel — столбец 3 (номер 16р слова)
    ws.cell(row=excel_row, column=3).value = word_range

# --- Третий проход: повторное использование слов ---

for row in rows_to_process:
    start_num = row["start_num"]
    end_num = row["end_num"]
    excel_row = row["excel_row"]
    words_to_add = row["words_to_add"] * row["count"]
    assigned_words = []

    for _ in range(words_to_add):
        for word in used_words:
            if int(word["конечный"]) < int(start_num):
                # Обновляем слово
                word["конечный"] = end_num

                # Добавляем повторно
                for w in words:
                    if w["номер"] == word["номер"]:
                        w["конечный"] = end_num
                        break
                assigned_words.append(word["номер"])
                break

    if assigned_words:
        word_range = assigned_words[0] if len(assigned_words) == 1 else f"{assigned_words[0]}-{assigned_words[-1]}"
        # Обновляем в Excel
        ws.cell(row=excel_row, column=3).value = word_range

# --- Сохраняем Excel ---
wb.save(r"C:\Users\User\Documents\диплом\результат.xlsx")
print("Файл сохранен!")

# --- Печать массива слов ---
print("\nМассив слов:")
for word in words:
    print(f"Номер: {word['номер']}, Начальный участок: {word['начальный']}, Конечный участок: {word['конечный']}")

