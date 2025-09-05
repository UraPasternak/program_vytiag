import re
import pandas as pd
from PyPDF2 import PdfReader

# === Налаштування ===
pdf_path = "ДРРП_3.pdf"   # шлях до твого PDF
output_path = "vityag_output.xlsx"  # шлях куди зберегти Excel

# === Читаємо PDF ===
reader = PdfReader(pdf_path)
full_text = ""
for page in reader.pages:
    txt = page.extract_text()
    if txt:
        full_text += txt + "\n"

# === Парсимо на блоки ===
blocks = re.split(r"Актуальна інформація про об’єкт речових прав", full_text)
data = []

for block in blocks:
    if "Реєстраційний номер об’єкта" not in block:
        continue

    reg_num = re.search(r"Реєстраційний номер об’єкта\s*нерухомого майна:\s*(\d+)", block)
    reg_num = reg_num.group(1) if reg_num else ""

    ident = re.search(r"Ідентифікатор об’єкта в\s*ЄДЕССБ:\s*([^\s]+)", block)
    ident = ident.group(1) if ident else ""

    # Загальна площа
    total_area = re.search(r"Загальна площа \(кв\.м\):\s*([\d,\.]+)", block)
    total_area = total_area.group(1).replace(",", ".") if total_area else ""

    addr = re.search(r"Адреса:\s*(.+?)(?=Актуальна інформація|Номер відомостей)", block, re.S)
    addr = addr.group(1).replace("\n", " ").strip() if addr else ""

    rights = re.findall(r"Номер відомостей про речове право:\s*(\d+)", block)
    rights = ", ".join(rights) if rights else ""

    dates = re.findall(r"Дата, час державної реєстрації:\s*([\d\.]+\s*[\d:]+)", block)
    dates = ", ".join(dates) if dates else ""

    # === Витягуємо номер об’єкта ===
    obj_num = ""
    m_flat = re.search(r"квартира\s*(\d+)", addr)
    m_pm = re.search(r"П/М-\d+", addr)
    m_room = re.search(r"\bП\d+\b", addr)
    m_digits = re.search(r"\b\d{3,5}\b", addr)  # просто число (3–5 цифр)

    if m_flat:
        obj_num = m_flat.group(1)
    elif m_pm:
        obj_num = m_pm.group(0)
    elif m_room:
        obj_num = m_room.group(0)
    elif m_digits:
        obj_num = m_digits.group(0)

    data.append({
        "Реєстраційний номер": reg_num,
        "Ідентифікатор ЄДЕССБ": ident,
        "Адреса": addr,
        "Номер відомостей": rights,
        "Дата реєстрації": dates,
        "Загальна площа (кв.м)": total_area,
        "Номер об’єкта": obj_num
    })

# === Формуємо Excel ===
df = pd.DataFrame(data)
df.to_excel(output_path, index=False)

print(f"Готово! Дані збережені у {output_path}")
