import streamlit as st
import os
import re
import json
import uuid
import pdfplumber
import urllib3
import requests
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Отключаем предупреждения
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= НАСТРОЙКИ =================
# Вставь сюда свой Authorization Key от GigaChat
CREDENTIALS = """MDFhMGNhODItZjljMy03ZWJkLWJmMzktM2E5NjZkZWQ4YmYyOjIyY2JmN2E2LWFkNDEtNGVkNi05NDkyLWI0YzIzMTdmNjBkOA=="""

FIELD_NAMES_RU = {
    'passport_number': 'Номер паспорта',
    'issued_to': 'Выдан',
    'issue_date': 'Дата выдачи',
    'manufacturing_date': 'Дата изготовления',
    'product_name': 'Наименование и марка изделия',
    'volume': 'Количество изделий',
    'gost': 'ГОСТ/ТУ',
    'concrete_class': 'Класс бетона',
    'strength_required': 'Требуемая прочность',
    'strength_release': 'Отпускная прочность',
    'strength_actual': 'Фактическая прочность',
    'frost_resistance': 'Морозостойкость',
    'water_absorption': 'Водопоглощение бетона по массе %',
    'product_weight': 'Масса изделия кг',
    'reinforcement': 'Арматура',
    'surface_category': 'Категория поверхности',
    'dimensions': 'Размеры изделия',
    'accuracy_class': 'Класс точности',
    'rebar_weight': 'Вес арматурного каркаса',
    'product_type': 'Тип продукции',
    'concrete_density': 'Средняя плотность бетона кг/см2',
    'concrete_humidity': 'Отпускная влажность бетона %',
    'standard_designation': 'Обозначение стандарта (ГОСТ, ТУ)',
    'series_number': 'Номер серии и выпуска рабочих чертежей',
    'radiation_compliance': 'Соответствие по радионуклидам'
}
# =============================================

# Создаем папку для временных файлов
TEMP_DIR = Path("temp_web_files")
TEMP_DIR.mkdir(exist_ok=True)

@st.cache_resource
def get_gigachat_token():
    """Получает и кэширует токен GigaChat"""
    url = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
    credentials = CREDENTIALS.strip()
    credentials = credentials.encode('ascii', 'ignore').decode('ascii')
    credentials = ''.join(c for c in credentials if c.isprintable() and not c.isspace())
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
        "Authorization": f"Basic {credentials}",
        "RqUID": str(uuid.uuid4())
    }
    
    try:
        response = requests.post(url, headers=headers, data="scope=GIGACHAT_API_PERS", verify=False, timeout=30)
        if response.status_code == 200:
            return response.json().get("access_token")
        return None
    except Exception:
        return None

def get_available_models(token: str):
    url = "https://api.giga.chat/v1/models"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=30)
        if response.status_code == 200:
            models = response.json().get("data", [])
            priority = ['GigaChat-3-Ultra', 'GigaChat-3-Pro', 'GigaChat-2-Max']
            for preferred in priority:
                for model in models:
                    if preferred in model.get('id', ''):
                        return model.get('id')
            return models[0].get('id') if models else "GigaChat"
        return "GigaChat"
    except Exception:
        return "GigaChat"

def call_gigachat_text(token: str, prompt: str, model_name: str) -> str:
    url = "https://api.giga.chat/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "stream": False
    }
    try:
        response = requests.post(url, headers=headers, json=payload, verify=False, timeout=120)
        if response.status_code == 200:
            return response.json()['choices'][0]['message']['content']
        return None
    except Exception:
        return None

def extract_text_from_pdf(pdf_path: Path) -> str:
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception:
        pass
    
    if len(text) < 50:
        try:
            import fitz
            doc = fitz.open(pdf_path)
            for page in doc:
                page_text = page.get_text()
                if page_text:
                    text += page_text + "\n"
            doc.close()
        except Exception:
            pass
    return text

def is_valid_text(text: str) -> bool:
    if len(text) < 50:
        return False
    cyrillic_and_digits = sum(1 for c in text if c.isalpha() and c.isascii() == False or c.isdigit())
    total_chars = sum(1 for c in text if not c.isspace())
    if total_chars == 0:
        return False
    return (cyrillic_and_digits / total_chars) > 0.3

def split_passports(text: str) -> list:
    patterns = [r'ТЕХНИЧЕСКИЙ ПАСПОРТ', r'Паспорт качества', r'ПАСПОРТ КАЧЕСТВА']
    positions = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            positions.append(match.start())
    positions.sort()
    
    if len(positions) == 0:
        return [text]
    
    passports = []
    for i, pos in enumerate(positions):
        if i + 1 < len(positions):
            passport_text = text[pos:positions[i+1]].strip()
        else:
            passport_text = text[pos:].strip()
        if len(passport_text) > 100:
            passports.append(passport_text)
    
    return passports if passports else [text]

def parse_with_regex(text: str) -> dict:
    data = {}
    match = re.search(r'(?:Паспорт качества|Паспорт|ТЕХНИЧЕСКИЙ ПАСПОРТ).*?№?\s*([A-Za-z0-9_\-/бн]+)', text, re.IGNORECASE)
    if match: data['passport_number'] = match.group(1).strip()
    match = re.search(r'Выдан\s+(.+?)(?:\n|Дата|$)', text, re.IGNORECASE)
    if match: data['issued_to'] = match.group(1).strip()
    match = re.search(r'(?:Дата выдачи|ОТ|от).*?(\d{2}\.\d{2}\.\d{4})', text)
    if match: data['issue_date'] = match.group(1)
    match = re.search(r'Дата изготовления.*?:\s*(.+?)(?:\n\d|$)', text, re.IGNORECASE)
    if match: data['manufacturing_date'] = match.group(1).strip()
    match = re.search(r'на изделия из.*?бетонов[:\s]*(.+?)(?:\n\d|$)', text, re.IGNORECASE)
    if match: data['product_type'] = match.group(1).strip()
    match = re.search(r'Наименование и марка изделия.*?:\s*(.+?)(?:\n\d|$)', text, re.IGNORECASE)
    if match: data['product_name'] = match.group(1).strip()
    match = re.search(r'(?:Количество|Объём|Партия|Кол-во).*?:\s*(\d+)\s*(?:шт|ед)', text, re.IGNORECASE)
    if match: data['volume'] = match.group(1).strip() + " шт."
    match = re.search(r'(?:ГОСТ|ТУ)\s*[:\-]?\s*([A-Za-z0-9\-\.]+)', text, re.IGNORECASE)
    if match: data['gost'] = match.group(1).strip()
    match = re.search(r'(?:Класс бетона|Марка бетона).*?(В\s*\d+(?:[.,]\d+)?)', text, re.IGNORECASE)
    if match: data['concrete_class'] = match.group(1).replace(',', '.').replace(' ', '')
    match = re.search(r'(?:Rт|Rтр|Требуемая прочность).*?(\d+(?:[.,]\d+)?)\s*МПа', text, re.IGNORECASE)
    if match: data['strength_required'] = match.group(1).replace(',', '.') + " МПа"
    match = re.search(r'(?:Rотп|Отпускная прочность).*?(\d+(?:[.,]\d+)?)\s*(?:%|МПа)', text, re.IGNORECASE)
    if match: data['strength_release'] = match.group(1).replace(',', '.') + "%"
    match = re.search(r'(?:Rф|Rфакт|Фактическая прочность).*?(\d+(?:[.,]\d+)?)\s*МПа', text, re.IGNORECASE)
    if match: data['strength_actual'] = match.group(1).replace(',', '.') + " МПа"
    match = re.search(r'(?:Морозостойкость|F).*?(F\s*\d+)', text, re.IGNORECASE)
    if match: data['frost_resistance'] = match.group(1).replace(' ', '')
    match = re.search(r'(?:Водопоглощение).*?(\d+)\s*%', text, re.IGNORECASE)
    if match: data['water_absorption'] = match.group(1).strip() + "%"
    match = re.search(r'(?:Масса|Вес изделия).*?:?\s*(\d+(?:[.,]\d+)?)\s*(?:кг|т)?', text, re.IGNORECASE)
    if match: data['product_weight'] = match.group(1).replace(',', '.') + " кг"
    match = re.search(r'Вес арматурного каркаса.*?:\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
    if match: data['rebar_weight'] = match.group(1).strip()
    match = re.search(r'Вид и класс стали.*?:\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
    if match: data['reinforcement'] = match.group(1).strip()
    match = re.search(r'Категория бетонной поверхности.*?:\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
    if match: data['surface_category'] = match.group(1).strip()
    match = re.search(r'(?:Проектные размеры|Размеры).*?:\s*(.+?)(?:\n\d|$)', text, re.IGNORECASE)
    if match: data['dimensions'] = match.group(1).strip()
    match = re.search(r'(?:Отклонение|Класс точности).*?:?\s*(\d+)', text, re.IGNORECASE)
    if match: data['accuracy_class'] = match.group(1).strip()
    match = re.search(r'Средняя плотность бетона.*?(\d+)', text, re.IGNORECASE)
    if match: data['concrete_density'] = match.group(1).strip()
    match = re.search(r'Отпускная влажность.*?(\d+)\s*%', text, re.IGNORECASE)
    if match: data['concrete_humidity'] = match.group(1).strip() + "%"
    match = re.search(r'Обозначение стандарта.*?(?:ГОСТ|ТУ)\s*([A-Za-z0-9\-\.]+)', text, re.IGNORECASE)
    if match: data['standard_designation'] = match.group(1).strip()
    match = re.search(r'Номер серии.*?([A-Za-z0-9\-\.]+)', text, re.IGNORECASE)
    if match: data['series_number'] = match.group(1).strip()
    return data

def get_russian_name(key: str) -> str:
    return FIELD_NAMES_RU.get(key, key.replace('_', ' ').title())

def create_passport_page(doc, data):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('ЗАО «ДСК-Столица»')
    run.font.size = Pt(26)
    run.font.name = 'Calibri Light'
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('ИНН 7721823772 / КПП 772101001')
    run.font.size = Pt(12)
    run.font.name = 'Calibri'
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Филиал "Центральный" Банка ВТБ ПАО г. МОСКВА, БИК 044525411, к/с 30101810145250000411, р/с 40702810801880000195')
    run.font.size = Pt(12)
    run.font.name = 'Calibri'
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('109428, город Москва, ул. Проспект Рязанский дом 30/15')
    run.font.size = Pt(12)
    run.font.name = 'Calibri'
    
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("ПАСПОРТ КАЧЕСТВА")
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    doc.add_paragraph()
    
    for key, value in data.items():
        p = doc.add_paragraph()
        run_name = p.add_run(f"{get_russian_name(key)}: ")
        run_name.bold = True
        run_name.font.size = Pt(11)
        run_name.font.name = 'Times New Roman'
        run_value = p.add_run(str(value))
        run_value.font.size = Pt(11)
        run_value.font.name = 'Times New Roman'
    
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run("ОТК")
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    
    p = doc.add_paragraph()
    run = p.add_run("  М.П.")
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'

def process_uploaded_file(uploaded_file, token, model_name):
    """Основная логика обработки файла"""
    # Сохраняем загруженный файл во временную папку
    file_path = TEMP_DIR / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Извлекаем текст
    text = extract_text_from_pdf(file_path)
    
    if not is_valid_text(text):
        return None, "Не удалось извлечь текст. Возможно, это скан без текстового слоя."
    
    passports_text = split_passports(text)
    all_passports_data = []
    
    for passport_text in passports_text:
        prompt = f"""Ты — эксперт по строительным паспортам качества.
В паспорте может быть таблица с несколькими изделиями.

Текст паспорта:
---
{passport_text}
---

ВАЖНО: Если в паспорте есть таблица с НЕСКОЛЬКИМИ изделиями, верни JSON МАССИВ, где каждый объект — это данные для ОДНОГО изделия.
Если в паспорте только одно изделие — верни один JSON объект (не массив).

Формат для одного изделия:
{{
  "passport_number": "номер паспорта",
  "issue_date": "дата выдачи",
  "manufacturing_date": "дата изготовления",
  "product_name": "наименование и марка изделия",
  "volume": "количество этого изделия",
  "gost": "ГОСТ или ТУ",
  "concrete_class": "класс бетона",
  "strength_release": "отпускная прочность",
  "frost_resistance": "морозостойкость",
  "water_absorption": "водонепроницаемость",
  "product_weight": "масса этого изделия кг",
  "concrete_density": "средняя плотность",
  "concrete_humidity": "отпускная влажность %",
  "standard_designation": "обозначение стандарта",
  "series_number": "номер серии чертежей"
}}

ПРАВИЛА:
- НЕ складывай количества и массы разных изделий
- Для каждого изделия из таблицы создай отдельный объект
- Общие поля (номер паспорта, дата, ГОСТ) дублируй для каждого изделия
- Если поле не найдено, ставь null
- Верни ТОЛЬКО JSON (объект или массив), без пояснений
"""
        
        response_text = call_gigachat_text(token, prompt, model_name)
        data_list = []
        
        if response_text:
            try:
                clean_text = response_text.strip()
                if clean_text.startswith("```"):
                    clean_text = clean_text.split("\n", 1)[1] if "\n" in clean_text else clean_text[3:]
                if clean_text.endswith("```"):
                    clean_text = clean_text[:-3]
                
                parsed = json.loads(clean_text.strip())
                if isinstance(parsed, list):
                    data_list = [{k: v for k, v in item.items() if v is not None} for item in parsed]
                elif isinstance(parsed, dict):
                    data_list = [{k: v for k, v in parsed.items() if v is not None}]
            except Exception:
                data_list = []
        
        if len(data_list) == 0:
            regex_data = parse_with_regex(passport_text)
            if regex_data:
                data_list = [regex_data]
        
        all_passports_data.extend(data_list)
    
    if len(all_passports_data) == 0:
        return None, "Не удалось извлечь данные из паспорта."
    
    # Создаем Word документ
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.25)
    section.right_margin = Inches(1.25)
    
    for passport_idx, data in enumerate(all_passports_data, 1):
        if passport_idx > 1:
            doc.add_page_break()
        create_passport_page(doc, data)
    
    passport_num = all_passports_data[0].get('passport_number', 'passport')
    passport_num = str(passport_num).replace('/', '_').replace('\\', '_').replace(' ', '_')
    output_filename = f"passport_{passport_num}_{uuid.uuid4().hex[:6]}.docx"
    output_path = TEMP_DIR / output_filename
    
    doc.save(output_path)
    
    # Очищаем исходный файл
    file_path.unlink(missing_ok=True)
    
    return output_path, None

# ================= ИНТЕРФЕЙС STREAMLIT =================
st.set_page_config(page_title="Парсер паспортов качества", page_icon="🏗️", layout="centered")

st.title("🏗️ Парсер паспортов качества")
st.markdown("Загрузите PDF-файл паспорта завода, и система автоматически создаст готовый Word-документ в корпоративном стиле.")

uploaded_file = st.file_uploader("Выберите PDF-файл", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Загружен файл: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} КБ)")
    
    if st.button("🚀 Обработать паспорт", type="primary"):
        with st.spinner("🔑 Получаем доступ к GigaChat..."):
            token = get_gigachat_token()
            if not token:
                st.error("❌ Ошибка: Не удалось получить токен GigaChat. Проверьте CREDENTIALS в коде.")
                st.stop()
            
            model_name = get_available_models(token)
            st.info(f"✅ Используем модель: {model_name}")
        
        with st.spinner("🧠 Анализируем документ и извлекаем данные... (это может занять 10-20 секунд)"):
            output_path, error_msg = process_uploaded_file(uploaded_file, token, model_name)
            
            if error_msg:
                st.error(f"❌ Ошибка: {error_msg}")
            else:
                st.success("✅ Документ успешно создан!")
                
                # Читаем файл для скачивания
                with open(output_path, "rb") as f:
                    file_bytes = f.read()
                
                st.download_button(
                    label="📥 Скачать готовый Word-документ",
                    data=file_bytes,
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Очищаем временный файл после скачивания (опционально, можно оставить)
                # output_path.unlink(missing_ok=True)

st.markdown("---")
st.caption("Разработано для автоматизации работы с паспортами качества ЗАО «ДСК-Столица»")