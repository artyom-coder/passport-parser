import streamlit as st
import os
import re
import json
import uuid
import base64
import io
import pdfplumber
import urllib3
import requests
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= НАСТРОЙКИ =================
CREDENTIALS = st.secrets["gigachat_credentials"]

FIELD_NAMES_RU = {
    'passport_number': 'Номер паспорта', 'issued_to': 'Выдан', 'issue_date': 'Дата выдачи',
    'manufacturing_date': 'Дата изготовления', 'product_name': 'Наименование и марка изделия',
    'volume': 'Количество изделий', 'gost': 'ГОСТ/ТУ', 'concrete_class': 'Класс бетона',
    'strength_required': 'Требуемая прочность', 'strength_release': 'Отпускная прочность',
    'strength_actual': 'Фактическая прочность', 'frost_resistance': 'Морозостойкость',
    'water_absorption': 'Водопоглощение бетона по массе %', 'product_weight': 'Масса изделия кг',
    'reinforcement': 'Арматура', 'surface_category': 'Категория поверхности',
    'dimensions': 'Размеры изделия', 'accuracy_class': 'Класс точности',
    'rebar_weight': 'Вес арматурного каркаса', 'product_type': 'Тип продукции',
    'concrete_density': 'Средняя плотность бетона кг/см2', 'concrete_humidity': 'Отпускная влажность бетона %',
    'standard_designation': 'Обозначение стандарта (ГОСТ, ТУ)', 'series_number': 'Номер серии и выпуска рабочих чертежей',
    'radiation_compliance': 'Соответствие по радионуклидам'
}

TEMP_DIR = Path("temp_web_files")
TEMP_DIR.mkdir(exist_ok=True)

def get_access_token(credentials: str) -> str:
    url = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
    credentials = ''.join(c for c in credentials.strip().encode('ascii', 'ignore').decode('ascii') if c.isprintable() and not c.isspace())
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json", "Authorization": f"Basic {credentials}", "RqUID": str(uuid.uuid4())}
    try:
        response = requests.post(url, headers=headers, data="scope=GIGACHAT_API_PERS", verify=False, timeout=30)
        return response.json().get("access_token") if response.status_code == 200 else None
    except Exception:
        return None

def get_available_models(token: str):
    try:
        response = requests.get("https://api.giga.chat/v1/models", headers={"Authorization": f"Bearer {token}", "Accept": "application/json"}, verify=False, timeout=30)
        if response.status_code == 200:
            models = response.json().get("data", [])
            for pref in ['GigaChat-3-Ultra', 'GigaChat-3-Pro', 'GigaChat-2-Max']:
                for m in models:
                    if pref in m.get('id', ''): return m.get('id')
            return models[0].get('id') if models else "GigaChat"
        return "GigaChat"
    except Exception:
        return "GigaChat"

def call_gigachat_text(token: str, prompt: str, model_name: str) -> str:
    try:
        response = requests.post("https://api.giga.chat/v1/chat/completions", 
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}", "Accept": "application/json"},
            json={"model": model_name, "messages": [{"role": "user", "content": prompt}], "temperature": 0.1, "stream": False},
            verify=False, timeout=120)
        return response.json()['choices'][0]['message']['content'] if response.status_code == 200 else None
    except Exception:
        return None

def call_gigachat_vision(token: str, img_base64: str, model_name: str) -> str:
    try:
        response = requests.post("https://api.giga.chat/v1/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}", "Accept": "application/json"},
            json={
                "model": model_name,
                "messages": [{"role": "user", "content": [
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}},
                    {"type": "text", "text": "Распознай весь текст на этом изображении паспорта качества. Верни ТОЛЬКО текст, сохраняя структуру строк и таблиц, без пояснений."}
                ]}],
                "temperature": 0.1, "stream": False
            }, verify=False, timeout=180)
        return response.json()['choices'][0]['message']['content'] if response.status_code == 200 else None
    except Exception:
        return None

def extract_text_from_pdf(file_bytes: bytes) -> str:
    temp_path = TEMP_DIR / "temp.pdf"
    with open(temp_path, "wb") as f: f.write(file_bytes)
    text = ""
    try:
        with pdfplumber.open(temp_path) as pdf:
            for page in pdf.pages:
                if page.extract_text(): text += page.extract_text() + "\n"
    except Exception: pass
    if len(text) < 50:
        try:
            import fitz
            doc = fitz.open(temp_path)
            for page in doc:
                if page.get_text(): text += page.get_text() + "\n"
            doc.close()
        except Exception: pass
    temp_path.unlink(missing_ok=True)
    return text

def extract_text_from_docx(file_bytes: bytes) -> str:
    try:
        doc = Document(io.BytesIO(file_bytes))
        text = [para.text for para in doc.paragraphs if para.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join([cell.text.strip() for cell in row.cells])
                if row_text.strip(): text.append(row_text)
        return "\n".join(text)
    except Exception: return ""

def extract_text_from_xlsx(file_bytes: bytes) -> str:
    try:
        wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
        text = []
        for sheet in wb.sheetnames:
            for row in wb[sheet].iter_rows(values_only=True):
                row_text = " | ".join([str(cell) if cell is not None else "" for cell in row])
                if row_text.strip(): text.append(row_text)
        return "\n".join(text)
    except Exception: return ""

def is_valid_text(text: str) -> bool:
    if len(text) < 50: return False
    cyrillic_and_digits = sum(1 for c in text if c.isalpha() and c.isascii() == False or c.isdigit())
    total_chars = sum(1 for c in text if not c.isspace())
    return (cyrillic_and_digits / total_chars) > 0.3 if total_chars > 0 else False

def split_passports(text: str) -> list:
    patterns = [r'ТЕХНИЧЕСКИЙ ПАСПОРТ', r'Паспорт качества', r'ПАСПОРТ КАЧЕСТВА']
    positions = sorted([match.start() for pattern in patterns for match in re.finditer(pattern, text, re.IGNORECASE)])
    if not positions: return [text]
    passports = []
    for i, pos in enumerate(positions):
        passports.append(text[pos:positions[i+1]].strip() if i + 1 < len(positions) else text[pos:].strip())
    return [p for p in passports if len(p) > 100] or [text]

def parse_with_regex(text: str) -> dict:
    data = {}
    rules = [
        (r'(?:Паспорт качества|Паспорт|ТЕХНИЧЕСКИЙ ПАСПОРТ).*?№?\s*([A-Za-z0-9_\-/бн]+)', 'passport_number'),
        (r'Выдан\s+(.+?)(?:\n|Дата|$)', 'issued_to'),
        (r'(?:Дата выдачи|ОТ|от).*?(\d{2}\.\d{2}\.\d{4})', 'issue_date'),
        (r'Дата изготовления.*?:\s*(.+?)(?:\n\d|$)', 'manufacturing_date'),
        (r'на изделия из.*?бетонов[:\s]*(.+?)(?:\n\d|$)', 'product_type'),
        (r'Наименование и марка изделия.*?:\s*(.+?)(?:\n\d|$)', 'product_name'),
        (r'(?:Количество|Объём|Партия|Кол-во).*?:\s*(\d+)\s*(?:шт|ед)', 'volume'),
        (r'(?:ГОСТ|ТУ)\s*[:\-]?\s*([A-Za-z0-9\-\.]+)', 'gost'),
        (r'(?:Класс бетона|Марка бетона).*?(В\s*\d+(?:[.,]\d+)?)', 'concrete_class'),
        (r'(?:Rт|Rтр|Требуемая прочность).*?(\d+(?:[.,]\d+)?)\s*МПа', 'strength_required'),
        (r'(?:Rотп|Отпускная прочность).*?(\d+(?:[.,]\d+)?)\s*(?:%|МПа)', 'strength_release'),
        (r'(?:Rф|Rфакт|Фактическая прочность).*?(\d+(?:[.,]\d+)?)\s*МПа', 'strength_actual'),
        (r'(?:Морозостойкость|F).*?(F\s*\d+)', 'frost_resistance'),
        (r'(?:Водопоглощение).*?(\d+)\s*%', 'water_absorption'),
        (r'(?:Масса|Вес изделия).*?:?\s*(\d+(?:[.,]\d+)?)\s*(?:кг|т)?', 'product_weight'),
        (r'Вес арматурного каркаса.*?:\s*(.+?)(?:\n|$)', 'rebar_weight'),
        (r'Вид и класс стали.*?:\s*(.+?)(?:\n|$)', 'reinforcement'),
        (r'Категория бетонной поверхности.*?:\s*(.+?)(?:\n|$)', 'surface_category'),
        (r'(?:Проектные размеры|Размеры).*?:\s*(.+?)(?:\n\d|$)', 'dimensions'),
        (r'(?:Отклонение|Класс точности).*?:?\s*(\d+)', 'accuracy_class'),
        (r'Средняя плотность бетона.*?(\d+)', 'concrete_density'),
        (r'Отпускная влажность.*?(\d+)\s*%', 'concrete_humidity'),
        (r'Обозначение стандарта.*?(?:ГОСТ|ТУ)\s*([A-Za-z0-9\-\.]+)', 'standard_designation'),
        (r'Номер серии.*?([A-Za-z0-9\-\.]+)', 'series_number')
    ]
    for pattern, key in rules:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            if key in ['volume', 'water_absorption', 'concrete_humidity']: val += " %" if "%" not in val else ""
            if key == 'product_weight': val = val.replace(',', '.') + " кг"
            if key in ['concrete_class', 'strength_required', 'strength_actual']: val = val.replace(',', '.')
            data[key] = val
    return data

def get_russian_name(key: str) -> str:
    return FIELD_NAMES_RU.get(key, key.replace('_', ' ').title())

def create_passport_page(doc, data):
    for txt, size, font, bold in [
        ('ЗАО «ДСК-Столица»', 26, 'Calibri Light', False),
        ('ИНН 7721823772 / КПП 772101001', 12, 'Calibri', False),
        ('Филиал "Центральный" Банка ВТБ ПАО г. МОСКВА, БИК 044525411, к/с 30101810145250000411, р/с 40702810801880000195', 12, 'Calibri', False),
        ('109428, город Москва, ул. Проспект Рязанский дом 30/15', 12, 'Calibri', False)
    ]:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(txt); run.font.size = Pt(size); run.font.name = font; run.bold = bold
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("ПАСПОРТ КАЧЕСТВА"); run.bold = True; run.font.size = Pt(14); run.font.name = 'Times New Roman'
    doc.add_paragraph()
    for key, value in data.items():
        p = doc.add_paragraph()
        run_name = p.add_run(f"{get_russian_name(key)}: "); run_name.bold = True; run_name.font.size = Pt(11); run_name.font.name = 'Times New Roman'
        run_value = p.add_run(str(value)); run_value.font.size = Pt(11); run_value.font.name = 'Times New Roman'
    doc.add_paragraph()
    for txt in ["ОТК", "  М.П."]:
        p = doc.add_paragraph(); run = p.add_run(txt); run.font.size = Pt(11); run.font.name = 'Times New Roman'

def process_file(uploaded_file, token, model_name):
    file_bytes = uploaded_file.getvalue()
    file_name = uploaded_file.name
    ext = file_name.split('.')[-1].lower()
    
    text = ""
    if ext == 'pdf':
        text = extract_text_from_pdf(file_bytes)
    elif ext in ['docx', 'doc']:
        text = extract_text_from_docx(file_bytes)
    elif ext in ['xlsx', 'xls']:
        text = extract_text_from_xlsx(file_bytes)
    elif ext in ['jpg', 'jpeg', 'png']:
        img_base64 = base64.b64encode(file_bytes).decode('utf-8')
        text = call_gigachat_vision(token, img_base64, model_name) or ""

    if not is_valid_text(text):
        return None, f"Не удалось извлечь текст из {file_name}. Возможно, это нечитаемый скан."
    
    all_passports_data = []
    for passport_text in split_passports(text):
        prompt = f"""Ты — эксперт по строительным паспортам качества. Если в паспорте таблица с НЕСКОЛЬКИМИ изделиями, верни JSON МАССИВ объектов. Если одно — один объект.
Текст: ---\n{passport_text}\n---
Формат: {{"passport_number": "...", "product_name": "...", "volume": "...", "product_weight": "...", "gost": "...", "concrete_class": "...", "strength_release": "...", "frost_resistance": "...", "water_absorption": "...", "concrete_density": "...", "concrete_humidity": "...", "standard_designation": "...", "series_number": "..."}}
ПРАВИЛА: НЕ складывай массы/количества разных изделий. Для каждого изделия отдельный объект. Общие поля дублируй. Верни ТОЛЬКО JSON."""
        
        response_text = call_gigachat_text(token, prompt, model_name)
        data_list = []
        if response_text:
            try:
                clean = response_text.strip().replace("```json", "").replace("```", "").strip()
                parsed = json.loads(clean)
                data_list = [{k: v for k, v in item.items() if v is not None} for item in parsed] if isinstance(parsed, list) else [{k: v for k, v in parsed.items() if v is not None}]
            except Exception: pass
        
        if not data_list:
            regex_data = parse_with_regex(passport_text)
            if regex_data: data_list = [regex_data]
        
        all_passports_data.extend(data_list)

    if not all_passports_data:
        return None, f"Не удалось распознать данные в {file_name}."

    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(1)
    section.left_margin = section.right_margin = Inches(1.25)
    
    for idx, data in enumerate(all_passports_data, 1):
        if idx > 1: doc.add_page_break()
        create_passport_page(doc, data)
    
    passport_num = all_passports_data[0].get('passport_number', 'passport').replace('/', '_').replace(' ', '_')
    output_path = TEMP_DIR / f"passport_{passport_num}_{uuid.uuid4().hex[:6]}.docx"
    doc.save(output_path)
    return output_path, None

# ================= ИНТЕРФЕЙС =================
st.set_page_config(page_title="Парсер паспортов качества", page_icon="🏗️", layout="centered")

# Попытка загрузить логотип (если файл logo.png добавлен в репозиторий)
try:
    st.image("logo.png", width=150)
except Exception:
    pass

st.title("🏗️ Парсер паспортов качества ЗАО «ДСК-Столица»")
st.markdown("Загрузите **один или несколько** файлов (PDF, Word, Excel, JPG, PNG). Система автоматически извлечёт данные и создаст готовые Word-документы.")

uploaded_files = st.file_uploader("Выберите файлы", type=["pdf", "docx", "doc", "xlsx", "xls", "jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    st.info(f"Выбрано файлов: **{len(uploaded_files)}**")
    
    if st.button("🚀 Обработать все файлы", type="primary"):
        token = get_access_token(CREDENTIALS)
        if not token:
            st.error("❌ Ошибка: Не удалось получить токен GigaChat.")
            st.stop()
        
        model_name = get_available_models(token)
        results = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Обработка {i+1}/{len(uploaded_files)}: {uploaded_file.name}...")
            progress_bar.progress((i + 1) / len(uploaded_files))
            
            output_path, error_msg = process_file(uploaded_file, token, model_name)
            if error_msg:
                results.append({"name": uploaded_file.name, "status": "error", "msg": error_msg})
            else:
                with open(output_path, "rb") as f:
                    results.append({"name": uploaded_file.name, "status": "success", "data": f.read(), "filename": output_path.name})
                output_path.unlink(missing_ok=True)
        
        status_text.text("✅ Обработка завершена!")
        progress_bar.empty()
        
        st.markdown("---")
        for res in results:
            if res["status"] == "success":
                st.success(f"✅ {res['name']}")
                st.download_button(label=f"📥 Скачать {res['filename']}", data=res["data"], file_name=res["filename"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.error(f"❌ {res['name']}: {res['msg']}")

st.markdown("---")
st.caption("Разработано для автоматизации работы с паспортами качества")
