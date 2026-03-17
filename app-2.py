import streamlit as st
import pandas as pd
import zipfile
import io
import os
import re
from docx_filler import fill_template, extract_placeholders

st.set_page_config(
    page_title="Автозаполнение документов",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Автозаполнение документов")
st.markdown("Загрузите шаблоны Word и таблицу с данными сотрудников — получите готовые документы.")

with st.expander("ℹ️ Инструкция по подготовке шаблонов и таблиц"):
    st.markdown("""
    **1. Как подготовить шаблон документа (Word):**
    - Откройте ваш документ `.doc` или `.docx`.
    - В нужных местах вставьте переменные, оборачивая их в специальные символы (по умолчанию `{{...}}`). Например: `{{ФИО}}` или `{{Должность}}`.
    - **Совет:** Убедитесь, что названия внутри скобок совпадают с названиями колонок в таблице. Программа автоматически очистит любые невидимые символы форматирования, которые Word мог вставить случайным образом.
    
    **2. Как подготовить таблицу с данными (Excel / CSV):**
    - Создайте таблицу, где каждая строка — это данные для отдельного комплекта документов.
    - В первой строке (заголовке) напишите названия ваших переменных **точно так же**, как они указаны в шаблоне (но уже без скобок). Например: `ФИО`, `Должность`.
    - Заполните ячейки соответствующими данными.
    
    **3. Что произойдет дальше:**
    - Приложение найдет в тексте шаблона все плейсхолдеры (например, `{{ФИО}}`) и заменит их на значение из колонки `ФИО` для каждой строки вашей таблицы.
    - На выходе вы получите ZIP-архив, внутри которого программа разложит готовые документы по папкам для каждой строчки данных.
    """)
# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Настройки")
    folder_column = st.text_input(
        "Колонка для имени папки",
        value="ФИО",
        help="Имя колонки, значение которой будет использоваться как название папки для каждого сотрудника"
    )
    placeholder_format = st.selectbox(
        "Формат плейсхолдеров",
        ["{{ПЕРЕМЕННАЯ}}", "{ПЕРЕМЕННАЯ}", "[[ПЕРЕМЕННАЯ]]", "<<ПЕРЕМЕННАЯ>>"],
        help="Формат, в котором переменные записаны в шаблонах"
    )
    st.divider()
    st.markdown("**Пример шаблона:**")
    fmt = placeholder_format.replace("ПЕРЕМЕННАЯ", "ФИО")
    st.code(f"Настоящий договор заключён с {fmt}...", language=None)

# Determine delimiter chars from format
fmt_map = {
    "{{ПЕРЕМЕННАЯ}}": ("{{", "}}"),
    "{ПЕРЕМЕННАЯ}": ("{", "}"),
    "[[ПЕРЕМЕННАЯ]]": ("[[", "]]"),
    "<<ПЕРЕМЕННАЯ>>": ("<<", ">>"),
}
open_delim, close_delim = fmt_map[placeholder_format]

import subprocess
import tempfile
import shutil
import os

# --- Step 1: Upload templates ---
st.header("1️⃣ Загрузите шаблоны документов")
uploaded_templates = st.file_uploader(
    "Выберите до 10 файлов Word (.docx, .doc)",
    type=["docx", "doc"],
    accept_multiple_files=True,
    help="Шаблоны с плейсхолдерами в формате {{КОЛОНКА}}"
)

def convert_doc_to_docx(doc_bytes: bytes, filename: str) -> bytes:
    """Converts a .doc file to .docx format using LibreOffice or macOS textutil."""
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_doc_path = os.path.join(temp_dir, filename)
        # Expected path of output from LibreOffice
        out_filename = filename + "x" if not filename.endswith(".docx") else filename
        temp_docx_path = os.path.join(temp_dir, out_filename)
        
        with open(temp_doc_path, "wb") as f:
            f.write(doc_bytes)
            
        try:
            if shutil.which("libreoffice"):
                subprocess.run(["libreoffice", "--headless", "--convert-to", "docx", "--outdir", temp_dir, temp_doc_path], check=True, capture_output=True)
            elif shutil.which("soffice"):
                subprocess.run(["soffice", "--headless", "--convert-to", "docx", "--outdir", temp_dir, temp_doc_path], check=True, capture_output=True)
            elif shutil.which("textutil"):
                subprocess.run(["textutil", "-convert", "docx", "-output", temp_docx_path, temp_doc_path], check=True, capture_output=True)
            else:
                 st.error(f"Не найдена утилита для конвертации (нужен LibreOffice или textutil).")
                 return b""

            # LibreOffice typically generates the file with .docx extension, replacing .doc
            lo_out_path = os.path.join(temp_dir, os.path.splitext(filename)[0] + ".docx")
            if os.path.exists(lo_out_path):
                 temp_docx_path = lo_out_path

            with open(temp_docx_path, "rb") as f:
                return f.read()
        except subprocess.CalledProcessError as e:
            st.error(f"Ошибка конвертации файла {filename}: {e.stderr.decode('utf-8', errors='ignore')}")
            return b""
        except Exception as e:
             st.error(f"Ошибка конвертации файла {filename}: {e}")
             return b""

if uploaded_templates:
    if len(uploaded_templates) > 10:
        st.warning("⚠️ Максимум 10 шаблонов. Будут использованы первые 10.")
        uploaded_templates = uploaded_templates[:10]

    st.success(f"✅ Загружено шаблонов: {len(uploaded_templates)}")
    
    # Pre-process uploaded files to ensure they are .docx bytes
    processed_templates = []
    for tmpl in uploaded_templates:
        tmpl.seek(0)
        file_bytes = tmpl.read()
        if tmpl.name.endswith(".doc"):
            converted_bytes = convert_doc_to_docx(file_bytes, tmpl.name)
            if converted_bytes:
                processed_templates.append({"name": tmpl.name + "x", "bytes": converted_bytes})
        else:
             processed_templates.append({"name": tmpl.name, "bytes": file_bytes})
             
    if not processed_templates:
        st.error("❌ Не удалось обработать загруженные шаблоны.")
        st.stop()
        
    # Show placeholders found in each template
    with st.expander("🔍 Найденные переменные в шаблонах", expanded=True):
        all_placeholders = set()
        template_placeholders = {}
        for tmpl_data in processed_templates:
            placeholders = extract_placeholders(tmpl_data["bytes"], open_delim, close_delim)
            all_placeholders.update(placeholders)
            template_placeholders[tmpl_data['name']] = placeholders
            cols = st.columns([2, 5])
            cols[0].markdown(f"**{tmpl_data['name']}**")
            if placeholders:
                cols[1].markdown(" ".join([f"`{p}`" for p in sorted(placeholders)]))
            else:
                cols[1].markdown("_Переменные не найдены_")

# --- Step 2: Upload data table ---
st.header("2️⃣ Загрузите таблицу с данными")
uploaded_table = st.file_uploader(
    "Файл Excel (.xlsx) или CSV (.csv)",
    type=["xlsx", "csv"],
    help="Каждая строка — один сотрудник. Заголовки колонок должны совпадать с переменными в шаблонах."
)

df = None
if uploaded_table:
    try:
        if uploaded_table.name.endswith(".csv"):
            df = pd.read_csv(uploaded_table)
        else:
            df = pd.read_excel(uploaded_table)
        df = df.fillna("").astype(str)
        
        st.success(f"✅ Загружено строк: {len(df)} | Колонок: {len(df.columns)}")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Check folder column exists
        if folder_column not in df.columns:
            st.warning(f"⚠️ Колонка **{folder_column}** не найдена в таблице. Папки будут пронумерованы (сотрудник_1, сотрудник_2, ...)")
    except Exception as e:
        st.error(f"❌ Ошибка при чтении файла: {e}")

# --- Step 3: Validate & Generate ---
st.header("3️⃣ Сгенерировать документы")

if uploaded_templates and df is not None:
    # Cross-check placeholders vs columns
    if 'all_placeholders' in locals() and all_placeholders:
        missing_overall = all_placeholders - set(df.columns)
        matched = all_placeholders & set(df.columns)
        
        col1, col2 = st.columns(2)
        with col1:
            if matched:
                st.success(f"**Совпали:** {', '.join(f'`{p}`' for p in sorted(matched))}")
        with col2:
            if missing_overall:
                st.warning("⚠️ **В таблице отсутствуют следующие колонки для шаблонов:**")
                for t_name, t_pl in template_placeholders.items():
                    missing_for_t = t_pl - set(df.columns)
                    if missing_for_t:
                        st.markdown(f"**{t_name}**: {', '.join(f'`{p}`' for p in sorted(missing_for_t))}")
    
    if st.button("🚀 Создать ZIP-архив", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="Начинаем...")
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        errors = []
        total = len(df) * len(uploaded_templates)
        processed = 0
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, row in df.iterrows():
                # Determine folder name
                if folder_column in df.columns and row.get(folder_column, "").strip():
                    folder_name = row[folder_column].strip().replace("/", "_").replace("\\", "_")
                else:
                    folder_name = f"сотрудник_{idx + 1}"
                
                replacements = {col: str(val) for col, val in row.items()}
                
                for tmpl_data in processed_templates:
                    try:
                        filled_bytes = fill_template(
                            tmpl_data["bytes"],
                            replacements,
                            open_delim,
                            close_delim
                        )
                        file_path = f"{folder_name}/{tmpl_data['name']}"
                        zf.writestr(file_path, filled_bytes)
                    except Exception as e:
                        errors.append(f"Строка {idx + 1}, файл {tmpl_data['name']}: {e}")
                    
                    processed += 1
                    progress = processed / total
                    progress_bar.progress(progress, text=f"Обрабатываем: {folder_name} / {tmpl_data['name']}")
        
        progress_bar.progress(1.0, text="✅ Готово!")
        
        if errors:
            with st.expander(f"⚠️ Ошибки ({len(errors)})", expanded=True):
                for err in errors:
                    st.error(err)
        
        zip_buffer.seek(0)
        st.download_button(
            label="⬇️ Скачать архив documents.zip",
            data=zip_buffer,
            file_name="documents.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )
        st.balloons()

elif not uploaded_templates:
    st.info("⬆️ Загрузите шаблоны документов для начала работы")
elif df is None:
    st.info("⬆️ Загрузите таблицу с данными сотрудников")
