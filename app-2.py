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

# --- Step 1: Upload templates ---
st.header("1️⃣ Загрузите шаблоны документов")
uploaded_templates = st.file_uploader(
    "Выберите до 5 файлов Word (.docx)",
    type=["docx"],
    accept_multiple_files=True,
    help="Шаблоны с плейсхолдерами в формате {{КОЛОНКА}}"
)

if uploaded_templates:
    if len(uploaded_templates) > 5:
        st.warning("⚠️ Максимум 5 шаблонов. Будут использованы первые 5.")
        uploaded_templates = uploaded_templates[:5]

    st.success(f"✅ Загружено шаблонов: {len(uploaded_templates)}")
    
    # Show placeholders found in each template
    with st.expander("🔍 Найденные переменные в шаблонах", expanded=True):
        all_placeholders = set()
        for tmpl in uploaded_templates:
            tmpl.seek(0)
            placeholders = extract_placeholders(tmpl.read(), open_delim, close_delim)
            tmpl.seek(0)
            all_placeholders.update(placeholders)
            cols = st.columns([2, 5])
            cols[0].markdown(f"**{tmpl.name}**")
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
        missing = all_placeholders - set(df.columns)
        matched = all_placeholders & set(df.columns)
        
        col1, col2 = st.columns(2)
        with col1:
            if matched:
                st.success(f"**Совпали:** {', '.join(f'`{p}`' for p in sorted(matched))}")
        with col2:
            if missing:
                st.warning(f"**Не найдены в таблице:** {', '.join(f'`{p}`' for p in sorted(missing))}")
    
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
                
                for tmpl_file in uploaded_templates:
                    tmpl_file.seek(0)
                    try:
                        filled_bytes = fill_template(
                            tmpl_file.read(),
                            replacements,
                            open_delim,
                            close_delim
                        )
                        file_path = f"{folder_name}/{tmpl_file.name}"
                        zf.writestr(file_path, filled_bytes)
                    except Exception as e:
                        errors.append(f"Строка {idx + 1}, файл {tmpl_file.name}: {e}")
                    finally:
                        tmpl_file.seek(0)
                    
                    processed += 1
                    progress = processed / total
                    progress_bar.progress(progress, text=f"Обрабатываем: {folder_name} / {tmpl_file.name}")
        
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
