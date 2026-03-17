"""
docx_filler.py — заменяет плейсхолдеры в .docx шаблонах, сохраняя форматирование.

Ключевая проблема: Word разбивает текст вроде {{ФИО}} на несколько <w:r> runs в XML.
Решение: склеиваем runs в параграфе перед поиском, затем делаем замену.
"""

import re
import zipfile
import io
from lxml import etree
from copy import deepcopy

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WORD_NS}}}"


def extract_placeholders(docx_bytes: bytes, open_delim: str, close_delim: str) -> set[str]:
    """Извлекает все плейсхолдеры из .docx файла."""
    placeholders = set()
    
    # Escape delimiters for regex
    esc_open = re.escape(open_delim)
    esc_close = re.escape(close_delim)
    pattern = re.compile(f"{esc_open}(.*?){esc_close}")
    xml_tag_pattern = re.compile(r"<[^>]+>")
    
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        for name in zf.namelist():
            if name.endswith(".xml"):
                content = zf.read(name).decode("utf-8", errors="ignore")
                clean_content = xml_tag_pattern.sub("", content)
                matches = pattern.findall(clean_content)
                placeholders.update(matches)
    
    return placeholders


def _merge_runs_in_paragraph(para_elem):
    """
    Склеивает соседние <w:r> с одинаковым форматированием в параграфе.
    Это решает проблему, когда Word разбивает {{ФИО}} на несколько runs.
    """
    runs = para_elem.findall(f".//{W}r")
    if len(runs) < 2:
        return
    
    # Группируем соседние runs с одинаковым rPr
    i = 0
    children = list(para_elem)
    
    while i < len(children) - 1:
        curr = children[i]
        nxt = children[i + 1]
        
        # Только если оба — w:r
        if curr.tag != f"{W}r" or nxt.tag != f"{W}r":
            i += 1
            continue
        
        # Сравниваем форматирование
        curr_rpr = etree.tostring(curr.find(f"{W}rPr")) if curr.find(f"{W}rPr") is not None else b""
        nxt_rpr = etree.tostring(nxt.find(f"{W}rPr")) if nxt.find(f"{W}rPr") is not None else b""
        
        if curr_rpr != nxt_rpr:
            i += 1
            continue
        
        # Склеиваем текст
        curr_t = curr.find(f"{W}t")
        nxt_t = nxt.find(f"{W}t")
        
        if curr_t is None or nxt_t is None:
            i += 1
            continue
        
        curr_text = curr_t.text or ""
        nxt_text = nxt_t.text or ""
        merged = curr_text + nxt_text
        
        curr_t.text = merged
        if merged != merged.strip():
            curr_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        
        # Удаляем следующий run
        para_elem.remove(nxt)
        children = list(para_elem)
        # Не увеличиваем i — проверяем снова с текущей позиции
    
    return


def _replace_in_run(run_elem, replacements: dict, open_delim: str, close_delim: str):
    """Заменяет плейсхолдеры в тексте одного run."""
    t_elem = run_elem.find(f"{W}t")
    if t_elem is None or not t_elem.text:
        return
    
    text = t_elem.text
    changed = False
    
    for key, value in replacements.items():
        placeholder = f"{open_delim}{key}{close_delim}"
        if placeholder in text:
            text = text.replace(placeholder, value)
            changed = True
    
    if changed:
        t_elem.text = text
        # Если в тексте есть пробелы по краям — нужен xml:space="preserve"
        if text != text.strip():
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _process_xml_content(xml_bytes: bytes, replacements: dict, open_delim: str, close_delim: str) -> bytes:
    """
    Обрабатывает XML-файл документа:
    1. Парсит XML
    2. Для каждого параграфа склеивает runs
    3. Заменяет плейсхолдеры в runs
    4. Возвращает изменённый XML
    """
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return xml_bytes
    
    # Обрабатываем все параграфы (w:p) в документе
    for para in root.iter(f"{W}p"):
        _merge_runs_in_paragraph(para)
        for run in para.findall(f".//{W}r"):
            _replace_in_run(run, replacements, open_delim, close_delim)
    
    # Также обрабатываем текст вне параграфов (на всякий случай)
    for run in root.iter(f"{W}r"):
        _replace_in_run(run, replacements, open_delim, close_delim)
    
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_template(
    docx_bytes: bytes,
    replacements: dict,
    open_delim: str = "{{",
    close_delim: str = "}}",
) -> bytes:
    """
    Заполняет .docx шаблон данными из словаря replacements.
    
    Args:
        docx_bytes: содержимое .docx файла
        replacements: словарь {имя_переменной: значение}
        open_delim: открывающий разделитель плейсхолдера (напр. "{{")
        close_delim: закрывающий разделитель плейсхолдера (напр. "}}")
    
    Returns:
        bytes: содержимое заполненного .docx файла
    """
    output_buffer = io.BytesIO()
    
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf_in:
        with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                data = zf_in.read(item.filename)
                
                # Обрабатываем только XML-файлы внутри word/
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    data = _process_xml_content(data, replacements, open_delim, close_delim)
                
                zf_out.writestr(item, data)
    
    output_buffer.seek(0)
    return output_buffer.read()
