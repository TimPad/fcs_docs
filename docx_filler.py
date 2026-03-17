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
    
    def build_regex(s):
        return r'(?:<[^>]+>)*'.join(re.escape(c) for c in s)
        
    esc_open = build_regex(open_delim)
    esc_close = build_regex(close_delim)
    pattern = re.compile(f"{esc_open}(.*?){esc_close}", flags=re.DOTALL)
    xml_tag_pattern = re.compile(r"<[^>]+>", flags=re.DOTALL)
    
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        for name in zf.namelist():
            if name.endswith(".xml"):
                content = zf.read(name).decode("utf-8", errors="ignore")
                clean_content = xml_tag_pattern.sub("", content)
                matches = pattern.findall(clean_content)
                placeholders.update(matches)
    
    return placeholders


def _process_xml_content(xml_bytes: bytes, replacements: dict, open_delim: str, close_delim: str) -> bytes:
    """
    Обрабатывает XML-файл документа, заменяя переменные напрямую в строке.
    """
    content = xml_bytes.decode("utf-8", errors="ignore")
    
    def build_regex(s):
        return r'(?:<[^>]+>)*'.join(re.escape(c) for c in s)
        
    esc_open = build_regex(open_delim)
    esc_close = build_regex(close_delim)
    pattern = re.compile(f"{esc_open}(.*?){esc_close}", flags=re.DOTALL)
    xml_tag_pattern = re.compile(r"<[^>]+>", flags=re.DOTALL)
    
    def replace_match(match):
        raw_inner = match.group(1)
        clean_name = xml_tag_pattern.sub("", raw_inner)
        
        if clean_name in replacements:
            val = str(replacements[clean_name])
            # Escape specially for XML
            val = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            return val
        
        return match.group(0)
        
    result_content = pattern.sub(replace_match, content)
    
    return result_content.encode("utf-8")


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
