#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор презентации <N>_Лекция.pptx из текста лекции <N>_Лекция.md.

Подход: открыть эталон-шаблон (1_Лекция.pptx) как источник темы/мастера/раскладок,
удалить все его слайды и наполнить заново из лекции, сохраняя фирменный стиль
курса (шрифт IBM Plex Sans, акцент #8E39E2, формат 16:9).

Слайд-маркеры `[Слайд N]` / `[Слайды N–M]` в лекции -> по одному слайду с контентом.
`## Часть X. <Название>` -> слайд-раздел (section divider).
`# Текст лекции. ...` -> титульный слайд.
`[иллюстрация]` -> слайд-заглушка «вставить вручную».

Полный текст выступления помещается в заметки (notes) слайда — как в эталоне,
где спикерский текст хранится именно в notes. Видимый текст слайда (body) —
ОБЯЗАТЕЛЬНО перерабатывается агентом в краткие тезисы после генерации.

Зависимости: python-pptx (pip install python-pptx).

Использование:
    python generate_pptx.py <LECTURE.md> <OUT.pptx> [--template <TEMPLATE.pptx>]
"""
import argparse
import io
import re
import sys

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# --- Дизайн-токены (извлечены из эталона 1_Лекция.pptx) -----------------------
SLIDE_W_IN = 13.33
SLIDE_H_IN = 7.5
FONT = "IBM Plex Sans"
TITLE_SZ = Pt(36)
TITLE_BOLD = True
TITLE_CLR = RGBColor(0x15, 0x11, 0x0E)   # почти чёрный заголовок
BODY_SZ = Pt(16)
BODY_CLR = RGBColor(0x00, 0x00, 0x00)    # чёрный основной текст
SECTION_SZ = Pt(44)
SECTION_CLR = RGBColor(0x8E, 0x39, 0xE2) # фирменный фиолетовый (разделы)
ACCENT_CLR = RGBColor(0x8E, 0x39, 0xE2)
COVER_TITLE_SZ = Pt(44)
COVER_SUB_SZ = Pt(20)

DEFAULT_TEMPLATE = r"Модуль 1. Введение в ML\1_Лекция.pptx"

# namespace relationship id
_RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# --- низкоуровневые помощники ----------------------------------------------
def delete_all_slides(prs):
    """Удалить все слайды: убрать sldId и связи (rels) -> чистая колода с темой."""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.attrib[_RNS + "id"]
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def find_layout(prs, *names):
    """Найти первую раскладку по одному из имён."""
    by = {lay.name: lay for lay in prs.slide_layouts}
    for n in names:
        if n in by:
            return by[n]
    return prs.slide_layouts[0] if prs.slide_layouts else None


def style_runs(tf, size, bold=None, color=None, name=FONT):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = name
            r.font.size = size
            if bold is not None:
                r.font.bold = bold
            if color is not None:
                r.font.color.rgb = color


def _set_placeholder(slide, idx, text, size, bold=None, color=None):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text or ""
            style_runs(ph.text_frame, size, bold, color)
            return ph
    return None


def set_notes(slide, notes):
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


# --- построение слайдов ------------------------------------------------------
def add_cover_slide(prs, layout, title, subtitle):
    s = prs.slides.add_slide(layout)
    _set_placeholder(s, 0, title or "", COVER_TITLE_SZ, True, TITLE_CLR)
    # subtitle обычно idx=1 (CENTER_TITLE layout) или idx=1 (SUBTITLE)
    if subtitle:
        _set_placeholder(s, 1, subtitle, COVER_SUB_SZ, None, ACCENT_CLR)
    return s


def add_section_slide(prs, layout, title):
    s = prs.slides.add_slide(layout)
    _set_placeholder(s, 0, title or "", SECTION_SZ, True, SECTION_CLR)
    return s


def add_content_slide(prs, layout, title, body, notes):
    s = prs.slides.add_slide(layout)
    _set_placeholder(s, 0, title or "", TITLE_SZ, TITLE_BOLD, TITLE_CLR)
    if body:
        _set_placeholder(s, 1, body, BODY_SZ, None, BODY_CLR)
    set_notes(s, notes)
    return s


# --- парсинг лекции ----------------------------------------------------------
_H1 = re.compile(r"^#\s+(.+?)\s*$")
_PART = re.compile(r"^##\s+(.+?)\s*$")
_SUB = re.compile(r"^###\s+(.+?)\s*$")          # ### N. <Заголовок> (подраздел)
_MARKER = re.compile(r"^\s*\[Слайд(?:ы)?\s*(\d+(?:\s*[–-]\s*\d+)?)\]\s*(.*)$")
_BULLET = re.compile(r"^\s*[-*]\s+(.+?)\s*$")    # пункт маркированного списка
_INLINE = re.compile(r"`([^`]+)`")              # inline-код
_TABLE = re.compile(r"^\s*\|.+\|\s*$")           # строка markdown-таблицы


_MATH_CMDS = [
    (r"\\mathbb\{([A-Za-z])\}", r"\1"),
    (r"\\hat\{(\w)\}", r"\1"),
    (r"\\bar\{(\w)\}", r"\1"),
    (r"\\arg\\?min", "argmin"),
    (r"\\nabla", "∇"), (r"\\eta", "η"), (r"\\in", "∈"), (r"\\dots", "…"),
    (r"\\le", "≤"), (r"\\ge", "≥"), (r"\\ne", "≠"), (r"\\times", "×"),
    (r"\\cdot", "·"), (r"\\rightarrow", "→"), (r"\\to", "→"),
]
_LATEX_CMD = re.compile(r"\\[a-zA-Z]+")


def _clean_math(s):
    for pat, rep in _MATH_CMDS:
        s = re.sub(pat, rep, s)
    s = s.replace("$", "")          # убрать ограничители inline-математики
    s = _LATEX_CMD.sub("", s)       # прочие \команды выкинуть
    s = s.replace("{", "").replace("}", "")  # остатки групп
    return s


def strip_md(text):
    """Убрать markdown-разметку и inline-математику для чистого текста слайда."""
    text = text.replace("**", "").replace("__", "")
    text = re.sub(r"(?<!\*)\*(?!\*)", "", text)   # курсив, но не внутри **
    text = _INLINE.sub(r"\1", text)               # `код` -> код
    text = _clean_math(text)                       # $...$ и \команды LaTeX
    return text.strip()


def first_sentence(text, limit=140):
    """Короткая выжимка для тела слайда по умолчанию (агент потом переработает)."""
    text = strip_md(text).strip()
    if not text:
        return ""
    for sep in (". ", "! ", "? "):
        i = text.find(sep)
        if 0 < i < limit:
            return text[: i + 1]
    return text[:limit].rstrip() + ("…" if len(text) > limit else "")


def bullets_from_block(lines, max_items=6, max_len=90):
    """Извлечь маркированные пункты из блока строк -> тело слайда тезисами."""
    items = []
    for ln in lines:
        m = _BULLET.match(ln)
        if m:
            t = strip_md(m.group(1))
            if t:
                items.append(t[:max_len])
        if len(items) >= max_items:
            break
    return "\n".join("• " + it for it in items)


def parse_by_markers(text):
    """Режим 1: лекция размечена слайд-маркерами [Слайд N] (стиль Модуля 1)."""
    items = []
    cover_title = None
    section = None
    for ln in text.split("\n"):
        mh = _H1.match(ln)
        if mh and cover_title is None:
            cover_title = strip_md(mh.group(1))
            continue
        pm = _PART.match(ln)
        if pm:
            section = strip_md(pm.group(1))
            items.append(("section", section))
            continue
        mk = _MARKER.match(ln)
        if mk:
            rng, rest = mk.group(1), mk.group(2).strip()
            low = rest.lower()
            is_ill = ("[иллюстрация]" in low) or (rest == "")
            title = section or ("Слайды " + rng.strip())
            body = ("[Иллюстрация — вставить вручную]") if is_ill else first_sentence(rest)
            items.append(("slide", title, body, rest, is_ill))
    if cover_title:
        sub = section if not items else None
        items.insert(0, ("cover", cover_title, sub))
    return items


def parse_by_subsections(text):
    """Режим 2: лекция без слайд-маркеров, разбивается по подразделам ### N. <Заголовок>.
    Каждый подраздел -> контентный слайд; буллеты -> тело, весь текст -> notes."""
    items = []
    cover_title = None
    lines = text.split("\n")

    cur_section_title = None  # для контекста заголовка раздела
    # накопители для текущего подраздела
    cur_title = None
    buf = []

    def flush():
        nonlocal cur_title, buf
        if cur_title is None:
            cur_title, buf = None, []
            return
        bl = bullets_from_block(buf)
        body = bl or first_sentence(" ".join(p for p in buf if p.strip()))
        if not body:
            body = ""
        notes = "\n".join(buf).strip()   # notes — сырой текст выступления (как в режиме маркеров)
        items.append(("slide", cur_title, body, notes, False))
        cur_title, buf = None, []

    for ln in lines:
        mh = _H1.match(ln)
        if mh and cover_title is None:
            cover_title = strip_md(mh.group(1))
            continue
        pm = _PART.match(ln)
        if pm:
            flush()
            cur_section_title = strip_md(pm.group(1))
            items.append(("section", cur_section_title))
            continue
        sm = _SUB.match(ln)
        if sm:
            flush()
            cur_title = strip_md(sm.group(1))
            buf = []
            continue
        if cur_title is not None:
            buf.append(ln)
    flush()

    if cover_title:
        sub = cur_section_title if not items else None
        items.insert(0, ("cover", cover_title, sub))
    return items


def parse_lecture(path):
    """Вернуть список элементов: ('cover', title, subtitle) / ('section', name) /
    ('slide', title, body, notes, is_illustration).

    Автовыбор режима: если в лекции есть слайд-маркеры [Слайд N] — режим 1,
    иначе разбивка по подразделам ### N. (режим 2)."""
    text = io.open(path, encoding="utf-8").read()
    if _MARKER.search("\n".join(text.split("\n"))):
        return parse_by_markers(text)
    return parse_by_subsections(text)


def build(lecture_path, out_path, template_path):
    prs = Presentation(template_path)
    delete_all_slides(prs)

    cover_l = find_layout(prs, "TITLE", "SECTION_HEADER", "OBJECT")
    section_l = find_layout(prs, "SECTION_HEADER", "TITLE", "OBJECT")
    content_l = find_layout(prs, "OBJECT", "TITLE_AND_BODY", "ONE_COLUMN_TEXT")

    items = parse_lecture(lecture_path)
    counts = {"cover": 0, "section": 0, "slide": 0, "illustration": 0}
    for it in items:
        kind = it[0]
        if kind == "cover":
            add_cover_slide(prs, cover_l, it[1], it[2])
            counts["cover"] += 1
        elif kind == "section":
            add_section_slide(prs, section_l, it[1])
            counts["section"] += 1
        elif kind == "slide":
            _, title, body, notes, is_ill = it
            add_content_slide(prs, content_l, title, body, notes)
            counts["slide"] += 1
            if is_ill:
                counts["illustration"] += 1

    prs.save(out_path)
    return counts, len(prs.slides)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Сгенерировать <N>_Лекция.pptx из <N>_Лекция.md")
    ap.add_argument("lecture", help="путь к <N>_Лекция.md")
    ap.add_argument("out", help="путь к <N>_Лекция.pptx (результат)")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE,
                    help="эталон-шаблон (по умолчанию 1_Лекция.pptx Модуля 1)")
    args = ap.parse_args(argv)

    counts, total = build(args.lecture, args.out, args.template)
    print("Готово: {} -> {}".format(args.lecture, args.out))
    print("  титулов: {cover}, разделов: {section}, контентных слайдов: {slide} "
          "(из них иллюстраций-заглушек: {illustration}), всего: {n}".format(
              n=total, **counts))
    return 0


if __name__ == "__main__":
    sys.exit(main())
