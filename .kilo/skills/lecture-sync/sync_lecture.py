#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Синхронизатор изменений между <N>_Лекция.md и <N>_Лекция.pptx (ИНКРЕМЕНТНО).

В отличие от generate_pptx.py (pptx-writer), который **полностью пересобирает** колоду
и затирает ручную работу (иллюстрации, правку слайдов), этот скрипт вносит только
дельту: добавляет новые слайды/текст, обновляет изменившееся — и не трогает остальное.

Поддерживает оба формата лекций курса:
  * Формат 1 (эталон, Модуль 1): md с маркерами `[Слайд N]` / `[Слайды N–M]`,
    текст выступления живёт в notes слайдов.
  * Формат 2 (модули 2+): md с подразделами `### N. <Заголовок>`, заголовок слайда
    pptx начинается с `N.`.

Две команды:
  report <md> <pptx>              — изучить расхождения, вывести человекочитаемый
                                    отчёт + JSON-сводку (ничего не меняет).
  apply  <md> <pptx>              — применить изменения в заданном направлении.

Параметры apply:
  --direction md2pptx|pptx2md     куда переносить изменения (по умолчанию md2pptx).
  --mode add|update|all           add — только добавить новое (по умолчанию, безопасно);
                                    update — ещё и перезаписать изменившиеся тексты/заголовки;
                                    all — ещё и удалить лишнее и пересортировать.
  --dry-run                       показать план, но не записывать файлы.
  --backup                        сохранить .bak перед перезаписью.

Зависимости: python-pptx. Переиспользует оформление/парсеры из generate_pptx.py.
"""
import argparse
import io
import json
import os
import re
import shutil
import sys

# --- переиспользуем оформление и парсеры из соседнего скилла pptx-writer ---------
_HERE = os.path.dirname(os.path.abspath(__file__))
_PPTX_WRITER = os.path.normpath(os.path.join(_HERE, "..", "pptx-writer"))
if _PPTX_WRITER not in sys.path:
    sys.path.insert(0, _PPTX_WRITER)

from pptx import Presentation  # noqa: E402

from generate_pptx import (  # noqa: E402
    _H1, _PART, _SUB, _MARKER,
    add_cover_slide, add_content_slide, add_section_slide,
    find_layout, first_sentence, strip_md,
)

# --- регулярки -----------------------------------------------------------------
_NUM_IN_TITLE = re.compile(r"^\s*(\d+)[\.\s]")          # "12. ..." в заголовке слайда
_SUB_NUM = re.compile(r"^###\s+(?:\*\*)?(\d+)\.\s*(.+?)\**\s*$")
_MARKER_NUM = re.compile(r"^\s*\[Слайд(?:ы)?\s*(\d+)")
_IS_SECTION_TITLE = re.compile(r"часть|раздел|заключение", re.I)


# =============================================================================
# ПАРСИНГ .md  (оба формата) — units с номерами и диапазонами строк
# =============================================================================
def _struct_kind(line):
    """К какому структурному элементу относится строка md."""
    if _H1.match(line):
        return "h1"
    if _PART.match(line):
        return "part"
    if _SUB.match(line):
        return "sub"
    if _MARKER.match(line):
        return "marker"
    return None


def parse_md(path):
    """Разобрать лекцию в упорядоченный список unit-ов (оба формата).

    Каждый unit — dict:
      kind      : 'cover' | 'section' | 'content'
      num       : int | None            (номер слайда/подраздела)
      title     : str                   (cover — название лекции; section — заголовок части;
                                         content — 'N. <name>' (формат 2) или имя части (формат 1))
      text      : str                   (текст выступления → notes; без маркера и без '---')
      is_ill    : bool                  ([иллюстрация])
      header_ln : int                   (0-based индекс строки-заголовка; -1 для cover)
      block_lo  : int                   (0-based первая строка блока unit-а в файле)
      block_hi  : int                   (0-based последняя строка блока включительно, с '---'/пустыми)
      inline    : str|None              (для формата 1 — текст после '[Слайд N]' в строке маркера)
    """
    lines = io.open(path, encoding="utf-8").read().split("\n")
    use_markers = any(_MARKER.match(ln) for ln in lines)
    units = []
    cover_title = None

    # проход 1: собрать unit-ы с координатой заголовка
    for i, ln in enumerate(lines):
        k = _struct_kind(ln)
        if k == "h1":
            if cover_title is None:
                cover_title = strip_md(_H1.match(ln).group(1))
            continue
        if k == "part":
            units.append(dict(kind="section", num=None,
                              title=strip_md(_PART.match(ln).group(1)),
                              header_ln=i, inline=None))
            continue
        if use_markers and k == "marker":
            num = int(_MARKER_NUM.match(ln).group(1))
            rest = _MARKER.match(ln).group(2).strip()
            units.append(dict(kind="content", num=num, title=None, header_ln=i,
                              inline=rest, is_ill=("[иллюстрация]" in rest.lower())))
            continue
        if (not use_markers) and k == "sub":
            sm = _SUB.match(ln)
            mn = _SUB_NUM.match(ln)
            num = int(mn.group(1)) if mn else None
            units.append(dict(kind="content", num=num,
                              title=strip_md(sm.group(1)),
                              header_ln=i, inline=None, is_ill=False))
            continue

    # формат 1: title content-юнита = заголовок текущей части (как в generate_pptx)
    cur_section = None
    for u in units:
        if u["kind"] == "section":
            cur_section = u["title"]
        elif u["kind"] == "content" and u.get("title") is None:
            u["title"] = cur_section

    # проход 2: диапазоны блоков и текст
    headers = [u["header_ln"] for u in units]
    n = len(lines)
    for idx, u in enumerate(units):
        lo = u["header_ln"]
        hi = (headers[idx + 1] - 1) if idx + 1 < len(headers) else (n - 1)
        u["block_lo"], u["block_hi"] = lo, hi
        u["text"] = _extract_text(lines, u, lo, hi)

    if cover_title:
        units.insert(0, dict(kind="cover", num=None, title=cover_title, text="",
                             is_ill=False, header_ln=-1, block_lo=-1, block_hi=-1, inline=None))
    return units, lines


def _extract_text(lines, u, lo, hi):
    """Текст выступления unit-а: для формата 1 убираем префикс '[Слайд N]',
    для формата 2 берём тело после '### ...'; хвостовые '---' и пустые строки отбрасываем."""
    if u["kind"] != "content":
        return ""
    if u.get("inline") is not None:               # формат 1: маркер
        first = u["inline"]
        body = ([first] if first else []) + lines[lo + 1:hi + 1]
    else:                                          # формат 2: подраздел
        body = list(lines[lo + 1:hi + 1])
    while body and body[-1].strip() in ("", "---"):
        body.pop()
    return "\n".join(body).strip()


def _make_md_block(num, title, text, use_markers):
    """Собрать строки md-блока для одного слайда (для вставки/замены)."""
    text = (text or "").strip()
    if use_markers:
        first, *rest = text.split("\n") if text else [""]
        head = "[Слайд{}] {}".format((" " + str(num)) if num is not None else "", first).strip()
        return [head] + rest + [""]
    name = re.sub(r"^\s*\d+\.\s*", "", title or "").strip() or "Слайд"
    name = name.replace("**", "")
    block = ["---", "", "### **{}. {}**".format(num if num is not None else "?", name), ""]
    if text:
        block += text.split("\n")
    block += ["", "---"]
    return block


# =============================================================================
# ПАРСИНГ .pptx  -> units
# =============================================================================
def parse_pptx(path):
    """Пройти по слайдам pptx -> упорядоченный список unit-ов (как в parse_md)."""
    prs = Presentation(path)
    units = []
    for idx, s in enumerate(prs.slides):
        title, body, notes = "", "", ""
        for ph in s.placeholders:
            if not ph.has_text_frame:
                continue
            t = ph.text_frame.text
            if ph.placeholder_format.idx == 0:
                title = t
            elif ph.placeholder_format.idx == 1 and t:
                body = (body + "\n" + t) if body else t
        if s.has_notes_slide:
            notes = s.notes_slide.notes_text_frame.text
        m = _NUM_IN_TITLE.match(title or "")
        num = int(m.group(1)) if m else None
        kind = _classify_slide(idx, title, body, notes, num)
        units.append(dict(kind=kind, num=num, title=title.strip(), body=body.strip(),
                          text=notes.strip(), is_ill=False, slide=s, idx=idx))
    return prs, units


def _classify_slide(idx, title, body, notes, num):
    if num is not None:
        return "content"
    if idx == 0:
        return "cover"
    if _IS_SECTION_TITLE.search(title or ""):
        return "section"
    if not title and not body and not notes:
        return "section"  # пустой слайд-раздел/разделитель
    return "content"


# =============================================================================
# ВЫРАВНИВАНИЕ md <-> pptx  (числовой или позиционный режим)
# =============================================================================
def align(md_units, pptx_units):
    """Сопоставить unit-ы двух сторон.

    Возвращает:
      pairs        : list[ (md_unit|None, pptx_unit|None) ]
      only_md      : list[md_unit]      — есть только в md
      only_pptx    : list[pptx_unit]    — есть только в pptx
      mode         : 'numeric' | 'positional'
    """
    md_num = [u for u in md_units if u["kind"] == "content" and u["num"] is not None]
    px_num = [u for u in pptx_units if u["kind"] == "content" and u["num"] is not None]
    numeric = bool(md_num) and bool(px_num) and \
        len(md_num) >= 0.5 * max(1, sum(1 for u in md_units if u["kind"] == "content")) and \
        len(px_num) >= 0.5 * max(1, sum(1 for u in pptx_units if u["kind"] == "content"))

    pairs, used_md, used_px = [], set(), set()
    if numeric:
        px_by_num = {}
        for u in pptx_units:
            if u["kind"] == "content" and u["num"] is not None and u["num"] not in px_by_num:
                px_by_num[u["num"]] = u
        # cover
        md_cov = next((u for u in md_units if u["kind"] == "cover"), None)
        px_cov = next((u for u in pptx_units if u["kind"] == "cover"), None)
        if md_cov or px_cov:
            pairs.append((md_cov, px_cov))
            if md_cov:
                used_md.add(id(md_cov))
            if px_cov:
                used_px.add(id(px_cov))
        # секции и контент — единым проходом по md, сохраняя порядок
        md_by_num = {u["num"]: u for u in md_num}
        all_nums = sorted(set(list(md_by_num) + list(px_by_num)))
        for n in all_nums:
            mu = md_by_num.get(n)
            pu = px_by_num.get(n)
            pairs.append((mu, pu))
            if mu and pu:                      # реально совпали — отмечаем использованными;
                used_md.add(id(mu))            # md-only (pu=None) останется в only_md
                used_px.add(id(pu))
        # секции без номера — по заголовку среди свободных
        for u in md_units:
            if u["kind"] != "section" or id(u) in used_md:
                continue
            match = _find_by_title(u["title"], [p for p in pptx_units if p["kind"] == "section"
                                                and id(p) not in used_px])
            if match:
                pairs.append((u, match))
                used_md.add(id(u))
                used_px.add(id(match))
            else:
                pairs.append((u, None))
    else:
        # позиционно: контент-юниты по порядку
        md_c = [u for u in md_units if u["kind"] in ("cover", "content", "section")]
        px_c = [u for u in pptx_units if u["kind"] in ("cover", "content", "section")]
        for mu, pu in zip(md_c, px_c):
            pairs.append((mu, pu))
            used_md.add(id(mu))
            used_px.add(id(pu))
        for u in md_c:
            if id(u) not in used_md:
                pairs.append((u, None))

    only_md = [u for u in md_units if id(u) not in used_md and u["kind"] == "content"]
    only_pptx = [u for u in pptx_units if id(u) not in used_px and u["kind"] == "content"]
    # упорядочить только-в-md по номеру/порядку
    only_md.sort(key=lambda u: (u["num"] if u["num"] is not None else 1e9,
                                md_units.index(u)))
    only_pptx.sort(key=lambda u: (u["num"] if u["num"] is not None else 1e9,
                                  pptx_units.index(u)))
    return pairs, only_md, only_pptx, ("numeric" if numeric else "positional")


def _find_by_title(title, candidates):
    tgt = _norm(title)
    if not tgt:
        return None
    for c in candidates:
        if _norm(c["title"]) == tgt:
            return c
    for c in candidates:
        if tgt and tgt in _norm(c["title"]):
            return c
    return None


# =============================================================================
# ДИФФ / ОТЧЁТ
# =============================================================================
def _norm(t):
    t = strip_md(t or "")
    t = "\n".join(l for l in t.split("\n") if l.strip() != "---")  # игнорировать разделители
    return " ".join(t.split())


def build_diff(md_path, pptx_path):
    md_units, _ = parse_md(md_path)
    prs, px_units = parse_pptx(pptx_path)
    pairs, only_md, only_pptx, mode = align(md_units, px_units)

    changed_text, changed_title = [], []
    for mu, pu in pairs:
        if not (mu and pu):
            continue
        if _norm(mu["text"]) != _norm(pu["text"]):
            changed_text.append(dict(num=_key(mu, pu), title=pu["title"] or mu["title"]))
        if _norm(mu["title"]) and _norm(pu["title"]) and \
                _norm(mu["title"]) != _norm(pu["title"]):
            changed_title.append(dict(num=_key(mu, pu),
                                      md=mu["title"], pptx=pu["title"]))
    return dict(mode=mode, md_count=len(md_units), pptx_count=len(px_units),
                pairs=pairs, only_md=only_md, only_pptx=only_pptx,
                changed_text=changed_text, changed_title=changed_title,
                prs=prs, px_units=px_units, md_units=md_units)


def _key(mu, pu):
    for u in (mu, pu):
        if u and u.get("num") is not None:
            return u["num"]
    return None


def render_report(diff):
    out = []
    out.append("РЕЖИМ СОПОСТАВЛЕНИЯ: {}".format(diff["mode"]))
    out.append("md unit-ов: {md_count} | pptx слайдов: {pptx_count}".format(**diff))
    out.append("")
    out.append("Только в md (нужно добавить в pptx): {}".format(len(diff["only_md"])))
    for u in diff["only_md"]:
        out.append("  - слайд {}: {}".format(u["num"], _short(u["title"] or u["text"])))
    out.append("")
    out.append("Только в pptx (нужно добавить в md): {}".format(len(diff["only_pptx"])))
    for u in diff["only_pptx"]:
        out.append("  - слайд {}: {}".format(u["num"], _short(u["title"] or u["text"])))
    out.append("")
    out.append("Изменился текст выступления (notes): {}".format(len(diff["changed_text"])))
    for c in diff["changed_text"]:
        out.append("  - слайд {}: {}".format(c["num"], _short(c["title"])))
    out.append("")
    out.append("Изменились заголовки: {}".format(len(diff["changed_title"])))
    for c in diff["changed_title"]:
        out.append("  - слайд {}: md={!r} -> pptx={!r}".format(c["num"], _short(c["md"]), _short(c["pptx"])))
    return "\n".join(out)


def _short(s, n=70):
    s = strip_md(s or "").replace("\n", " ").strip()
    return s if len(s) <= n else s[:n - 1] + "…"


# =============================================================================
# ПРИМЕНЕНИЕ:  md -> pptx
# =============================================================================
def _sldid_for_slide(prs, slide):
    """Найти sldId-элемент слайда по стабильному partname (Slide-объекты не хэшируемы)."""
    target = slide.part.partname
    for s, el in zip(prs.slides, prs.slides._sldIdLst):
        if s.part.partname == target:
            return el
    return None


def _reorder(prs, desired_slides):
    """Переупорядочить sldIdLst точно под desired_slides (существующие слайды)."""
    by_partname = {}
    for s, el in zip(prs.slides, prs.slides._sldIdLst):
        by_partname[s.part.partname] = el
    ordered = [by_partname[s.part.partname] for s in desired_slides
               if s.part.partname in by_partname]
    sldIdLst = prs.slides._sldIdLst
    for el in list(sldIdLst):
        sldIdLst.remove(el)
    for el in ordered:
        sldIdLst.append(el)


def apply_md2pptx(diff, md_path, pptx_path, mode, dry_run, backup, template):
    prs = diff["prs"]
    pairs = diff["pairs"]
    only_md = diff["only_md"]

    cover_l = find_layout(prs, "TITLE", "SECTION_HEADER", "OBJECT")
    section_l = find_layout(prs, "SECTION_HEADER", "TITLE", "OBJECT")
    content_l = find_layout(prs, "OBJECT", "TITLE_AND_BODY", "ONE_COLUMN_TEXT")

    used_pptx_ids = {id(pu) for (_, pu) in pairs if pu}
    # 1) обновить существующие пары (update/all)
    updates = 0
    for mu, pu in pairs:
        if not (mu and pu):
            continue
        if mode in ("update", "all"):
            if _norm(mu["text"]) != _norm(pu["text"]) and mu["text"]:
                _set_notes(pu["slide"], mu["text"])
                updates += 1
            if _norm(mu["title"]) and _norm(pu["title"]) and \
                    _norm(mu["title"]) != _norm(pu["title"]):
                _set_title(pu["slide"], mu["title"])
                updates += 1
        elif mode == "add":
            # добавить notes только если его нет
            if not pu["text"] and mu["text"]:
                _set_notes(pu["slide"], mu["text"])
                updates += 1

    # 2) план порядка: проходим md по порядку, вставляем новые слайды
    plan = []  # list of slide-объектов в желаемом порядке
    md_units = diff["md_units"]
    pair_by_mu = {id(mu): pu for (mu, pu) in pairs}
    only_md_ids = {id(u) for u in only_md}
    created = 0
    for mu in md_units:
        pu = pair_by_mu.get(id(mu))
        if mu["kind"] == "cover":
            if pu:
                plan.append(pu["slide"])
            continue
        if pu:
            plan.append(pu["slide"])
            continue
        if id(mu) in only_md_ids:
            s = _build_slide_from_md(prs, mu, content_l, section_l, cover_l)
            plan.append(s)
            created += 1

    # 3) preserved-only-pptx слайды: в режиме add/update оставляем, якоря после
    #    предшествующего совпавшего. в all — удаляем.
    preserved = []
    for pu in diff["px_units"]:
        if id(pu) in used_pptx_ids or pu["kind"] == "cover":
            continue
        if mode == "all":
            _drop_slide(prs, pu["slide"])
        else:
            preserved.append(pu)

    # вставить preserved в план после их исходного соседа
    if preserved:
        plan = _interleave_preserved(plan, preserved, diff["px_units"])

    # 4) пересортировать колоду под план
    _reorder(prs, plan)

    summary = dict(direction="md2pptx", mode=mode,
                   added=created, updated=updates,
                   deleted=(len(preserved) if mode == "all" else 0),
                   preserved=(0 if mode == "all" else len(preserved)),
                   total_slides=len(list(prs.slides)))
    _save_or_dry(prs, pptx_path, dry_run, backup)
    return summary


def _build_slide_from_md(prs, mu, content_l, section_l, cover_l):
    title = mu["title"]
    if mu["kind"] == "section":
        return add_section_slide(prs, section_l, title)
    num = mu.get("num")
    if title and num is not None and not _NUM_IN_TITLE.match(title):
        title = "{}. {}".format(num, title)
    elif not title:
        title = "Слайд {}".format(num) if num is not None else "Слайд"
    body = "[Иллюстрация — вставить вручную]" if mu.get("is_ill") else \
        (first_sentence(mu["text"]) or "")
    return add_content_slide(prs, content_l, title, body, mu["text"])


def _interleave_preserved(plan, preserved, px_units):
    """Вернуть новый план, в который preserved-слайды вставлены после того слайда,
    что в исходной pptx-колоде шёл прямо перед ними (или в начало, если нет)."""
    slide_to_pu = {id(pu["slide"]): pu for pu in px_units}
    plan_set = [id(s) for s in plan]
    result = []
    for s in plan:
        result.append(s)
        sid = id(s)
        pu = slide_to_pu.get(sid)
        if pu is None:
            continue
        idx = pu["idx"]
        # preserved, чей исходный индекс = idx+1, idx+2 ... подряд, пока идут в плане сразу после
        # простой эвристик: прикрепляем preserved-слайды с наибольшим idx <= текущего соседа
        attached = [p for p in preserved
                    if p["idx"] > idx and (slide_to_pu.get(id(_prev_in(plan, s))) is None)]
        # упростим: preserved с idx == idx+1 попадают сюда
        attached = sorted([p for p in preserved if p["idx"] > idx],
                          key=lambda p: p["idx"])
        # берём только те, что непосредственно следуют в исходной колоде без разрыва
        chain = []
        expect = idx + 1
        for p in attached:
            if p["idx"] == expect:
                chain.append(p)
                expect += 1
            else:
                break
        for p in chain:
            if id(p["slide"]) in plan_set:
                continue
            result.append(p["slide"])
    return result


def _prev_in(plan, s):
    i = plan.index(s)
    return plan[i - 1] if i > 0 else None


def _set_notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def _set_title(slide, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = text
            return


def _drop_slide(prs, slide):
    el = _sldid_for_slide(prs, slide)
    if el is None:
        return
    from pptx.oxml.ns import qn
    rid = el.get(qn("r:id"))
    prs.slides._sldIdLst.remove(el)
    if rid:
        prs.part.drop_rel(rid)


# =============================================================================
# ПРИМЕНЕНИЕ:  pptx -> md
# =============================================================================
def apply_pptx2md(diff, md_path, pptx_path, mode, dry_run, backup):
    md_units, lines = parse_md(md_path)
    pairs = diff["pairs"]
    use_markers = any(_MARKER.match(ln) for ln in lines)

    used_md_ids = {id(mu) for (mu, _) in pairs if mu}
    ops = []  # ('replace', block_lo, block_hi, new_lines) | ('insert', at_line, block_lines)

    # 1) обновить текст совпавших слайдов
    updates = 0
    for mu, pu in pairs:
        if not (mu and pu) or mu["header_ln"] < 0:
            continue
        changed = pu["text"] and _norm(mu["text"]) != _norm(pu["text"])
        fill = (mode == "add" and not mu["text"] and pu["text"])
        if (mode in ("update", "all") and changed) or fill:
            block = _make_md_block(mu["num"], mu["title"], pu["text"], use_markers)
            ops.append(("replace", mu["block_lo"], mu["block_hi"], block))
            updates += 1

    # 2) только-в-pptx -> вставить новые md-блоки
    created = 0
    for pu in diff["only_pptx"]:
        anchor = _md_anchor_for(pu, pairs)
        at_line = _md_insertion_point(anchor, md_units)
        block = _make_md_block(pu.get("num"), pu["title"], pu["text"], use_markers)
        ops.append(("insert", at_line, at_line, block))
        created += 1

    # 3) удалить только-в-md (режим all)
    deleted = 0
    if mode == "all":
        del_units = sorted((u for u in md_units if id(u) not in used_md_ids
                            and u["header_ln"] >= 0), key=lambda u: u["block_lo"])
        for u in del_units:
            ops.append(("delete", u["block_lo"], u["block_hi"], None))
            deleted += 1

    new_text = _apply_line_ops(lines, ops)
    summary = dict(direction="pptx2md", mode=mode, added=created, updated=updates,
                   deleted=deleted, ops=len(ops))
    if not dry_run:
        if backup:
            _backup(md_path)
        with io.open(md_path, "w", encoding="utf-8") as f:
            f.write(new_text)
    return summary


def _md_anchor_for(pu, pairs):
    """md-unit, предшествующий pptx-unit-у pu по номеру/порядку среди совпавших."""
    best = None
    best_key = -1
    for mu, x in pairs:
        if not (mu and x):
            continue
        if x.get("num") is not None and pu.get("num") is not None and x["num"] < pu["num"]:
            if x["num"] > best_key:
                best, best_key = mu, x["num"]
        elif x.get("idx", -1) < pu.get("idx", 1e9):
            if x["idx"] > best_key:
                best, best_key = mu, x["idx"]
    return best


def _md_insertion_point(anchor, md_units):
    if anchor is not None and anchor["block_hi"] >= 0:
        return anchor["block_hi"] + 1
    last = max((u["block_hi"] for u in md_units if u["block_hi"] >= 0), default=-1)
    return (last + 1) if last >= 0 else 0


def _apply_line_ops(lines, ops):
    """Применить операции к списку строк; сортируем по позиции убывающе."""
    arr = list(lines)

    def poskey(o):
        return o[1]
    for kind, p1, p2, payload in sorted(ops, key=poskey, reverse=True):
        if kind == "replace":
            arr[p1:p2 + 1] = payload
        elif kind == "insert":
            arr[p1:p1] = payload
        elif kind == "delete":
            del arr[p1:p2 + 1]
    return "\n".join(arr)


# =============================================================================
# IO помощники
# =============================================================================
def _save_or_dry(prs, path, dry_run, backup):
    if dry_run:
        return
    if backup:
        _backup(path)
    prs.save(path)


def _backup(path):
    bak = path + ".bak"
    if os.path.exists(path):
        shutil.copy2(path, bak)


# =============================================================================
# CLI
# =============================================================================
def main(argv=None):
    # надёжный вывод кириллицы в консоли Windows
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(
        description="Синхронизировать изменения между <N>_Лекция.md и <N>_Лекция.pptx")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_rep = sub.add_parser("report", help="изучить расхождения (ничего не менять)")
    p_rep.add_argument("md")
    p_rep.add_argument("pptx")
    p_rep.add_argument("--json", help="дополнительно записать JSON-сводку в файл")

    p_app = sub.add_parser("apply", help="применить изменения")
    p_app.add_argument("md")
    p_app.add_argument("pptx")
    p_app.add_argument("--direction", choices=["md2pptx", "pptx2md"], default="md2pptx")
    p_app.add_argument("--mode", choices=["add", "update", "all"], default="add")
    p_app.add_argument("--dry-run", action="store_true", help="только показать план")
    p_app.add_argument("--backup", action="store_true", help="сохранить .bak")
    args = ap.parse_args(argv)

    if args.cmd == "report":
        diff = build_diff(args.md, args.pptx)
        print(render_report(diff))
        if args.json:
            j = dict(mode=diff["mode"], md_count=diff["md_count"],
                     pptx_count=diff["pptx_count"],
                     only_md=[dict(num=u["num"], title=_short(u["title"] or u["text"]))
                              for u in diff["only_md"]],
                     only_pptx=[dict(num=u["num"], title=_short(u["title"] or u["text"]))
                                for u in diff["only_pptx"]],
                     changed_text=diff["changed_text"], changed_title=diff["changed_title"])
            with io.open(args.json, "w", encoding="utf-8") as f:
                json.dump(j, f, ensure_ascii=False, indent=2)
            print("\nJSON-сводка записана: {}".format(args.json))
        return 0

    if args.cmd == "apply":
        diff = build_diff(args.md, args.pptx)
        if args.dry_run:
            print("=== DRY RUN (изменения не записываются) ===")
            print(render_report(diff))
        if args.direction == "md2pptx":
            summary = apply_md2pptx(diff, args.md, args.pptx, args.mode,
                                    args.dry_run, args.backup, None)
        else:
            summary = apply_pptx2md(diff, args.md, args.pptx, args.mode,
                                    args.dry_run, args.backup)
        action = "ПЛАН" if args.dry_run else "ГОТОВО"
        print("{} ({}, mode={}): добавлено {added}, обновлено {updated}, "
              "удалено {deleted}".format(action, summary["direction"], summary["mode"], **summary))
        if not args.dry_run:
            tgt = args.pptx if args.direction == "md2pptx" else args.md
            print("Записано: {}".format(tgt))
        return 0


if __name__ == "__main__":
    sys.exit(main())
