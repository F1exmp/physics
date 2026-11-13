# -*- coding: utf-8 -*-
"""Генератор презентаций PowerPoint для уроков физики.

Использование:
    from tools.make_pptx import make_pptx
    slides = [
        ("title", "Заголовок презентации", "Подзаголовок"),
        ("section", "Название раздела"),
        ("bullets", "Заголовок слайда", ["Пункт 1", "Пункт 2", ...]),
        ("two_col", "Заголовок", ["левые пункты"], ["правые пункты"]),
        ("formula", "Заголовок", "Формула", "Пояснение"),
        ("task", "Задача № 1", "Условие задачи..."),          # слайд 1: условие
        ("solution", "Задача № 1 — решение", ["шаг 1", "шаг 2"], "Ответ: ..."),
        ("end", "Домашнее задание", ["п. 1, вопросы", ...]),
    ]
    make_pptx("path/to/file.pptx", slides, title_slide_footer="7 класс")
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

DARK = RGBColor(0x1F, 0x3B, 0x63)
ACCENT = RGBColor(0x2E, 0x74, 0xB5)
GRAY = RGBColor(0x44, 0x44, 0x44)
LIGHT = RGBColor(0xF2, 0xF5, 0xFA)

W, H = Inches(13.333), Inches(7.5)


def _bg(slide, color=LIGHT):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _title(slide, text, size=32):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), W - Inches(1), Inches(0.9))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = DARK
    return box


def _body(slide, items, left=0.7, top=1.4, width=11.9, size=20):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), H - Inches(top + 0.4))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        level = 0
        text = it
        if isinstance(it, tuple):
            level, text = it
        p.text = ("• " if level >= 0 else "") + text if level >= 0 else text
        p.level = max(level, 0)
        p.font.size = Pt(size - 2 * max(level, 0))
        p.font.color.rgb = GRAY
        p.space_after = Pt(8)
    return box


def _add(prs, layout_blank=True):
    return prs.slides.add_slide(prs.slide_layouts[6])


def make_pptx(path, slides, footer="7 класс"):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    for s in slides:
        kind = s[0]
        if kind == "title":
            sl = _add(prs)
            _bg(sl, DARK)
            box = sl.shapes.add_textbox(Inches(0.8), Inches(2.2), W - Inches(1.6), Inches(2.6))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s[1]
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            if len(s) > 2 and s[2]:
                p2 = tf.add_paragraph()
                p2.text = s[2]
                p2.font.size = Pt(22)
                p2.font.color.rgb = RGBColor(0xCF, 0xDD, 0xEE)
            if len(s) > 3 and s[3]:
                p3 = tf.add_paragraph()
                p3.text = s[3]
                p3.font.size = Pt(18)
                p3.font.color.rgb = RGBColor(0xAF, 0xC5, 0xE0)
        elif kind == "section":
            sl = _add(prs)
            _bg(sl, ACCENT)
            box = sl.shapes.add_textbox(Inches(0.8), Inches(2.8), W - Inches(1.6), Inches(2))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s[1]
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif kind == "bullets":
            sl = _add(prs)
            _bg(sl)
            _title(sl, s[1])
            _body(sl, list(s[2]))
        elif kind == "two_col":
            sl = _add(prs)
            _bg(sl)
            _title(sl, s[1])
            _body(sl, list(s[2]), top=1.5, width=5.6, size=18)
            _body(sl, list(s[3]), left=6.9, top=1.5, width=5.6, size=18)
        elif kind == "formula":
            sl = _add(prs)
            _bg(sl)
            _title(sl, s[1])
            fbox = sl.shapes.add_textbox(Inches(0.7), Inches(1.7), W - Inches(1.4), Inches(1.4))
            tf = fbox.text_frame
            p = tf.paragraphs[0]
            p.text = s[2]
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = DARK
            p.alignment = 1
            if len(s) > 3 and s[3]:
                _body(sl, s[3] if isinstance(s[3], list) else [s[3]], top=3.3, size=18)
        elif kind == "task":
            sl = _add(prs)
            _bg(sl)
            _title(sl, s[1])
            _body(sl, [s[2]], top=1.5, size=22)
        elif kind == "solution":
            sl = _add(prs)
            _bg(sl)
            _title(sl, s[1])
            _body(sl, list(s[2]), top=1.4, size=18)
            if len(s) > 3 and s[3]:
                abox = sl.shapes.add_textbox(Inches(0.7), Inches(6.2), W - Inches(1.4), Inches(1))
                tf = abox.text_frame
                p = tf.paragraphs[0]
                p.text = s[3]
                p.font.size = Pt(24)
                p.font.bold = True
                p.font.color.rgb = ACCENT
        elif kind == "end":
            sl = _add(prs)
            _bg(sl, DARK)
            box = sl.shapes.add_textbox(Inches(0.8), Inches(2.4), W - Inches(1.6), Inches(2.5))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s[1]
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            if len(s) > 2:
                for it in s[2]:
                    p2 = tf.add_paragraph()
                    p2.text = "• " + it
                    p2.font.size = Pt(20)
                    p2.font.color.rgb = RGBColor(0xCF, 0xDD, 0xEE)
        else:
            raise ValueError("unknown slide kind: " + kind)
    prs.save(path)
    return path
