#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create a beautiful PPTX presentation about Niklāvs Strunke
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def add_gradient_background(slide, color1, color2):
    """Add a gradient background to a slide"""
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 135
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[1].color.rgb = color2

def create_title_slide(prs):
    """Create title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    add_gradient_background(slide, RGBColor(102, 126, 234), RGBColor(118, 75, 162))

    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = "Niklāvs Strunke"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(60)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Subtitle 1
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(3.7), Inches(8), Inches(0.6))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Latviešu modernistu gleznotājs"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.italic = True
    subtitle_para.font.color.rgb = RGBColor(230, 230, 230)
    subtitle_para.alignment = PP_ALIGN.CENTER

    # Subtitle 2
    subtitle2_box = slide.shapes.add_textbox(Inches(1), Inches(4.4), Inches(8), Inches(0.6))
    subtitle2_frame = subtitle2_box.text_frame
    subtitle2_frame.text = "(1894–1966)"
    subtitle2_para = subtitle2_frame.paragraphs[0]
    subtitle2_para.font.size = Pt(28)
    subtitle2_para.font.italic = True
    subtitle2_para.font.color.rgb = RGBColor(230, 230, 230)
    subtitle2_para.alignment = PP_ALIGN.CENTER

def create_content_slide(prs):
    """Create table of contents slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Saturs"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Content items
    items = [
        "1. Biogrāfija",
        "2. Mākslinieciskā darbība",
        "3. Ieguldījums grāmatu ilustrācijā",
        "4. Mākslinieciskās grupas",
        "5. Pazīstamākie darbi",
        "6. Mantojums",
        "7. Avoti"
    ]

    y_pos = 1.8
    for item in items:
        # Create colored box for each item
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1), Inches(y_pos),
            Inches(8), Inches(0.6)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(248, 250, 252)
        shape.line.color.rgb = RGBColor(102, 126, 234)
        shape.line.width = Pt(3)

        text_frame = shape.text_frame
        text_frame.text = item
        text_frame.margin_left = Inches(0.2)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = text_frame.paragraphs[0]
        para.font.size = Pt(24)
        para.font.color.rgb = RGBColor(50, 50, 50)

        y_pos += 0.75

def create_biography_slide(prs):
    """Create biography slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Biogrāfija"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Timeline items
    timeline_data = [
        ("1894", "Dzimis 6. oktobrī Gostiņinā, Krievijas impērijā"),
        ("1909-1911", "Mācījās pie Nikolaja Rēriha un Ivana Bilibina Sanktpēterburgā"),
        ("1915", "Iestājās 6. Latviešu strēlnieku pulka izlūkošanas nodaļā"),
        ("1944", "Emigrēja uz Zviedriju"),
        ("1966", "Miris 13. oktobrī Romā")
    ]

    y_pos = 1.6
    for year, event in timeline_data:
        # Year box
        year_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(1.5), Inches(0.5))
        year_frame = year_box.text_frame
        year_frame.text = year
        year_para = year_frame.paragraphs[0]
        year_para.font.size = Pt(22)
        year_para.font.bold = True
        year_para.font.color.rgb = RGBColor(102, 126, 234)

        # Event box
        event_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.5), Inches(y_pos),
            Inches(6.5), Inches(0.5)
        )
        event_shape.fill.solid()
        event_shape.fill.fore_color.rgb = RGBColor(243, 244, 246)
        event_shape.line.color.rgb = RGBColor(118, 75, 162)
        event_shape.line.width = Pt(2)

        text_frame = event_shape.text_frame
        text_frame.text = event
        text_frame.margin_left = Inches(0.2)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = text_frame.paragraphs[0]
        para.font.size = Pt(16)
        para.font.color.rgb = RGBColor(50, 50, 50)

        y_pos += 0.8

def create_artistic_activity_slide(prs):
    """Create artistic activity slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Mākslinieciskā darbība"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Main text
    text_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(0.8))
    text_frame = text_box.text_frame
    text_frame.text = "Niklāvs Strunke bija viens no oriģinālākajiem latviešu modernistu paaudzes māksliniekiem."
    text_frame.word_wrap = True
    para = text_frame.paragraphs[0]
    para.font.size = Pt(20)
    para.font.color.rgb = RGBColor(50, 50, 50)

    # Left card - Genres
    left_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8), Inches(2.5),
        Inches(4), Inches(3.2)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = RGBColor(248, 250, 252)
    left_card.line.color.rgb = RGBColor(229, 231, 235)
    left_card.line.width = Pt(2)

    left_text = left_card.text_frame
    left_text.text = "Darbu žanri\n\n• Glezniecība\n• Grafika\n• Vitrāžas\n• Scenogrāfija"
    left_text.margin_left = Inches(0.3)
    left_text.margin_top = Inches(0.3)
    for i, para in enumerate(left_text.paragraphs):
        if i == 0:
            para.font.size = Pt(24)
            para.font.bold = True
            para.font.color.rgb = RGBColor(102, 126, 234)
        else:
            para.font.size = Pt(18)
            para.font.color.rgb = RGBColor(50, 50, 50)

    # Right card - Literary work
    right_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.2), Inches(2.5),
        Inches(4), Inches(3.2)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = RGBColor(248, 250, 252)
    right_card.line.color.rgb = RGBColor(229, 231, 235)
    right_card.line.width = Pt(2)

    right_text = right_card.text_frame
    right_text.text = "Literārā darbība\n\n• Rakstīja par mākslu\n• Publicējās ar pseidonīmu Pālmeņu Klāvs"
    right_text.margin_left = Inches(0.3)
    right_text.margin_top = Inches(0.3)
    for i, para in enumerate(right_text.paragraphs):
        if i == 0:
            para.font.size = Pt(24)
            para.font.bold = True
            para.font.color.rgb = RGBColor(102, 126, 234)
        else:
            para.font.size = Pt(18)
            para.font.color.rgb = RGBColor(50, 50, 50)

def create_book_illustration_slide(prs):
    """Create book illustration contribution slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Ieguldījums grāmatu ilustrācijā"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(8.4), Inches(0.6))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Periodā no 1920. līdz 1944. gadam Strunke sniedza milzīgu ieguldījumu latviešu grāmatu kultūrā"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(18)
    subtitle_para.font.color.rgb = RGBColor(80, 80, 80)

    # Left card - Book covers
    left_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.5), Inches(2.3),
        Inches(3.5), Inches(3)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = RGBColor(248, 250, 252)
    left_card.line.color.rgb = RGBColor(229, 231, 235)
    left_card.line.width = Pt(2)

    left_text = left_card.text_frame
    left_text.text = "Grāmatu vāki\n\n700+\n\nVairāk nekā 700 grāmatu vāki"
    left_text.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, para in enumerate(left_text.paragraphs):
        para.alignment = PP_ALIGN.CENTER
        if i == 0:
            para.font.size = Pt(26)
            para.font.bold = True
            para.font.color.rgb = RGBColor(102, 126, 234)
        elif i == 2:
            para.font.size = Pt(56)
            para.font.bold = True
            para.font.color.rgb = RGBColor(102, 126, 234)
        else:
            para.font.size = Pt(18)
            para.font.color.rgb = RGBColor(50, 50, 50)

    # Right card - Illustrations
    right_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.5), Inches(2.3),
        Inches(3.5), Inches(3)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = RGBColor(248, 250, 252)
    right_card.line.color.rgb = RGBColor(229, 231, 235)
    right_card.line.width = Pt(2)

    right_text = right_card.text_frame
    right_text.text = "Ilustrācijas\n\n~30\n\nAptuveni 30 romāni, dzejas un pasakas"
    right_text.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, para in enumerate(right_text.paragraphs):
        para.alignment = PP_ALIGN.CENTER
        if i == 0:
            para.font.size = Pt(26)
            para.font.bold = True
            para.font.color.rgb = RGBColor(118, 75, 162)
        elif i == 2:
            para.font.size = Pt(56)
            para.font.bold = True
            para.font.color.rgb = RGBColor(118, 75, 162)
        else:
            para.font.size = Pt(18)
            para.font.color.rgb = RGBColor(50, 50, 50)

def create_artistic_groups_slide(prs):
    """Create artistic groups slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Mākslinieciskās grupas"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(0.5))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Rīgas Mākslas grupa"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.bold = True
    subtitle_para.font.color.rgb = RGBColor(102, 126, 234)

    # Description
    desc_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.1), Inches(8.4), Inches(1.2))
    desc_frame = desc_box.text_frame
    desc_frame.text = "Strunke pievienojās inovatīvajai ekspresionistu grupai, kas orientējās uz franču mākslu un pēc Pirmā pasaules kara kļuva par Rīgas Mākslas grupu."
    desc_frame.word_wrap = True
    desc_para = desc_frame.paragraphs[0]
    desc_para.font.size = Pt(20)
    desc_para.font.color.rgb = RGBColor(50, 50, 50)

    # Style characteristics
    style_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.5), Inches(8.4), Inches(0.5))
    style_frame = style_box.text_frame
    style_frame.text = "Stila īpašības"
    style_para = style_frame.paragraphs[0]
    style_para.font.size = Pt(28)
    style_para.font.bold = True
    style_para.font.color.rgb = RGBColor(102, 126, 234)

    # Bullet points
    bullets = [
        "• Modernistiskie virzieni",
        "• Ekspresionisms",
        "• Franču ietekme",
        "• Oriģināls autorraksts"
    ]

    y_pos = 4.2
    for bullet in bullets:
        bullet_box = slide.shapes.add_textbox(Inches(1.2), Inches(y_pos), Inches(7.6), Inches(0.4))
        bullet_frame = bullet_box.text_frame
        bullet_frame.text = bullet
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.font.size = Pt(22)
        bullet_para.font.color.rgb = RGBColor(50, 50, 50)
        y_pos += 0.5

def create_famous_works_slide(prs):
    """Create famous works slide with image placeholder"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Pazīstamākie darbi"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # Image placeholder
    img_placeholder = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8), Inches(1.5),
        Inches(4.5), Inches(4)
    )
    img_placeholder.fill.solid()
    img_placeholder.fill.fore_color.rgb = RGBColor(200, 210, 230)
    img_placeholder.line.color.rgb = RGBColor(102, 126, 234)
    img_placeholder.line.width = Pt(3)

    placeholder_text = img_placeholder.text_frame
    placeholder_text.text = "[Vieta attēlam]\n\n\"Cilvēks, kas ieiet istabā\"\n(1927)"
    placeholder_text.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, para in enumerate(placeholder_text.paragraphs):
        para.alignment = PP_ALIGN.CENTER
        if i == 0:
            para.font.size = Pt(18)
            para.font.italic = True
            para.font.color.rgb = RGBColor(100, 100, 100)
        else:
            para.font.size = Pt(20)
            para.font.bold = True
            para.font.color.rgb = RGBColor(50, 50, 50)

    # Description
    desc_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.5), Inches(4), Inches(1.5))
    desc_frame = desc_box.text_frame
    desc_frame.text = "\"Cilvēks, kas ieiet istabā\" (1927)\n\nViena no slavenākajām mākslinieka gleznām, iekļauta Latvijas kultūras kanonā."
    desc_frame.word_wrap = True
    for i, para in enumerate(desc_frame.paragraphs):
        if i == 0:
            para.font.size = Pt(22)
            para.font.bold = True
            para.font.color.rgb = RGBColor(102, 126, 234)
        else:
            para.font.size = Pt(18)
            para.font.color.rgb = RGBColor(50, 50, 50)

    # Characteristics title
    char_title = slide.shapes.add_textbox(Inches(5.5), Inches(3.2), Inches(4), Inches(0.5))
    char_title_frame = char_title.text_frame
    char_title_frame.text = "Radošā darba raksturojums:"
    char_title_para = char_title_frame.paragraphs[0]
    char_title_para.font.size = Pt(20)
    char_title_para.font.bold = True
    char_title_para.font.color.rgb = RGBColor(102, 126, 234)

    # Characteristics bullets
    bullets = [
        "• Unikāls modernistiskais stils",
        "• Darbi Latvijas muzeju kolekcijās",
        "• Grafiskie darbi LU bibliotēkā",
        "• Izstādīti Tate galerijā"
    ]

    y_pos = 3.8
    for bullet in bullets:
        bullet_box = slide.shapes.add_textbox(Inches(5.7), Inches(y_pos), Inches(3.8), Inches(0.35))
        bullet_frame = bullet_box.text_frame
        bullet_frame.text = bullet
        bullet_frame.word_wrap = True
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.font.size = Pt(16)
        bullet_para.font.color.rgb = RGBColor(50, 50, 50)
        y_pos += 0.4

def create_legacy_slide(prs):
    """Create legacy slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Mantojums"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    # International recognition
    int_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(0.5))
    int_title_frame = int_title.text_frame
    int_title_frame.text = "Starptautiskā atzinība"
    int_title_para = int_title_frame.paragraphs[0]
    int_title_para.font.size = Pt(28)
    int_title_para.font.bold = True
    int_title_para.font.color.rgb = RGBColor(102, 126, 234)

    int_bullets = [
        "• Darbi izstādīti prestižās pasaules galerijās un muzejās",
        "• Pārstāvēts Tate galerijā (Lielbritānija)",
        "• Darbi pārdoti starptautiskos izsoles namos"
    ]

    y_pos = 2.1
    for bullet in int_bullets:
        bullet_box = slide.shapes.add_textbox(Inches(1.2), Inches(y_pos), Inches(7.6), Inches(0.4))
        bullet_frame = bullet_box.text_frame
        bullet_frame.text = bullet
        bullet_frame.word_wrap = True
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.font.size = Pt(20)
        bullet_para.font.color.rgb = RGBColor(50, 50, 50)
        y_pos += 0.5

    # Latvian culture significance
    lv_title = slide.shapes.add_textbox(Inches(0.8), Inches(3.7), Inches(8.4), Inches(0.5))
    lv_title_frame = lv_title.text_frame
    lv_title_frame.text = "Nozīme latviešu kultūrai"
    lv_title_para = lv_title_frame.paragraphs[0]
    lv_title_para.font.size = Pt(28)
    lv_title_para.font.bold = True
    lv_title_para.font.color.rgb = RGBColor(102, 126, 234)

    lv_bullets = [
        "• Viens no galvenajiem latviešu modernisma pārstāvjiem",
        "• Milzīga ietekme uz latviešu grāmatu grafiku",
        "• Darbi iekļauti Latvijas kultūras kanonā"
    ]

    y_pos = 4.3
    for bullet in lv_bullets:
        bullet_box = slide.shapes.add_textbox(Inches(1.2), Inches(y_pos), Inches(7.6), Inches(0.4))
        bullet_frame = bullet_box.text_frame
        bullet_frame.text = bullet
        bullet_frame.word_wrap = True
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.font.size = Pt(20)
        bullet_para.font.color.rgb = RGBColor(50, 50, 50)
        y_pos += 0.5

def create_sources_slide(prs):
    """Create sources slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(240, 240, 250), RGBColor(255, 255, 255))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Avoti"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(118, 75, 162)

    sources = [
        "Wikipedia: Niklāvs Strunke (en.wikipedia.org)",
        "Tate Gallery: Niklavs Strunke 1894–1966 (tate.org.uk)",
        "Latvijas Kultūras kanons: Glezna \"Cilvēks, kas ieiet istabā\" (kulturaskanons.lv)",
        "Rīgas Centrālā bibliotēka: Kultūras trešdienā – Niklāvs Strunke (rcb.lv)",
        "Art UK: Strunke, Niklāvs, 1894–1966 (artuk.org)",
        "MutualArt: Niklavs Strunke Artworks at Auction (mutualart.com)"
    ]

    y_pos = 1.6
    for source in sources:
        source_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8), Inches(y_pos),
            Inches(8.4), Inches(0.6)
        )
        source_shape.fill.solid()
        source_shape.fill.fore_color.rgb = RGBColor(248, 250, 252)
        source_shape.line.color.rgb = RGBColor(118, 75, 162)
        source_shape.line.width = Pt(2)

        text_frame = source_shape.text_frame
        text_frame.text = source
        text_frame.margin_left = Inches(0.2)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.word_wrap = True
        para = text_frame.paragraphs[0]
        para.font.size = Pt(16)
        para.font.color.rgb = RGBColor(50, 50, 50)

        y_pos += 0.75

def create_thank_you_slide(prs):
    """Create thank you slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, RGBColor(102, 126, 234), RGBColor(118, 75, 162))

    # Thank you text
    thank_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    thank_frame = thank_box.text_frame
    thank_frame.text = "Paldies par uzmanību!"
    thank_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    thank_para = thank_frame.paragraphs[0]
    thank_para.font.size = Pt(60)
    thank_para.font.bold = True
    thank_para.font.color.rgb = RGBColor(255, 255, 255)
    thank_para.alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Niklāvs Strunke (1894–1966)"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.italic = True
    subtitle_para.font.color.rgb = RGBColor(230, 230, 230)
    subtitle_para.alignment = PP_ALIGN.CENTER

def main():
    """Main function to create presentation"""
    print("Creating presentation about Niklāvs Strunke...")

    # Create presentation object
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Create all slides
    print("Creating slide 1: Title...")
    create_title_slide(prs)

    print("Creating slide 2: Contents...")
    create_content_slide(prs)

    print("Creating slide 3: Biography...")
    create_biography_slide(prs)

    print("Creating slide 4: Artistic Activity...")
    create_artistic_activity_slide(prs)

    print("Creating slide 5: Book Illustration...")
    create_book_illustration_slide(prs)

    print("Creating slide 6: Artistic Groups...")
    create_artistic_groups_slide(prs)

    print("Creating slide 7: Famous Works...")
    create_famous_works_slide(prs)

    print("Creating slide 8: Legacy...")
    create_legacy_slide(prs)

    print("Creating slide 9: Sources...")
    create_sources_slide(prs)

    print("Creating slide 10: Thank You...")
    create_thank_you_slide(prs)

    # Save presentation
    output_file = 'Niklavs_Strunke_Presentation.pptx'
    prs.save(output_file)
    print(f"\n✓ Presentation created successfully: {output_file}")
    print(f"✓ Total slides: 10")
    print("\nNote: Slide 7 has a placeholder for the painting 'Man Entering a Room' (1927).")
    print("You can add the image from: https://kulturaskanons.lv/en/archive/niklavs-strunke/")

if __name__ == "__main__":
    main()
