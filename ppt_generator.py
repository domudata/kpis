# -*- coding: utf-8 -*-
"""
Générateur de PowerPoint à partir du JSON de l'analyse IA
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io

# Palette de couleurs industrielle
PRIMARY = RGBColor(30, 58, 95)      # Bleu marine foncé
ACCENT = RGBColor(37, 99, 235)      # Bleu vif
WHITE = RGBColor(255, 255, 255)
DARK_TEXT = RGBColor(30, 41, 59)
LIGHT_GRAY = RGBColor(241, 245, 249)
RED = RGBColor(220, 38, 38)
ORANGE = RGBColor(217, 119, 6)
GREEN = RGBColor(5, 150, 105)

def add_title_shape(slide, text, left, top, width, height, font_size=32, bold=True, color=WHITE):
    """Ajoute un texte avec fond coloré pour les titres de slides."""
    shape = slide.shapes.add_shape(1, left, top, width, height) # 1 = Rectangle
    shape.fill.solid()
    shape.fill.fore_color.rgb = PRIMARY
    shape.line.fill.background()
    
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = PP_ALIGN.LEFT
    tf.margin_left = Inches(0.5)

def create_pptx(json_data: dict) -> bytes:
    """Génère le fichier PowerPoint en mémoire et retourne les bytes."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6] # Layout vide

    # ==========================================
    # SLIDE 0 : TITRE
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Fond bleu
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = PRIMARY
    
    # Titre principal
    txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.333), Inches(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = json_data.get("title", "Analyse des KPI Maintenance")
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    # Sous-titre
    txBox2 = slide.shapes.add_textbox(Inches(1), Inches(3.8), Inches(11.333), Inches(2))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = json_data.get("subtitle", "")
    p2.font.size = Pt(24)
    p2.font.color.rgb = RGBColor(203, 213, 225)
    p2.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDES 1 à 3 : CONCLUSIONS
    # ==========================================
    slide_titles = [
        ("Slide 1 : Analyse des KPI Performance et Qualité", "slide1"),
        ("Slide 2 : Analyse du Tableau des Anomalies", "slide2"),
        ("Slide 3 : Analyse des Tendances", "slide3")
    ]

    for title, key in slide_titles:
        slide = prs.slides.add_slide(blank_layout)
        add_title_shape(slide, title, Inches(0), Inches(0), prs.slide_width, Inches(1.2))
        
        conclusion_text = json_data.get(key, {}).get("conclusion", "Aucune donnée.")
        
        # Boîte de texte pour la conclusion (grande et centrée)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.333), Inches(3))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"« {conclusion_text} »"
        p.font.size = Pt(32)
        p.font.color.rgb = PRIMARY
        p.font.italic = True
        p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 4 : CAUSES RACINES
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    add_title_shape(slide, "Slide 4 : Causes Racines Probables", Inches(0), Inches(0), prs.slide_width, Inches(1.2))
    
    causes = json_data.get("slide4", {}).get("root_causes", [])
    y_pos = Inches(1.8)
    
    for cause in causes:
        cause_text = cause.get("cause", "")
        criticite = cause.get("criticite", "Moyenne")
        
        # Déterminer la couleur de la criticité
        crit_color = ORANGE
        if criticite == "Critique": crit_color = RED
        elif criticite == "Faible": crit_color = GREEN
        
        # Pastille de couleur
        shape = slide.shapes.add_shape(9, Inches(0.8), y_pos + Inches(0.15), Inches(0.3), Inches(0.3)) # Oval
        shape.fill.solid()
        shape.fill.fore_color.rgb = crit_color
        shape.line.fill.background()
        
        # Texte de la cause
        txBox = slide.shapes.add_textbox(Inches(1.4), y_pos, Inches(10), Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cause_text
        p.font.size = Pt(24)
        p.font.color.rgb = DARK_TEXT
        
        # Texte criticité
        txBox2 = slide.shapes.add_textbox(Inches(11.5), y_pos, Inches(1.5), Inches(0.5))
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = criticite
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = crit_color
        p2.alignment = PP_ALIGN.RIGHT
        
        y_pos += Inches(1.0)

    # ==========================================
    # SLIDE 5 : PLAN D'ACTION (TABLEAU)
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    add_title_shape(slide, "Slide 5 : Plan d'Action", Inches(0), Inches(0), prs.slide_width, Inches(1.2))
    
    recs = json_data.get("slide5", {}).get("recommendations", [])
    if recs:
        rows, cols = len(recs) + 1, 5
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.5)).table
        
        # En-têtes
        headers = ["Action", "Responsable", "Priorité", "Échéance", "Impact attendu"]
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = PRIMARY
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(16)
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Données
        for r_idx, rec in enumerate(recs, 1):
            data = [
                rec.get("action", ""),
                rec.get("responsable", ""),
                rec.get("priorite", ""),
                rec.get("echeance", ""),
                rec.get("impact_attendu", "")
            ]
            for c_idx, text in enumerate(data):
                cell = table.cell(r_idx, c_idx)
                cell.text = str(text)
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(14)
                p.font.color.rgb = DARK_TEXT
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                
                # Couleur pour la priorité
                if c_idx == 2:
                    p.alignment = PP_ALIGN.CENTER
                    if text == "Haute": p.font.color.rgb = RED
                    elif text == "Moyenne": p.font.color.rgb = ORANGE
                    elif text == "Basse": p.font.color.rgb = GREEN
                    p.font.bold = True

    else:
        txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11.333), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "Aucune action générée."
        p.font.size = Pt(24)
        p.font.color.rgb = DARK_TEXT
        p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 6 : CONCLUSION EXÉCUTIVE
    # ==========================================
    slide = prs.slides.add_slide(blank_layout)
    
    # Fond dégradé simulé avec un gros rectangle
    bg_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = PRIMARY
    bg_shape.line.fill.background()
    
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11.333), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Conclusion Exécutive"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    conclusion = json_data.get("slide6", {}).get("final_conclusion", "")
    txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.333), Inches(4))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = f"« {conclusion} »"
    p2.font.size = Pt(32)
    p2.font.color.rgb = WHITE
    p2.font.italic = True
    p2.alignment = PP_ALIGN.CENTER

    # Sauvegarder en mémoire
    pptx_buffer = io.BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()
