#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  🤖 AI COPILOT — Assistant IA Maintenance Industrielle                       ║
║  Page indépendante se connectant directement aux données de l'app.          ║
║  Exécution :  streamlit run ai_copilot.py                                   ║
║  Prérequis : pip install openai python-pptx python-docx                      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════
import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import json
import re
import html as html_lib
from datetime import datetime
from openai import OpenAI

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt as DocxPt, Inches as DocxInches, RGBColor as DocxRGB
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION DE LA PAGE
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(layout="wide", page_title="AI Copilot — Maintenance", page_icon="🤖", initial_sidebar_state="expanded")

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTES KPI
# ═══════════════════════════════════════════════════════════════════════════════
QK = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois","Performance Graissage","Performance Inspection","Performance Appels Systématiques"]
PK = ["Taux d'approbation des Avis","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL","OT Fiabilité","Total Avis de Panne"]
ALL_KPI = QK + PK

CIBLE = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,"OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,"OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,"OT exécution 1mois< <3mois":15,"Taux d'approbation des Avis":95,"OT LANC ESTIME":100,"Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,"OT CONFIME":100,"OT_COR_EGAL":100,"Performance Graissage":95,"Performance Inspection":95,"Performance Appels Systématiques":95,"OT Fiabilité":100,"Total Avis de Panne":100}

ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Améliorer le taux de réalisation des OT correctifs.","OT préparation <1 mois":"Réduire l'âge de préparation des OT (< 1 mois).","OT préparation >3 mois":"Traiter en priorité les OT avec préparation > 3 mois.","OT préparation 1mois< <3mois":"Réduire les OT dont la préparation est entre 1 et 3 mois.","OT planification <1 mois":"Réduire l'âge de planification des OT (< 1 mois).","OT planification >3 mois":"Traiter en priorité les OT avec planification > 3 mois.","OT planification 1mois< <3mois":"Réduire les OT dont la planification est entre 1 et 3 mois.","OT exécution <1 mois":"Réduire l'âge d'exécution des OT (< 1 mois).","OT exécution >3 mois":"Traiter en priorité les OT avec exécution > 3 mois.","OT exécution 1mois< <3mois":"Réduire les OT dont l'exécution est entre 1 et 3 mois.","Taux d'approbation des Avis":"Créer un OT pour chaque avis sans ordre associé.","OT LANC ESTIME":"Estimer les coûts de tous les OT lancés.","Backlog préparation caractérisé":"Caractériser l'intégralité du backlog de préparation.","Backlog planification caractérisé":"Caractériser l'intégralité du backlog de planification.","OT CONFIME":"Confirmer systématiquement les OT terminés.","OT_COR_EGAL":"Rapprocher les coûts réels et les coûts budgétés.","Performance Graissage":"Améliorer le taux de réalisation des OT de graissage (Type 350).","Performance Inspection":"Améliorer le taux de réalisation des OT d'inspection (Types 290, 300, 310).","Performance Appels Systématiques":"Améliorer le taux de réalisation des appels systématiques (Type 360).","OT Fiabilité":"Maintenir la fiabilité des coûts OT à 100%.","Total Avis de Panne":"Maintenir le suivi exhaustif des avis de panne."}

KPI_RESP_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Chef d'atelier","OT préparation <1 mois":"Préparateur BM","OT préparation 1mois< <3mois":"Préparateur BM","OT préparation >3 mois":"Préparateur BM","OT planification <1 mois":"Planificateur BM","OT planification 1mois< <3mois":"Planificateur BM","OT planification >3 mois":"Planificateur BM","OT exécution <1 mois":"Chef d'atelier","OT exécution 1mois< <3mois":"Chef d'atelier","OT exécution >3 mois":"Chef d'atelier","Taux d'approbation des Avis":"Chef d'atelier","OT LANC ESTIME":"Fiabilité","Backlog préparation caractérisé":"Préparateur BM","Backlog planification caractérisé":"Planificateur BM","OT CONFIME":"Agent de saisie","OT_COR_EGAL":"Agent de saisie","Performance Graissage":"Chef d'atelier","Performance Inspection":"Chef d'atelier","Performance Appels Systématiques":"Chef d'atelier","OT Fiabilité":"Fiabilité","Total Avis de Panne":"Fiabilité"}

LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois","OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]
MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
QUICK_SUGGESTIONS = ["🔍 Quels KPI sont critiques ?","⚡ Analyse les anomalies détectées","📊 Compare SF1 et SF2","🛠️ Propose un plan d'action","📉 Pourquoi les KPI de performance sont-ils faibles ?","📋 Quels sont les postes les plus problématiques ?","📈 Analyse les tendances des KPI","🎯 Résume la situation de la maintenance"]

# ═══════════════════════════════════════════════════════════════════════════════
#  CATALOGUE COMPLET DES MODÈLES IA
# ═══════════════════════════════════════════════════════════════════════════════
MODELS_CATALOG = {
    "── OpenAI ──": ["gpt-4o","gpt-4o-mini","gpt-4-turbo","gpt-4","gpt-3.5-turbo","o1-preview","o1-mini"],
    "── Anthropic ──": ["anthropic/claude-sonnet-4","anthropic/claude-3.5-sonnet","anthropic/claude-3.5-haiku","anthropic/claude-3-opus"],
    "── Google ──": ["google/gemini-pro-1.5","google/gemini-flash-1.5","google/gemini-2.0-flash-exp:free","google/gemma-2-9b-it:free"],
    "── Meta ──": ["meta-llama/llama-3.1-405b-instruct","meta-llama/llama-3.1-70b-instruct","meta-llama/llama-3.3-70b-instruct","meta-llama/llama-3.2-3b-instruct:free"],
    "── Mistral ──": ["mistralai/mistral-large","mistralai/mistral-small","mistralai/mixtral-8x22b-instruct","mistralai/mistral-7b-instruct:free"],
    "── Qwen ──": ["qwen/qwen-2.5-72b-instruct","qwen/qwen-2.5-32b-instruct","qwen/qwen-2.5-7b-instruct","qwen/qwen-2.5-coder-32b-instruct"],
    "── DeepSeek ──": ["deepseek/deepseek-chat","deepseek/deepseek-reasoner"],
    "── Autres ──": ["x-ai/grok-2-1212","cohere/command-r-plus","microsoft/phi-4"],
    "── 🆓 Gratuits ──": ["google/gemini-2.0-flash-exp:free","meta-llama/llama-3.2-3b-instruct:free","mistralai/mistral-7b-instruct:free","qwen/qwen-2.5-1.5b-instruct:free"]
}

# ═══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS UTILITAIRES — LECTURE & PRÉPARATION DES DONNÉES
# ═══════════════════════════════════════════════════════════════════════════════
def contient_mot(t, lm):
    t = str(t); return any(m in t for l in lm for m in l.split())

def cat_age(a):
    if pd.isna(a): return "Inconnu"
    if a <= 1: return "<1 mois"
    elif a >= 3: return ">3 mois"
    return "1 mois < <3 mois"

def excr(df):
    if "Poste travail princ." in df.columns:
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
    return df

@st.cache_data(show_spinner=False)
def read_excel_safe(bytes_data):
    bio = io.BytesIO(bytes_data)
    header = bytes_data[:8]
    if header[:4] in (b'PK\x03\x04', b'PK\x05\x06'):
        for engine in ['openpyxl', 'calamine']:
            try: return pd.read_excel(bio, engine=engine)
            except: bio.seek(0)
    if header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        for engine in ['xlrd', 'calamine']:
            try: return pd.read_excel(bio, engine=engine)
            except: bio.seek(0)
    for engine in ['openpyxl', 'xlrd', 'calamine']:
        try: bio.seek(0); return pd.read_excel(bio, engine=engine)
        except: continue
    raise ValueError("Format de fichier non reconnu.")

@st.cache_data(show_spinner=False)
def prepare_data(ot_bytes, av_bytes, date_str):
    raw_ot = read_excel_safe(ot_bytes); raw_av = read_excel_safe(av_bytes)
    raw_ot = excr(raw_ot); raw_av = excr(raw_av)
    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in raw_ot.columns: raw_ot[c] = pd.to_datetime(raw_ot[c], errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in raw_av.columns: raw_av[c] = pd.to_datetime(raw_av[c], errors="coerce")
    now_ts = pd.Timestamp.today(); df = raw_ot.copy()
    df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
    df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
    df["Type Carac Prep"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MP_KW if kw in str(x)), "NON CARACTERISE"))
    df["Type Carac Plan"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MPLAN_KW if kw in str(x)), "NON CARACTERISE"))
    for dc,am,ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
        if dc in df.columns:
            df[am]=((now_ts.year-df[dc].dt.year)*12+(now_ts.month-df[dc].dt.month)).round(2); df[ac]=df[am].apply(cat_age)
        else: df[am]=np.nan; df[ac]="Inconnu"
    df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLOT|TCLO",na=False) & df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
    df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    df["_tw_num"]=pd.to_numeric(df.get("Type de travail",pd.Series(dtype=float)),errors="coerce")
    if "Statut système" in df.columns: df["Statut OT"]=df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
    avf = raw_av[(raw_av["Ordre"].isna()|(raw_av["Ordre"].astype(str).str.strip()==""))&(raw_av["Type d'avis"].isin(["ZU","Z4","ZR","ZP"]))].copy()
    apm = sorted(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
    return df, avf, apm, now_ts

def detect_files():
    ot_path, av_path = None, None
    for f in os.listdir("."):
        if f.lower().endswith((".xlsx", ".xls")) and f != "indicateurs_kpis.xlsx":
            try:
                test = pd.read_excel(f, nrows=5)
                if "Ordre" in test.columns and "Poste travail princ." in test.columns and not ot_path: ot_path = f
                elif "Ordre" in test.columns and "Type d'avis" in test.columns and not av_path: av_path = f
            except: continue
    return ot_path, av_path

def get_date_from_file():
    for p in ["date.txt", "./date.txt"]:
        if os.path.exists(p):
            try: 
                with open(p,"r",encoding="utf-8") as f: return f.read().strip()
            except: pass
    return pd.Timestamp.today().strftime("%d/%m/%Y")

# ═══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS — CALCUL DES KPI
# ═══════════════════════════════════════════════════════════════════════════════
def calc_kpis(df_i, av_i, posts):
    res = {}; df = df_i.copy()
    def ckpi(n, d, sz=100): return np.where(d == 0, sz, (n / d) * 100)
    def cpiv(df_sub, flag, col, p): return pd.pivot_table(df_sub[flag], index="Poste travail princ.", columns=col, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)
    def statut_pivot(df_sub, p):
        piv = pd.pivot_table(df_sub, index="Poste travail princ.", columns="Statut OT", values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: piv[c] = piv.get(c, 0)
        piv["Realises"] = piv["CLOT"] + piv["TCLO"]; piv["Total"] = piv[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        return piv

    # 1. TAUX REALISATION
    filt_corr = (df["Nº appel pl.entret."].fillna(0)==0) & (df["Contient SOPL"]==1)
    an = cpiv(df, filt_corr, "Statut OT", posts)
    for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c] = an.get(c, 0)
    an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]; an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
    an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"])); res['an']=an

    # 2-4. PREPARATION
    pr = cpiv(df, (df["Statut OT"]=="CRÉÉ") & (df["Statut utilisateur"].str.contains(r"\bCRPR\b",case=False,na=False)), "ap", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pr[c]=pr.get(c,0)
    pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"]); pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0); pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0); res['pr']=pr

    # 5-7. PLANIFICATION
    pl = cpiv(df, (df["Statut OT"]=="LANC") & (df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)), "alp", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
    pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"]); pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0); pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0); res['pl']=pl

    # 8-10. EXECUTION
    ex = cpiv(df, (df["Statut OT"]=="LANC") & (df["Contient SOPL"]==1), "aex", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
    ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"]); ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0); ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0); res['ex']=ex

    # 11. OT LANC ESTIME
    la = pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: la[c]=la.get(c,0)
    la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"]); res['la']=la

    # 12-13. BACKLOGS
    pc = pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
    pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"]); res['pc']=pc

    pcl = pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pcl[c]=pcl.get(c,0)
    pcl["Total"]=pcl["CARACTERISE"]+pcl["NON CARACTERISE"]; pcl["Backlog planification caractérisé"]=ckpi(pcl["CARACTERISE"],pcl["Total"]); res['pcl']=pcl

    # 14-15. CONFIME & COR_EGAL
    cf = pd.pivot_table(df,index="Poste travail princ.",columns="OT CONFIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: cf[c]=cf.get(c,0)
    cf["Total"]=cf["OUI"]+cf["NON"]; cf["OT CONFIME"]=ckpi(cf["OUI"],cf["Total"]); res['cf']=cf

    ce = pd.pivot_table(df,index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: ce[c]=ce.get(c,0)
    ce["Total"]=ce["OUI"]+ce["NON"]; ce["OT_COR_EGAL"]=ckpi(ce["OUI"],ce["Total"]); res['ce']=ce

    # 16-18. PERFORMANCES TYPE
    gra = statut_pivot(df[df["_tw_num"]==350], posts); gra["Performance Graissage"]=ckpi(gra["Realises"],gra["Total"]); res['gra']=gra
    ins = statut_pivot(df[df["_tw_num"].isin([290,300,310])], posts); ins["Performance Inspection"]=ckpi(ins["Realises"],ins["Total"]); res['ins']=ins
    app_sys = statut_pivot(df[df["_tw_num"]==360], posts); app_sys["Performance Appels Systématiques"]=ckpi(app_sys["Realises"],app_sys["Total"]); res['app_sys']=app_sys

    # 19. FIABILITE
    fiab = pd.pivot_table(df[df["Statut OT"].isin(["CLOT","TCLO"])],index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: fiab[c]=fiab.get(c,0)
    fiab["Total"]=fiab["OUI"]+fiab["NON"]; fiab["OT Fiabilité"]=ckpi(fiab["OUI"],fiab["Total"]); res['fiab']=fiab

    # 20. TAUX APPROBATION AVIS
    av_by_poste = av_i.groupby("Poste travail princ." if "Poste travail princ." in av_i.columns else av_i.columns[0]).size().reindex(posts, fill_value=0)
    total_av = df.groupby("Poste travail princ.").size().reindex(posts, fill_value=0)
    taux_approb = pd.Series(np.where((total_av+av_by_poste)==0,100.0,ckpi(total_av,total_av+av_by_poste)),index=posts,name="Taux d'approbation des Avis")
    res['taux_approb']=taux_approb; res['av_count']=av_by_poste; res['dfp']=df

    # ASSEMBLAGE
    perf_data = {"TAUX_REALISATION_CORRECTIF/PT":res['an']["TAUX_REALISATION_CORRECTIF/PT"],"OT préparation <1 mois":res['pr']["OT préparation <1 mois"],"OT préparation >3 mois":res['pr']["OT préparation >3 mois"],"OT préparation 1mois< <3mois":res['pr']["OT préparation 1mois< <3mois"],"OT planification <1 mois":res['pl']["OT planification <1 mois"],"OT planification >3 mois":res['pl']["OT planification >3 mois"],"OT planification 1mois< <3mois":res['pl']["OT planification 1mois< <3mois"],"OT exécution <1 mois":res['ex']["OT exécution <1 mois"],"OT exécution >3 mois":res['ex']["OT exécution >3 mois"],"OT exécution 1mois< <3mois":res['ex']["OT exécution 1mois< <3mois"],"Performance Graissage":res['gra']["Performance Graissage"],"Performance Inspection":res['ins']["Performance Inspection"],"Performance Appels Systématiques":res['app_sys']["Performance Appels Systématiques"]}
    res['perf_df'] = pd.DataFrame(perf_data, index=posts)
    
    qual_data = {"Taux d'approbation des Avis":taux_approb,"OT LANC ESTIME":res['la']["OT LANC ESTIME"],"Backlog préparation caractérisé":res['pc']["Backlog préparation caractérisé"],"Backlog planification caractérisé":res['pcl']["Backlog planification caractérisé"],"OT CONFIME":res['cf']["OT CONFIME"],"OT_COR_EGAL":res['ce']["OT_COR_EGAL"],"OT Fiabilité":res['fiab']["OT Fiabilité"],"Total Avis de Panne":100.0}
    res['qual_df'] = pd.DataFrame(qual_data, index=posts)

    def calc_score(df_kpi, kpi_list):
        scores = []
        for poste in df_kpi.index:
            s, count = 0, 0
            for kpi in kpi_list:
                if kpi in df_kpi.columns and kpi in CIBLE:
                    val = df_kpi.loc[poste, kpi]; cible = CIBLE[kpi]
                    if cible > 0:
                        ratio = val / cible
                        if kpi in LOWER_BETTER: s += min(ratio, 1.0) * 100 if ratio <= 1.0 else max(0, 2.0 - ratio) * 100
                        else: s += min(ratio, 1.0) * 100
                    else: s += 100
                    count += 1
            scores.append(s / count if count > 0 else 0)
        return pd.Series(scores, index=df_kpi.index)

    res['perf_df']['Score Performance'] = calc_score(res['perf_df'], QK)
    res['qual_df']['Score Qualite'] = calc_score(res['qual_df'], PK)
    return res

# ═══════════════════════════════════════════════════════════════════════════════
#  HISTORIQUE & CONTEXTE IA
# ═══════════════════════════════════════════════════════════════════════════════
def load_historical_kpis(filepath):
    if not os.path.exists(filepath): return pd.DataFrame()
    try:
        from openpyxl import load_workbook
        wb = load_workbook(filepath, read_only=True, data_only=True)
    except: return pd.DataFrame()
    records, section, headers = [], None, None
    for sheet_name in wb.sheetnames:
        try:
            ws = wb[sheet_name]
            for row in ws.iter_rows(values_only=True):
                cell0 = str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper(): section="perf"; headers=None; continue
                elif "INDICATEURS DE QUALITE" in cell0.upper(): section="qual"; headers=None; continue
                elif "ANOMALIES" in cell0.upper(): section=None; continue
                if section and headers is None and cell0: headers=[str(c).strip() if c else "" for c in row]; continue
                if section and headers and cell0 and cell0 not in ("Cible","Total general",""):
                    entry={"Date":sheet_name}
                    for j,h in enumerate(headers):
                        if j<len(row): entry[h]=row[j]
                    entry["_section"]=section; records.append(entry)
        except: continue
    wb.close()
    if not records: return pd.DataFrame()
    df = pd.DataFrame(records)
    df["Date_parsed"]=pd.to_datetime(df["Date"].str.replace("-","/"),format="%d/%m/%Y",errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)

def calculate_variations(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns: return pd.DataFrame()
    dates = sorted(hist_df["Date"].unique())
    if len(dates) < 2: return pd.DataFrame()
    variations = []
    for i in range(1, len(dates)):
        prev_date, curr_date = dates[i-1], dates[i]
        for sec_val in ["Performance", "Qualite"]:
            prev_d = hist_df[(hist_df["Date"]==prev_date) & (hist_df["_section"]==sec_val[0].lower())]
            curr_d = hist_df[(hist_df["Date"]==curr_date) & (hist_df["_section"]==sec_val[0].lower())]
            if "Poste de travail" not in prev_d.columns or "Poste de travail" not in curr_d.columns: continue
            prev_idx, curr_idx = prev_d.set_index("Poste de travail"), curr_d.set_index("Poste de travail")
            score_col = f"Score {sec_val}"
            if score_col not in prev_idx.columns or score_col not in curr_idx.columns: continue
            for poste in set(prev_idx.index) & set(curr_idx.index):
                try: pv, cv = float(prev_idx.loc[poste, score_col]), float(curr_idx.loc[poste, score_col])
                except: continue
                diff, pct = cv - pv, ((cv-pv)/pv*100) if pv != 0 else 0
                sens = "Stable" if abs(diff) <= 0.5 else ("Amelioration" if diff > 0 else "Degradation")
                variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,"Type":sec_val,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),"Ecart":round(diff,2),"Ecart %":round(pct,2),"Sens":sens})
    return pd.DataFrame(variations)

def build_ai_context(kpis, posts, date_str, hist_df, var_df):
    ctx = [f"=== DATE DES DONNÉES : {date_str} ===\n"]
    df = kpis.get('dfp', pd.DataFrame())
    ctx.append(f"Nombre total d'OT chargés : {len(df)}")
    av_count = kpis.get('av_count', pd.Series(dtype=int))
    ctx.append(f"Nombre d'avis sans ordre : {int(av_count.sum())}")
    ctx.append(f"Postes analysés : {', '.join(posts)}\n")

    def format_table(df_kpi, kpi_list, title):
        ctx.append(f"=== {title} ===")
        header = "Poste | " + " | ".join([f"{k} (cible:{CIBLE.get(k,'?')})" for k in kpi_list])
        ctx.append(header); ctx.append("-"*len(header))
        for poste in df_kpi.index:
            vals = [f"{df_kpi.loc[poste, k]:.1f}" if k in df_kpi.columns else "N/A" for k in kpi_list]
            ctx.append(f"{poste} | " + " | ".join(vals))
        means = [f"{df_kpi[k].mean():.1f}" if k in df_kpi.columns else "N/A" for k in kpi_list]
        ctx.append(f"MOYENNE | " + " | ".join(means) + "\n")

    perf_df, qual_df = kpis.get('perf_df', pd.DataFrame()), kpis.get('qual_df', pd.DataFrame())
    if not perf_df.empty: format_table(perf_df, QK, "KPI DE PERFORMANCE")
    if not qual_df.empty: format_table(qual_df, PK, "KPI DE QUALITÉ")

    ctx.append("=== ANOMALIES DÉTECTÉES (KPI sous cible) ===")
    anomaly_count = 0
    all_kpi_df = {**{k: perf_df for k in QK}, **{k: qual_df for k in PK}}
    for kpi_name, kpi_df in all_kpi_df.items():
        if kpi_df.empty or kpi_name not in kpi_df.columns: continue
        cible = CIBLE.get(kpi_name)
        if cible is None: continue
        for poste in kpi_df.index:
            val = kpi_df.loc[poste, kpi_name]
            is_anomaly = (val > cible) if kpi_name in LOWER_BETTER else (val < cible)
            if is_anomaly:
                anomaly_count += 1
                ctx.append(f"• {poste} — {kpi_name} : {val:.1f}% (cible: {cible}%, écart: {val-cible:+.1f} pts) → Responsable: {KPI_RESP_MAP.get(kpi_name,'N/A')} → Action: {ACT_MAP.get(kpi_name,'N/A')}")
    if anomaly_count == 0: ctx.append("Aucune anomalie détectée.")
    ctx.append(f"\nTotal anomalies : {anomaly_count}\n")
    if not var_df.empty:
        ctx.append("=== TENDANCES ===")
        for _, row in var_df.iterrows(): ctx.append(f"  {row['Poste']} — {row['Type']} : {row['Valeur precedente']:.1f} → {row['Valeur actuelle']:.1f} ({row['Sens']})")
    return "\n".join(ctx)

# ═══════════════════════════════════════════════════════════════════════════════
#  IA — PROMPTS & APPEL API
# ═══════════════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT = """Tu es **AI Copilot**, un assistant IA expert en maintenance industrielle spécialisé SAP PM.
## RÈGLES STRICTES
1. Utilise UNIQUEMENT les données fournies. Ne JAMAIS inventer de chiffres.
2. Si une information n'est pas dans le contexte, dis-le : *"Je n'ai pas cette information."*
3. Réponds en **français**. Cite les **chiffres exacts**. Structure avec **Markdown** (titres, tableaux, listes).
4. Sois précis, concis et professionnel. Quand tu identifies un problème, propose une **action concrète**.
5. Pour les KPI "OT >3 mois" et "OT 1-3 mois" : plus c'est bas, mieux c'est. Pour les autres taux, plus c'est haut, mieux c'est.
6. Un KPI est 🔴 Critique si >10 pts sous cible, 🟡 Alerte si 5-10 pts sous cible, 🟢 Conforme sinon.
"""

REPORT_PROMPT = """Génère un rapport professionnel structuré en Markdown avec EXACTEMENT ces sections :
## 1. Résumé Exécutif (5-8 lignes)
## 2. Analyse des KPI de Performance (Tableau + analyse)
## 3. Analyse des KPI de Qualité (Tableau + analyse)
## 4. Analyse des Anomalies (🔴🟡 sévérité, écart, responsable)
## 5. Analyse des Causes (pour chaque anomalie majeure)
## 6. Analyse des Risques (sécurité, production, coûts)
## 7. Recommandations (numérotées, avec responsable et délai)
## 8. Plan d'Action (Tableau: Action | Responsable | Priorité | Délai)
## 9. Conclusion
Chaque affirmation doit être étayée par un chiffre du contexte."""

PPTX_PROMPT = """Génère le contenu pour 8 diapositives PowerPoint. Format strict :
===SLIDE===
TITRE: [titre]
CONTENU:
[liste à puces]
===SLIDE===
Diapositives : 1.Titre et date, 2.Contexte, 3.Performance, 4.Qualité, 5.Anomalies, 6.Causes, 7.Plan d'action, 8.Conclusion. Contenu factuel et percutant."""

def ask_ai(question, context, history, api_key, base_url, model, temperature, max_tokens):
    messages = [{"role": "system", "content": SYSTEM_PROMPT + "\n\n## DONNÉES DE CONTEXTE\n\n" + context}]
    for msg in history[-10:]:
        if msg["role"] in ("user", "assistant"): messages.append({"role": msg["role"], "content": msg["content"]})
    messages.append({"role": "user", "content": question})
    try:
        client_kwargs = {"base_url": base_url, "api_key": api_key if api_key and api_key.strip() else "not-needed"}
        client = OpenAI(**client_kwargs)
        response = client.chat.completions.create(model=model, messages=messages, temperature=temperature, max_tokens=max_tokens)
        return response.choices[0].message.content
    except Exception as e:
        return f"❌ **Erreur API** : {str(e)}\n\nVérifiez votre clé API et l'URL dans la barre latérale."

# ═══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION RAPPORT & POWERPOINT
# ═══════════════════════════════════════════════════════════════════════════════
def create_report_docx(content, date_str):
    if not DOCX_AVAILABLE: return None
    try:
        doc = DocxDocument()
        for line in content.split('\n'):
            s = line.strip()
            if not s: doc.add_paragraph('')
            elif s.startswith('## '):
                p = doc.add_heading(s[3:], level=2)
                for run in p.runs: run.font.color.rgb = DocxRGB(0x1E, 0x3A, 0x5F)
            elif s.startswith('### '): doc.add_heading(s[4:], level=3)
            elif re.match(r'^\d+\.', s): doc.add_paragraph(s, style='List Number')
            elif s.startswith('- ') or s.startswith('* '): doc.add_paragraph(s[2:], style='List Bullet')
            else: doc.add_paragraph(s)
        buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0); return buffer
    except: return None

def create_powerpoint(content, date_str):
    if not PPTX_AVAILABLE: return None
    try:
        prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
        PRIMARY, WHITE, LIGHT_BG, ACCENT, DARK = RGBColor(0x1E,0x3A,0x5F), RGBColor(0xFF,0xFF,0xFF), RGBColor(0xF0,0xF4,0xF8), RGBColor(0x25,0x63,0xEB), RGBColor(0x1E,0x29,0x3B)
        slides_data = [s.strip() for s in re.split(r'===SLIDE===', content) if s.strip()]
        for idx, slide_text in enumerate(slides_data):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = LIGHT_BG
            title, body_lines = "", []
            for line in slide_text.split('\n'):
                l = line.strip()
                if l.upper().startswith('TITRE:'): title = l[6:].strip()
                elif l.upper().startswith('CONTENU:'): continue
                elif l: body_lines.append(l)
            if title:
                tb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
                tb.fill.solid(); tb.fill.fore_color.rgb = PRIMARY; tb.line.fill.background()
                txBox = slide.shapes.add_textbox(Inches(0.6), Inches(0.15), Inches(12.5), Inches(1.0))
                p = txBox.text_frame.paragraphs[0]; p.text = title; p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = WHITE
            if body_lines:
                txBox = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12.2), Inches(5.5))
                tf = txBox.text_frame; tf.word_wrap = True
                for i, line in enumerate(body_lines):
                    clean = re.sub(r'[🔴🟡🟢⚠️✅❌📊📋🎯📈📉🔍⚡🛠️*#`]', '', line).strip()
                    if not clean: continue
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = ("▸  " + clean[2:]) if clean.startswith('- ') or clean.startswith('• ') else clean
                    p.font.size = Pt(18); p.font.color.rgb = DARK; p.space_after = Pt(8)
        buffer = io.BytesIO(); prs.save(buffer); buffer.seek(0); return buffer
    except: return None

# ═══════════════════════════════════════════════════════════════════════════════
#  CSS & INTERFACE
# ═══════════════════════════════════════════════════════════════════════════════
def inject_copilot_css():
    st.markdown("""<style>
    [data-testid="stHeaderActionElements"], [data-testid="stActionButtonContainer"] { display: none !important; }
    .stApp { background: #f0f4f8; }
    div[data-testid="stSidebar"] { background: linear-gradient(180deg, #1e40af 0%, #1e3a8a 50%, #1e3a5f 100%) !important; }
    div[data-testid="stSidebar"] * { color: rgba(255,255,255,0.9) !important; }
    div[data-testid="stSidebar"] .stSelectbox label, div[data-testid="stSidebar"] .stTextInput label, div[data-testid="stSidebar"] .stFileUploader label { color: rgba(255,255,255,0.9) !important; font-weight: 600; font-size: 13px; text-transform: uppercase; }
    div[data-testid="stSidebar"] div[data-testid="stWidget"] { background: rgba(255,255,255,0.1); border-radius: 6px; padding: 5px 10px; margin-bottom: 5px; border: 1px solid rgba(255,255,255,0.15); }
    div[data-testid="stSidebar"] .stSelectbox > div > div, div[data-testid="stSidebar"] .stTextInput > div > div { background: rgba(255,255,255,0.95) !important; border-radius: 5px; }
    .copilot-header { background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 60%, #3b82f6 100%); padding: 24px 32px; border-radius: 14px; margin-bottom: 12px; box-shadow: 0 8px 32px rgba(30,58,95,0.2); display: flex; align-items: center; gap: 16px; }
    .copilot-header h1 { color: #fff; font-size: 32px; font-weight: 800; margin: 0; flex: 1; }
    .copilot-header .badge { background: rgba(255,255,255,0.2); padding: 6px 16px; border-radius: 20px; color: #fff; font-size: 13px; font-weight: 700; border: 1px solid rgba(255,255,255,0.3); }
    .copilot-desc { color: #64748b; font-size: 15px; margin-bottom: 12px; }
    [data-testid="stChatMessage"] { padding: 16px 20px !important; font-size: 14.5px; line-height: 1.65; }
    [data-testid="stChatMessage"] table { font-size: 13px; width: 100%; border-collapse: collapse; margin: 10px 0; }
    [data-testid="stChatMessage"] table th { background: #1e3a5f; color: #fff; padding: 8px 12px; font-weight: 700; text-align: left; font-size: 12px; }
    [data-testid="stChatMessage"] table td { padding: 6px 12px; border-bottom: 1px solid #e2e8f0; }
    [data-testid="stChatMessage"] table tr:nth-child(even) td { background: #f8fafc; }
    [data-testid="stChatMessage"] h2, [data-testid="stChatMessage"] h3 { color: #1e3a5f; margin-top: 16px; }
    </style>""", unsafe_allow_html=True)

def render_sidebar():
    st.sidebar.markdown("## ⚙️ Configuration")
    with st.sidebar.expander("📂 Sources de données", expanded=True):
        source_mode = st.radio("Mode", ["Auto-détection", "Upload manuel"], label_visibility="collapsed", key="source_mode")
        ot_file, av_file = None, None
        if source_mode == "Auto-détection":
            ot_path, av_path = detect_files()
            if ot_path: st.success(f"✅ OT : `{ot_path}`"); ot_file = open(ot_path, "rb").read()
            else: st.warning("❌ Fichier OT non trouvé")
            if av_path: st.success(f"✅ Avis : `{av_path}`"); av_file = open(av_path, "rb").read()
            else: st.warning("❌ Fichier Avis non trouvé")
            if not ot_path or not av_path: st.info("💡 Placez les fichiers .xlsx à côté de ce script.")
        else:
            ot_up = st.file_uploader("Fichier OT", type=["xlsx","xls"], key="ot_upload")
            av_up = st.file_uploader("Fichier Avis", type=["xlsx","xls"], key="av_upload")
            if ot_up: ot_file = ot_up.read(); st.success(f"✅ OT chargé")
            if av_up: av_file = av_up.read(); st.success(f"✅ Avis chargé")

    with st.sidebar.expander("🔑 Configuration IA", expanded=True):
        api_key = st.text_input("Clé API", type="password", value=st.session_state.get("saved_api_key",""), key="api_key_input")
        st.session_state.saved_api_key = api_key
        
        providers = {"🌐 OpenAI":"https://api.openai.com/v1","🔀 OpenRouter":"https://openrouter.ai/api/v1","🖥️ Ollama (Local)":"http://localhost:11434/v1","🖥️ LM Studio (Local)":"http://localhost:1234/v1","⚙️ Personnalisée":"custom"}
        provider = st.selectbox("Prestataire", list(providers.keys()), key="provider_select")
        base_url = providers[provider] if providers[provider] != "custom" else st.text_input("URL Base", value=st.session_state.get("saved_base_url",""), key="base_url_input")
        st.session_state.saved_base_url = base_url

        use_custom = st.toggle("✏️ Saisie manuelle du modèle", value=False)
        if use_custom:
            model = st.text_input("Modèle", value=st.session_state.get("saved_model","gpt-4o"), key="custom_model_input")
        else:
            display_labels, label_to_model = [], {}
            for section, models in MODELS_CATALOG.items():
                display_labels.append(section)
                for m in models:
                    tag = " 🆓" if ":free" in m else ""
                    display_labels.append(f"  {m}{tag}"); label_to_model[f"  {m}{tag}"] = m
            selected_label = st.selectbox("Modèle", display_labels, index=1, key="model_select")
            model = label_to_model.get(selected_label, next(iter(MODELS_CATALOG.values()))[0] if selected_label in MODELS_CATALOG else "gpt-4o")
        st.session_state.saved_model = model
        
        if "localhost" in base_url: st.caption("🖥️ Mode local détecté (Clé API optionnelle)")
        temperature = st.slider("Température", 0.0, 1.0, 0.3, 0.05)
        max_tokens = st.slider("Max Tokens", 256, 8192, 4096, 256)

    with st.sidebar.expander("📊 État", expanded=False):
        if st.session_state.get("data_loaded"):
            kpis, posts = st.session_state.kpis, st.session_state.posts
            st.success("✅ Données prêtes")
            st.metric("OT", len(kpis.get('dfp', pd.DataFrame())))
            st.metric("Postes", len(posts))
        else: st.error("❌ Aucune donnée")

    return ot_file, av_file, api_key, base_url, model, temperature, max_tokens

def _process_question(question, context, api_key, base_url, model, temperature, max_tokens):
    st.session_state.copilot_messages.append({"role": "user", "content": question})
    with st.spinner("🤖 Analyse en cours..."):
        response = ask_ai(question, context, st.session_state.copilot_messages, api_key, base_url, model, temperature, max_tokens)
    st.session_state.copilot_messages.append({"role": "assistant", "content": response})

def render_main():
    inject_copilot_css()
    st.markdown("""<div class="copilot-header"><span style="font-size:44px">🤖</span><h1>AI Copilot</h1><span class="badge">Maintenance Industrielle</span></div>""", unsafe_allow_html=True)
    st.markdown('<p class="copilot-desc">💬 Posez vos questions concernant vos données maintenance.</p>', unsafe_allow_html=True)

    if not st.session_state.get("data_loaded"):
        st.warning("⚠️ Chargez les données dans la barre latérale."); return

    if not st.session_state.get("copilot_messages"):
        cols = st.columns(4)
        for i, sug in enumerate(QUICK_SUGGESTIONS):
            with cols[i % 4]:
                if st.button(sug, key=f"sug_{i}", use_container_width=True):
                    st.session_state.pending_question = sug.replace(sug[:2], "").strip(); st.rerun()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("🗑️ Effacer", use_container_width=True): st.session_state.copilot_messages = []; st.rerun()
    with c2:
        if st.button("📄 Générer Rapport", use_container_width=True, type="primary"): st.session_state.generate_report = True; st.rerun()
    with c3:
        if st.button("📽️ Générer PowerPoint", use_container_width=True, type="primary", disabled=not PPTX_AVAILABLE): st.session_state.generate_pptx = True; st.rerun()
    with c4:
        if st.button("📊 Résumé rapide", use_container_width=True): st.session_state.pending_question = "Donne-moi un résumé complet avec les chiffres clés."; st.rerun()

    for msg in st.session_state.copilot_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if "download" in msg:
                dl = msg["download"]
                st.download_button(label=dl["label"], data=dl["data"], file_name=dl["filename"], mime=dl["mime"], key=f"dl_{id(dl)}")

    api_key, base_url = st.session_state.get("saved_api_key",""), st.session_state.get("saved_base_url","https://api.openai.com/v1")
    model, temperature = st.session_state.get("current_model","gpt-4o"), st.session_state.get("current_temperature",0.3)
    max_tokens, context = st.session_state.get("current_max_tokens",4096), st.session_state.get("ai_context","")

    if st.session_state.get("generate_report"):
        st.session_state.generate_report = False
        if not api_key and "localhost" not in base_url: st.error("❌ Clé API requise.")
        else:
            with st.spinner("📝 Génération du rapport..."):
                st.session_state.copilot_messages.append({"role":"user","content":"📄 Génère un rapport complet."})
                report = ask_ai(REPORT_PROMPT, context, st.session_state.copilot_messages, api_key, base_url, model, 0.2, max_tokens)
                docx_buf = create_report_docx(report, st.session_state.date_str)
                dl_info = {"label":"📥 Télécharger le rapport (.docx)","data":docx_buf if docx_buf else report.encode('utf-8'),"filename":f"rapport_maintenance_{st.session_state.date_str.replace('/','-')}.{'docx' if docx_buf else 'md'}","mime":"application/vnd.openxmlformats-officedocument.wordprocessingml.document" if docx_buf else "text/markdown"}
                st.session_state.copilot_messages.append({"role":"assistant","content":report,"download":dl_info})
            st.rerun()

    if st.session_state.get("generate_pptx"):
        st.session_state.generate_pptx = False
        if not api_key and "localhost" not in base_url: st.error("❌ Clé API requise.")
        else:
            with st.spinner("📽️ Génération du PowerPoint..."):
                st.session_state.copilot_messages.append({"role":"user","content":"📽️ Génère une présentation PowerPoint."})
                pptx_content = ask_ai(PPTX_PROMPT, context, st.session_state.copilot_messages, api_key, base_url, model, 0.2, max_tokens)
                pptx_buf = create_powerpoint(pptx_content, st.session_state.date_str)
                if pptx_buf: st.session_state.copilot_messages.append({"role":"assistant","content":"✅ Présentation générée.","download":{"label":"📥 Télécharger (.pptx)","data":pptx_buf,"filename":f"presentation_kpi_{st.session_state.date_str.replace('/','-')}.pptx","mime":"application/vnd.openxmlformats-officedocument.presentationml.presentation"}})
                else: st.session_state.copilot_messages.append({"role":"assistant","content":"❌ Erreur de génération PPTX."})
            st.rerun()

    if pending := st.session_state.get("pending_question"):
        st.session_state.pending_question = None
        if not api_key and "localhost" not in base_url: st.error("❌ Clé API requise.")
        else: _process_question(pending, context, api_key, base_url, model, temperature, max_tokens); st.rerun()

    if prompt := st.chat_input("Posez votre question...", key="copilot_input", disabled=not st.session_state.get("data_loaded")):
        if not api_key and "localhost" not in base_url: st.error("❌ Clé API requise.")
        else: _process_question(prompt, context, api_key, base_url, model, temperature, max_tokens); st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    for k, v in [("copilot_messages",[]),("data_loaded",False),("generate_report",False),("generate_pptx",False),("pending_question",None)]:
        if k not in st.session_state: st.session_state[k] = v

    ot_file, av_file, api_key, base_url, model, temperature, max_tokens = render_sidebar()
    st.session_state.current_model, st.session_state.current_temperature, st.session_state.current_max_tokens = model, temperature, max_tokens

    if ot_file and av_file and not st.session_state.data_loaded:
        with st.spinner("🔄 Chargement des données..."):
            try:
                date_str = get_date_from_file(); st.session_state.date_str = date_str
                df, avf, posts, now_ts = prepare_data(ot_file, av_file, date_str)
                if not posts: st.error("❌ Aucun poste SF1/SF2 détecté."); return
                kpis = calc_kpis(df, avf, posts)
                st.session_state.kpis, st.session_state.posts, st.session_state.now_ts = kpis, posts, now_ts
                st.session_state.data_loaded = True
                hist_path = "kpis/indicateurs_kpis.xlsx"
                st.session_state.hist_df = load_historical_kpis(hist_path) if os.path.exists(hist_path) else pd.DataFrame()
                st.session_state.var_df = calculate_variations(st.session_state.hist_df)
                st.session_state.ai_context = build_ai_context(kpis, posts, date_str, st.session_state.hist_df, st.session_state.var_df)
                st.success(f"✅ {len(df)} OT et {len(avf)} avis chargés — {len(posts)} postes")
            except Exception as e: st.error(f"❌ Erreur : {e}"); st.exception(e)
    render_main()

if __name__ == "__main__":
    main()
