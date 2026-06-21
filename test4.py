# -*- coding: utf-8 -*-
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import io, locale, random, time, os, hashlib, json, base64
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ===========================================================
st.set_page_config(layout="wide", page_title="Dashboard KPI", initial_sidebar_state="expanded")
# ============================================================

QK = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois",
      "OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois",
      "OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois",
      "OT exécution 1mois< <3mois",
      "Performance Graissage","Performance Inspection","Performance Appels Systématiques"]
PK = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé",
      "Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL",
      "OT Fiabilité","Total Avis de Panne"]
ALL_KPI = QK + PK

CIBLE = {"TAUX_REALISATION_CORRECTIF/PT":95,"OT préparation <1 mois":80,"OT préparation >3 mois":5,
         "OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,
         "OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,
         "OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,
         "Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,
         "OT CONFIME":100,"OT_COR_EGAL":95,
         "Performance Graissage":95,"Performance Inspection":95,"Performance Appels Systématiques":85,
         "OT Fiabilité":100,"Total Avis de Panne":100}

LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois",
                "OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]

ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT.",
           "OT préparation <1 mois":"Reduire l'age de preparation des OT (< 1 mois).",
           "OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.",
           "OT planification <1 mois":"Reduire l'age de planification des OT (< 1 mois).",
           "OT planification >3 mois":"Traiter les OT avec planification > 3 mois.",
           "OT exécution <1 mois":"Reduire l'age d'execution des OT (< 1 mois).",
           "OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.",
           "OT LANC ESTIME":"Estimer les couts des OT lances.",
           "Backlog préparation caractérisé":"Caracteriser le backlog de preparation.",
           "Backlog planification caractérisé":"Caracteriser le backlog de planification.",
           "OT CONFIME":"Confirmer les OT termines.",
           "OT_COR_EGAL":"Rapprocher les couts reels et budgetes.",
           "appel avis approuvé":"Creer un OT pour les avis sans ordre.",
           "OT préparation 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT planification 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "Performance Graissage":"Ameliorer le taux de realisation des OT de graissage (Type 350).",
           "Performance Inspection":"Ameliorer le taux de realisation des OT d'inspection (Types 290,300,310).",
           "Performance Appels Systématiques":"Ameliorer le taux de realisation des appels systematiques (Type 360).",
           "OT Fiabilité":"Maintenir la fiabilite des OT a 100%.",
           "Total Avis de Panne":"Maintenir le suivi des avis de panne a 100%."}

KPI_RESP_MAP = {
    "TAUX_REALISATION_CORRECTIF/PT": "Chef d'atelier",
    "OT préparation <1 mois": "Préparateur BM",
    "OT préparation 1mois< <3mois": "Préparateur BM",
    "OT préparation >3 mois": "Préparateur BM",
    "OT planification <1 mois": "Planificateur BM",
    "OT planification 1mois< <3mois": "Planificateur BM",
    "OT planification >3 mois": "Planificateur BM",
    "OT exécution <1 mois": "Chef d'atelier",
    "OT exécution 1mois< <3mois": "Chef d'atelier",
    "OT exécution >3 mois": "Chef d'atelier",
    "appel avis approuvé": "Chef d'atelier",
    "OT LANC ESTIME": "Fiabilité",
    "Backlog préparation caractérisé": "Préparateur BM",
    "Backlog planification caractérisé": "Planificateur BM",
    "OT CONFIME": "Agent de saisie",
    "OT_COR_EGAL": "Agent de saisie",
    "Performance Graissage": "Chef d'atelier",
    "Performance Inspection": "Chef d'atelier",
    "Performance Appels Systématiques": "Chef d'atelier",
    "OT Fiabilité": "Fiabilité",
    "Total Avis de Panne": "Fiabilité"
}

MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]

CONSIGNES_HSE = [
    "Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de securite.",
    "Port obligatoire des lunettes de protection.","Port obligatoire des gants adaptes au travail.",
    "Utiliser les protections auditives dans les zones bruyantes.","Verifier l'absence de tension avant toute intervention electrique.",
    "Respecter la procedure de consignation et deconsignation.","Ne jamais intervenir sur un equipement en marche.",
    "Baliser et securiser la zone de travail.","Maintenir le poste de travail propre et ordonne.",
    "Verifier l'etat des outils avant utilisation.","Utiliser uniquement du materiel homologue.",
    "Respecter les permis de travail en vigueur.","Identifier les risques avant de commencer une tache.",
    "Signaler immediatement toute situation dangereuse.","Signaler tout incident ou presque accident.",
    "Ne jamais neutraliser un dispositif de securite.","Verifier les detecteurs de gaz avant utilisation.",
    "Verifier la bonne ventilation des zones de travail.","Respecter les regles des espaces confines.",
    "Controler l'atmosphere avant d'entrer dans un espace confine.","Utiliser les points d'ancrage pour les travaux en hauteur.",
    "Verifier l'etat des echafaudages avant utilisation.","Securiser les outils lors des travaux en hauteur.",
    "Ne pas travailler seul lors des operations a risque.","Controler les elingues avant chaque levage.",
    "Respecter les limites de charge des equipements.","Verifier l'etat des appareils de levage.",
    "Maintenir les voies de circulation degagees.","Respecter la signalisation de securite.",
    "Verifier les extincteurs a proximite du chantier.","Connaitre les issues de secours les plus proches.",
    "Respecter les procedures d'arret d'urgence.","Verifier les flexibles et raccords avant mise en service.",
    "Controler les fuites avant demarrage d'un equipement.","Respecter les distances de securite.",
    "Ne jamais contourner une procedure HSE.","Porter les EPI adaptes au risque identifie.",
    "Prevenir son responsable avant toute intervention particuliere.","Analyser les risques avant chaque demarrage de chantier.",
    "Verifier la stabilite des equipements.","Utiliser les bons outils pour la bonne tache.",
    "Respecter les consignes specifiques du chantier.","Ne jamais prendre de raccourci au detriment de la securite.",
    "Arreter immediatement les travaux en cas de danger.","Proteger l'environnement lors des interventions.",
    "Collecter et trier correctement les dechets.","Eviter toute pollution accidentelle.",
    "Respecter les consignes de stockage des produits dangereux.","Lire les fiches de securite avant manipulation.",
    "Verifier les equipements avant chaque prise de poste.","S'assurer de la disponibilite des moyens de secours.",
    "Communiquer clairement avec l'equipe avant intervention.","Respecter les regles de circulation des engins.",
    "Garder une vigilance permanente sur son environnement.","Prendre le temps d'effectuer le travail en securite.",
    "La securite est l'affaire de tous.","Chaque incident peut etre evite par la prevention.",
    "Aucun travail n'est plus urgent que la securite.","Zero accident commence par un comportement sur."]

def get_logo_base64():
    for path in ["logo.png", "./logo.png", "../logo.png"]:
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    return base64.b64encode(f.read()).decode()
            except Exception:
                pass
    return None

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return "18/06/2026"

def contient_mot(t,lm):
    t=str(t); return any(m in t for l in lm for m in l.split())

def cat_age(a):
    if pd.isna(a): return "Inconnu"
    if a<=1: return "<1 mois"
    elif a>=3: return ">3 mois"
    return "1 mois < <3 mois"

def excr(df):
    if "Poste travail princ." in df.columns:
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
    return df

@st.cache_data(show_spinner=False)
def prepare_data(ot_bytes, av_bytes, date_str):
    raw_ot = pd.read_excel(io.BytesIO(ot_bytes))
    raw_av = pd.read_excel(io.BytesIO(av_bytes))
    raw_ot = excr(raw_ot)
    raw_av = excr(raw_av)
    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in raw_ot.columns: raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in raw_av.columns: raw_av[c]=pd.to_datetime(raw_av[c],errors="coerce")
    now_ts = pd.to_datetime(date_str, format="%d/%m/%Y", errors='coerce')
    if pd.isna(now_ts): now_ts = pd.Timestamp.now()
    df = raw_ot.copy()
    df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
    df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
    df["Type Carac Prep"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MP_KW if kw in str(x)), "NON CARACTERISE"))
    df["Type Carac Plan"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MPLAN_KW if kw in str(x)), "NON CARACTERISE"))
    for dc,am,ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
        if dc in df.columns:
            df[am]=((now_ts.year-df[dc].dt.year)*12+(now_ts.month-df[dc].dt.month)).round(2)
            df[ac]=df[am].apply(cat_age)
        else: df[am]=np.nan; df[ac]="Inconnu"
    df["OT CONFIME"]=np.where((df["Statut système"].str.contains("CLOT",na=False)|df["Statut système"].str.contains("TCLO",na=False))&~df["Statut système"].str.contains("CONF",na=False),"NON","OUI")
    df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    df["_tw_num"]=pd.to_numeric(df.get("Type de travail",pd.Series(dtype=float)),errors="coerce")
    if "Statut système" in df.columns: df["Statut OT"]=df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
    avf = raw_av[(raw_av["Ordre"].isna())|(raw_av["Ordre"].astype(str).str.strip()=="")].copy()
    apm = sorted(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
    return df, avf, apm, now_ts

def save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,sheet_name):
    kpis_dir="kpis"; os.makedirs(kpis_dir,exist_ok=True)
    filepath=os.path.join(kpis_dir,"indicateurs_kpis.xlsx")
    sn=str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf=Font(bold=True,color="FFFFFF",size=10); hfl=PatternFill(start_color="1E3A5F",end_color="1E3A5F",fill_type="solid")
    tf=Font(bold=True,size=12,color="1E3A5F")
    tb=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    try: wb=load_workbook(filepath)
    except Exception: wb=Workbook()
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    if sn in wb.sheetnames: del wb[sn]
    ws=wb.create_sheet(sn); rn=1
    def ws_sec(title,cols,rows,sr):
        ws.cell(row=sr,column=1,value=title).font=tf; sr+=1
        for j,c in enumerate(cols,1):
            cl=ws.cell(row=sr,column=j,value=c); cl.font=hf; cl.fill=hfl; cl.alignment=Alignment(horizontal='center'); cl.border=tb
        sr+=1
        for r in rows:
            for j,c in enumerate(cols,1):
                cl=ws.cell(row=sr,column=j,value=r.get(c,"")); cl.border=tb; cl.alignment=Alignment(horizontal='center')
            sr+=1
        return sr+1
    rn=ws_sec("INDICATEURS DE PERFORMANCE",pcols,prows,rn)
    if ano_p_c and ano_p_r: rn=ws_sec("ANOMALIES PERFORMANCE",ano_p_c,ano_p_r,rn)
    rn=ws_sec("INDICATEURS DE QUALITE",qcols,qrows,rn)
    if ano_q_c and ano_q_r: rn=ws_sec("ANOMALIES QUALITE",ano_q_c,ano_q_r,rn)
    try: wb.save(filepath)
    except Exception: pass

def load_historical_kpis(filepath):
    if not os.path.exists(filepath): return pd.DataFrame()
    try: wb=load_workbook(filepath,read_only=True,data_only=True)
    except Exception: return pd.DataFrame()
    records=[]; section=None; headers=None
    for sheet_name in wb.sheetnames:
        try:
            ws=wb[sheet_name]; rows_data=list(ws.iter_rows(values_only=True))
            for row in rows_data:
                cell0=str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper(): section="perf"; headers=None; continue
                elif "INDICATEURS DE QUALITE" in cell0.upper(): section="qual"; headers=None; continue
                elif "ANOMALIES" in cell0.upper(): section=None; continue
                if section and headers is None and cell0:
                    headers=[str(c).strip() if c else "" for c in row]; continue
                if section and headers and cell0 and cell0 not in ("Cible","Total general",""):
                    entry={"Date":sheet_name}
                    for j,h in enumerate(headers):
                        if j<len(row): entry[h]=row[j]
                    entry["_section"]=section; records.append(entry)
        except Exception: continue
    wb.close()
    if not records: return pd.DataFrame()
    df=pd.DataFrame(records)
    df["Date_parsed"]=pd.to_datetime(df["Date"].str.replace("-","/"),format="%d/%m/%Y",errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)

def calculate_variations(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns: return pd.DataFrame()
    dates=sorted(hist_df["Date"].unique())
    if len(dates)<2: return pd.DataFrame()
    perf_df=hist_df[hist_df["_section"]=="perf"].copy()
    qual_df=hist_df[hist_df["_section"]=="qual"].copy()
    variations=[]
    for i in range(1,len(dates)):
        prev_date,curr_date=dates[i-1],dates[i]
        prev_perf=perf_df[perf_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        curr_perf=perf_df[perf_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        prev_qual=qual_df[qual_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        curr_qual=qual_df[qual_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        for sec_name,prev_d,curr_d,kpi_list in [("Performance",prev_perf,curr_perf,QK+["Score Performance"]),("Qualite",prev_qual,curr_qual,PK+["Score Qualite"])]:
            for poste in set(prev_d.index)&set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns: continue
                    try: pv=float(prev_d.loc[poste,kpi])
                    except Exception: continue
                    try: cv=float(curr_d.loc[poste,kpi])
                    except Exception: continue
                    diff=cv-pv; pct=(diff/pv*100) if pv!=0 else (100 if cv!=0 else 0)
                    if abs(diff)<=0.5: trend="stabilite"
                    elif diff>0.5: trend="hausse"
                    else: trend="baisse"
                    sens = "Stable"
                    if trend != "stabilite":
                        if (trend == "hausse" and kpi not in LOWER_BETTER) or (trend == "baisse" and kpi in LOWER_BETTER):
                            sens = "Amelioration"
                        else:
                            sens = "Degradation"
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,
                        "Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),
                        "Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend, "Sens":sens})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty: return pd.DataFrame()
    j=var_df.copy(); j["Significatif"]=j["Ecart %"].abs()>=5
    j=j[j["Significatif"]].copy()
    return j.sort_values(["Date actuelle","Ecart %"],ascending=[True,False])

def calculate_rankings(var_df):
    if var_df.empty: return pd.DataFrame(),pd.DataFrame()
    scores={}
    for poste in var_df["Poste"].unique():
        pv=var_df[var_df["Poste"]==poste].copy()
        scores[poste]=sum((-r["Ecart %"] if r["KPI"] in LOWER_BETTER else r["Ecart %"]) for _,r in pv.iterrows())
    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    return pd.DataFrame(ranked[:5],columns=["Poste","Score variation"]),pd.DataFrame(ranked[-5:][::-1],columns=["Poste","Score variation"])

# ============================================================
# SCORE POSTE : valeurs brutes, cibles brutes
# vert(1) si : valeur >= cible (normal) ou valeur <= cible (LOWER_BETTER)
# Score = (vert / total) * 100
# ============================================================
def calc_score_poste(vals_dict, kpi_list):
    vert = 0; total = 0
    for kpi in kpi_list:
        if kpi not in vals_dict: continue
        try: v = float(vals_dict[kpi])
        except: continue
        cible = CIBLE.get(kpi, 100.0)
        if kpi in LOWER_BETTER:
            if v <= cible: vert += 1
        else:
            if v >= cible: vert += 1
        total += 1
    return round((vert / total) * 100, 2) if total > 0 else 0.0

# ============================================================
# SCORE TOTAL GÉNÉRAL : pour les 6 LOWER_BETTER on fait 100-valeur
# puis tout est comparé avec "plus haut = mieux"
# vert(1) si : valeur_affichée >= cible_affichée
# Score = (vert / total) * 100
# ============================================================
def calc_score_total(raw_vals_dict, kpi_list):
    vert = 0; total = 0
    for kpi in kpi_list:
        if kpi not in raw_vals_dict: continue
        try: v = float(raw_vals_dict[kpi])
        except: continue
        cible = CIBLE.get(kpi, 100.0)
        if kpi in LOWER_BETTER:
            v_aff = 100.0 - v
            c_aff = 100.0 - cible
            if v_aff >= c_aff: vert += 1
        else:
            if v >= cible: vert += 1
        total += 1
    return round((vert / total) * 100, 2) if total > 0 else 0.0

# ============================================================
# VALEUR AFFICHÉE pour Total Général
# ============================================================
def get_total_display_val(kpi, raw_val):
    """Retourne la valeur affichée dans la ligne Total Général."""
    try: v = float(raw_val)
    except: return raw_val
    if kpi in LOWER_BETTER:
        return round(100.0 - v, 1)
    return round(v, 1)

def get_total_display_cible(kpi):
    """Retourne la cible affichée dans la ligne Total Général."""
    cible = CIBLE.get(kpi, 100.0)
    if kpi in LOWER_BETTER:
        return round(100.0 - cible, 1)
    return cible

# ============================================================
# COULEURS CELLULES - Postes (valeurs brutes)
# ============================================================
def get_cell_bg_poste(kpi, val):
    try: v = float(val)
    except: return ""
    cible = CIBLE.get(kpi, 100.0)
    if kpi in LOWER_BETTER:
        if v <= cible: return "background:#d1fae5;color:#065f46;font-weight:700;"
        elif v <= cible + 5: return "background:#fef3c7;color:#92400e;font-weight:600;"
        else: return "background:#fee2e2;color:#991b1b;font-weight:700;"
    else:
        if v >= cible: return "background:#d1fae5;color:#065f46;font-weight:700;"
        elif v >= cible - 5: return "background:#fef3c7;color:#92400e;font-weight:600;"
        else: return "background:#fee2e2;color:#991b1b;font-weight:700;"

# ============================================================
# COULEURS CELLULES - Total Général (valeurs inversées pour LOWER_BETTER)
# ============================================================
def get_cell_bg_total(kpi, raw_val):
    try: v = float(raw_val)
    except: return ""
    cible = CIBLE.get(kpi, 100.0)
    if kpi in LOWER_BETTER:
        v_aff = 100.0 - v
        c_aff = 100.0 - cible
        if v_aff >= c_aff: return "background:#d1fae5;color:#065f46;font-weight:700;"
        elif v_aff >= c_aff - 5: return "background:#fef3c7;color:#92400e;font-weight:600;"
        else: return "background:#fee2e2;color:#991b1b;font-weight:700;"
    else:
        if v >= cible: return "background:#d1fae5;color:#065f46;font-weight:700;"
        elif v >= cible - 5: return "background:#fef3c7;color:#92400e;font-weight:600;"
        else: return "background:#fee2e2;color:#991b1b;font-weight:700;"

# ============================================================
# COULEURS BARRES - Postes (valeurs brutes)
# ============================================================
def get_bar_color_poste(kpi, val):
    try: v = float(val)
    except: return "#cbd5e0"
    cible = CIBLE.get(kpi, 100.0)
    if kpi in LOWER_BETTER:
        if v <= cible: return "#38a169"
        elif v <= cible + 5: return "#f59e0b"
        else: return "#e53e3e"
    else:
        if v >= cible: return "#38a169"
        elif v >= cible - 5: return "#f59e0b"
        else: return "#e53e3e"

def get_score_bg(score):
    if score >= 80: return "background:#d1fae5;color:#065f46;font-weight:800;font-size:13px"
    elif score >= 60: return "background:#fef3c7;color:#92400e;font-weight:800;font-size:13px"
    else: return "background:#fee2e2;color:#991b1b;font-weight:800;font-size:13px"

def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    .main .block-container{max-width:100%!important;width:100%!important;padding-left:0.5rem!important;padding-right:0.5rem!important}
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--primary:#1e3a5f;--primary-light:#2c5282;--success:#10b981;--success-dark:#059669;--warning:#f59e0b;--warning-dark:#d97706;--danger:#ef4444;--danger-dark:#dc2626;--info:#3b82f6;--border:#e2e8f0;--radius:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#f8fafc;font-family:'Inter',sans-serif}
    .main .block-container{padding-top:.8rem;padding-bottom:.8rem}
    .mh{background:linear-gradient(135deg,#1e3a5f 0%,#2563eb 100%);padding:16px 24px;border-radius:12px;margin-bottom:12px;box-shadow:0 8px 24px rgba(30,58,95,0.15);display:flex;align-items:center;gap:16px}
    .mh h1{color:#fff;font-size:42px;font-weight:800;margin:0;flex:1}
    .mh .logo{height:50px;width:auto;max-width:150px;object-fit:contain;border-radius:6px}
    .mh .db{background:rgba(255,255,255,0.2);padding:6px 16px;border-radius:16px;color:#fff;font-size:20px;font-weight:600;border:1px solid rgba(255,255,255,0.3);backdrop-filter:blur(10px)}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px}
    .cc{background:#fff;border-radius:12px;padding:18px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);border-left:4px solid;transition:transform 0.2s,box-shadow 0.2s;text-align:center}
    .cc:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(0,0,0,0.1)}
    .cc .cv{font-size:32px;font-weight:900;line-height:1.1}
    .cc .cl{font-size:16px;color:#1e293b;font-weight:800;text-transform:uppercase;letter-spacing:.5px;margin-top:8px}
    .cc.c1{border-left-color:#3b82f6}.cc.c1 .cv{color:#2563eb}
    .cc.c2{border-left-color:#10b981}.cc.c2 .cv{color:#059669}
    .cc.c3{border-left-color:#8b5cf6}.cc.c3 .cv{color:#7c3aed}
    .cc.c4{border-left-color:#ef4444}.cc.c4 .cv{color:#dc2626}
    .stl{font-size:16px;font-weight:800;color:var(--primary);margin:10px 0 5px 0;padding-left:12px;border-left:4px solid var(--info)}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:13px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--primary);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:6px 8px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2563eb,#3b82f6)}
    .tw.pt thead th{background:linear-gradient(135deg,#059669,#10b981)}
    .tw.at thead th{background:linear-gradient(135deg,#dc2626,#ef4444)}
    .tw thead th:first-child{z-index:11;left:0}
    .tw tbody td:first-child{position:sticky;left:0;background:#fff;z-index:5;border-right:1px solid var(--border);color:#1e293b!important}
    .tw tbody tr:nth-child(even) td:first-child{background:#f8fafc}
    .tw tbody tr:hover td:first-child{background:#eff6ff}
    .tw tbody td{padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;color:#1e293b!important;text-align:center}
    .tw tbody tr:nth-child(even) td{background:#f8fafc}
    .tw tbody tr:hover td{background:#eff6ff!important}
    .cb td{background:#1e3a5f!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .ca{background:#fff;border-radius:var(--radius);padding:12px;margin-top:6px;border:1px solid var(--border);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:14px;font-weight:700;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid var(--border)}
    .car{display:flex;align-items:center;margin-bottom:6px;font-size:12px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:260px;font-weight:600;color:var(--primary);text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:26px;background:#edf2f7;border-radius:4px;overflow:visible;position:relative}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .target-mark{position:absolute;top:-4px;bottom:-4px;width:3px;background:var(--info)!important;z-index:20;transform:translateX(-50%);box-shadow:0 0 6px rgba(59,130,246,.9),0 0 2px rgba(0,0,0,.4);border-radius:2px}
    .car .cav-out{font-size:12px;font-weight:800;color:#1e293b;min-width:70px;text-align:right;padding-left:6px}
    .car .cav-tgt{font-size:10px;font-weight:700;color:#1e293b;min-width:50px;text-align:right;padding-left:4px;opacity:.7}
    .es{text-align:center;padding:14px;color:#64748b;font-size:14px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--primary),var(--primary-light));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,#1e40af 0%,#1e3a8a 50%,#1e3a5f 100%)!important}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stTextInput label{color:rgba(255,255,255,.9)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.1);border-radius:6px;padding:5px 10px;margin-bottom:5px;border:1px solid rgba(255,255,255,.15)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stTextInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    [data-testid="stHeaderActionElements"]{display:none!important}
    [data-testid="stActionButtonContainer"]{display:none!important}
    .footer{text-align:center;margin-top:30px;padding:15px;color:#64748b;font-size:13px;border-top:1px solid var(--border);font-weight:600}
    .stTabs [data-baseweb="tab-list"]{gap:6px;background:#e2e8f0;padding:6px;border-radius:8px;margin-bottom:8px}
    .stTabs [data-baseweb="tab"]{border-radius:6px;padding:12px 22px;font-weight:700;font-size:20px;line-height:1.5;min-height:48px}
    .stTabs [data-baseweb="tab"] span,.stTabs [data-baseweb="tab"]>div{font-size:22px!important}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--primary)!important;box-shadow:0 3px 8px rgba(0,0,0,.1);font-size:21px}
    .stTabs [data-baseweb="tab"] svg{width:22px;height:22px}
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg>div{background:#fff;border-radius:var(--radius);padding:8px 10px;border:1px solid var(--border)}
    .cg .ct{font-size:13px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--border)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f1f5f9}
    .cgr:last-child{border:none}
    .cgr .rk{width:18px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1e293b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:55px;text-align:right}
    @media(max-width:768px){
        .cr{grid-template-columns:repeat(2,1fr)}
        .mh{padding:8px 10px;gap:8px}.mh h1{font-size:18px}.mh .db{font-size:11px;padding:2px 8px}
        .cg{grid-template-columns:1fr}
        .car{flex-wrap:wrap;gap:2px}.car .cal{width:100%;text-align:left;padding-right:0;margin-bottom:2px}.car .cab{flex:1 1 70%}.car .cav-out,.car .cav-tgt{flex:1 1 15%;min-width:40px}
        .tw{font-size:10px}.tw thead th,.tw tbody td{padding:3px 4px}
        .stTabs [data-baseweb="tab"]{padding:8px 12px;font-size:15px}.stTabs [data-baseweb="tab"] span{font-size:16px!important}
    }
    @media print{section[data-testid="stSidebar"],header[data-testid="stHeader"],div[data-testid="stToolbar"],div[data-testid="stHeaderActionElements"],footer,.stDeployButton,#MainMenu{display:none!important}.main .block-container{padding-top:0!important;padding-left:0!important;padding-right:0!important;max-width:100%!important}.stButton,.stDownloadButton{display:none!important}*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}.tw{page-break-inside:avoid;overflow:visible!important}}
    </style>""",unsafe_allow_html=True)

# ============================================================
# FONCTION PRINCIPALE
# ============================================================
def main():
    try: locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception:
        try: locale.setlocale(locale.LC_ALL,'fr_FR')
        except Exception: pass
    inject_custom_css()
    fichier_date=get_date_from_file()

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche=False
    if not st.session_state.hse_affiche:
        c=random.choice(CONSIGNES_HSE)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:22px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>"""%c,unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche=True; st.rerun(); st.stop()

    # ---- Fonctions internes ----
    def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)
    def cpiv(df,f,c,p):
        return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)

    # ---- SIDEBAR ----
    with st.sidebar:
        st.markdown("### 📁 Fichiers")
        ot_file = st.file_uploader("Fichier OT", type=["xlsx","xls"], key="ot_up")
        av_file = st.file_uploader("Fichier Avis", type=["xlsx","xls"], key="av_up")
        st.markdown("### 📅 Date")
        date_input = st.text_input("Date de référence", value=fichier_date, key="date_in")
        st.markdown("### 🔧 Postes")
        all_posts = []
        if ot_file:
            try:
                _tmp = pd.read_excel(io.BytesIO(ot_file.getvalue()))
                if "Poste travail princ." in _tmp.columns:
                    _tmp = excr(_tmp)
                    all_posts = sorted(_tmp[_tmp["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
            except Exception: pass
        posts_sel = st.multiselect("Sélectionner postes", options=all_posts, default=all_posts, key="posts_sel")
        st.markdown("---")
        if st.button("📥 Exporter Excel", key="btn_export"):
            st.session_state["do_export"] = True
        else:
            st.session_state["do_export"] = False

    if not ot_file or not av_file:
        st.markdown('<div class="es" style="margin-top:60px"><div style="font-size:80px;margin-bottom:20px">📂</div>Veuillez charger les deux fichiers Excel<br>(OT et Avis de Panne)</div>', unsafe_allow_html=True)
        return

    # ---- PRÉPARATION DES DONNÉES ----
    df, avf, apm, now_ts = prepare_data(ot_file.getvalue(), av_file.getvalue(), date_input)
    posts = [p for p in posts_sel if p in df["Poste travail princ."].values]
    if not posts:
        st.warning("Aucun poste sélectionné trouvé dans les données.")
        return

    # ---- CALCUL DES KPIs (valeurs BRUTES, aucune inversion) ----
    res = calc_kpis(df, avf, now_ts, posts)
    ckdf = res['ckdf']  # contient les valeurs brutes

    # ================================================================
    # LIGNES PERFORMANCE - Postes (valeurs brutes, cibles brutes)
    # ================================================================
    pcols = ["Poste de travail"] + QK + ["Score Performance"]
    prows = []
    for poste in posts:
        row = {"Poste de travail": poste}
        for kpi in QK:
            row[kpi] = round(float(ckdf.loc[poste, kpi]), 1) if poste in ckdf.index and kpi in ckdf.columns else 0.0
        row["Score Performance"] = calc_score_poste(row, QK)
        prows.append(row)

    # ================================================================
    # LIGNE TOTAL GÉNÉRAL PERFORMANCE
    # - Moyennes brutes par KPI
    # - Pour les 6 LOWER_BETTER : affichage = 100 - moyenne brute
    # - Score = méthode vert(1)/rouge(0) avec valeurs inversées
    # ================================================================
    total_p_raw = {}  # moyennes brutes
    for kpi in QK:
        total_p_raw[kpi] = round(float(ckdf[kpi].mean()), 1) if kpi in ckdf.columns else 0.0

    total_p = {"Poste de travail": "Total general"}
    for kpi in QK:
        total_p[kpi] = get_total_display_val(kpi, total_p_raw[kpi])
    total_p["Score Performance"] = calc_score_total(total_p_raw, QK)
    prows.append(total_p)

    # ================================================================
    # LIGNES QUALITÉ - Postes (valeurs brutes, cibles brutes)
    # ================================================================
    qcols = ["Poste de travail"] + PK + ["Score Qualite"]
    qrows = []
    for poste in posts:
        row = {"Poste de travail": poste}
        for kpi in PK:
            row[kpi] = round(float(ckdf.loc[poste, kpi]), 1) if poste in ckdf.index and kpi in ckdf.columns else 0.0
        row["Score Qualite"] = calc_score_poste(row, PK)
        qrows.append(row)

    # ================================================================
    # LIGNE TOTAL GÉNÉRAL QUALITÉ
    # ================================================================
    total_q_raw = {}
    for kpi in PK:
        total_q_raw[kpi] = round(float(ckdf[kpi].mean()), 1) if kpi in ckdf.columns else 0.0

    total_q = {"Poste de travail": "Total general"}
    for kpi in PK:
        total_q[kpi] = get_total_display_val(kpi, total_q_raw[kpi])
    total_q["Score Qualite"] = calc_score_total(total_q_raw, PK)
    qrows.append(total_q)

    # ================================================================
    # ANOMALIES PERFORMANCE (valeurs brutes, pas d'inversion)
    # ================================================================
    ano_p_c = ["Poste de travail", "KPI", "Valeur", "Cible", "Ecart", "Responsable", "Action"]
    ano_p_r = []
    for row in prows:
        poste = row["Poste de travail"]
        if poste == "Total general": continue
        for kpi in QK:
            val = row[kpi]
            cible = CIBLE.get(kpi, 100.0)
            is_bad = False
            if kpi in LOWER_BETTER:
                if val > cible: is_bad = True
            else:
                if val < cible: is_bad = True
            if is_bad:
                ecart = round(val - cible, 1)
                ano_p_r.append({
                    "Poste de travail": poste, "KPI": kpi, "Valeur": val,
                    "Cible": cible, "Ecart": ecart,
                    "Responsable": KPI_RESP_MAP.get(kpi, ""), "Action": ACT_MAP.get(kpi, "")
                })

    # ================================================================
    # ANOMALIES QUALITÉ
    # ================================================================
    ano_q_c = ["Poste de travail", "KPI", "Valeur", "Cible", "Ecart", "Responsable", "Action"]
    ano_q_r = []
    for row in qrows:
        poste = row["Poste de travail"]
        if poste == "Total general": continue
        for kpi in PK:
            val = row[kpi]
            cible = CIBLE.get(kpi, 100.0)
            is_bad = False
            if kpi in LOWER_BETTER:
                if val > cible: is_bad = True
            else:
                if val < cible: is_bad = True
            if is_bad:
                ecart = round(val - cible, 1)
                ano_q_r.append({
                    "Poste de travail": poste, "KPI": kpi, "Valeur": val,
                    "Cible": cible, "Ecart": ecart,
                    "Responsable": KPI_RESP_MAP.get(kpi, ""), "Action": ACT_MAP.get(kpi, "")
                })

    # ---- EXPORT EXCEL ----
    if st.session_state.get("do_export"):
        save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, date_input)
        st.session_state["do_export"] = False
        st.success("✅ Export terminé : kpis/indicateurs_kpis.xlsx")

    # ================================================================
    # EN-TÊTE
    # ================================================================
    logo_b64 = get_logo_base64()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="logo">' if logo_b64 else ""
    st.markdown(f'''<div class="mh">{logo_html}<h1>Dashboard KPI Maintenance</h1><div class="db">📅 {date_input}</div></div>''', unsafe_allow_html=True)

    # ================================================================
    # CARTES RÉSUMÉ
    # ================================================================
    score_p_global = total_p["Score Performance"]
    score_q_global = total_q["Score Qualite"]
    total_ot = len(df)
    taux_corr = round(float(ckdf["TAUX_REALISATION_CORRECTIF/PT"].mean()), 1)

    st.markdown(f'''<div class="cr">
        <div class="cc c1"><div class="cv">{score_p_global:.0f}%</div><div class="cl">Score Performance</div></div>
        <div class="cc c2"><div class="cv">{score_q_global:.0f}%</div><div class="cl">Score Qualité</div></div>
        <div class="cc c3"><div class="cv">{total_ot}</div><div class="cl">Total OT</div></div>
        <div class="cc c4"><div class="cv">{taux_corr:.0f}%</div><div class="cl">Taux Réal. Correctif</div></div>
    </div>''', unsafe_allow_html=True)

    # ================================================================
    # NOTE EXPLICATIVE
    # ================================================================
    st.markdown('''<div style="background:#eff6ff;border:1px solid #3b82f6;border-radius:8px;padding:10px 16px;margin-bottom:10px;font-size:12px;color:#1e40af">
    <b>ℹ️ Méthode de calcul :</b> Score = (Nombre de cellules vertes / Nombre total de KPIs) × 100.<br>
    <b>Postes individuels :</b> valeurs brutes affichées. Pour les KPIs "&gt;3 mois" et "1mois&lt;&lt;3mois", vert si valeur ≤ cible (plus bas = mieux).<br>
    <b>Total Général uniquement :</b> les 6 KPIs "&gt;3 mois" et "1mois&lt;&lt;3mois" sont affichés avec <b>100 - valeur</b> et cible <b>100 - cible</b> pour un calcul uniforme.
    </div>''', unsafe_allow_html=True)

    # ================================================================
    # FONCTIONS DE RENDU HTML
    # ================================================================
    def render_kpi_table(rows, kpi_list, score_col, table_class):
        """Tableau HTML : postes = valeurs brutes, Total = valeurs inversées pour LOWER_BETTER."""
        cols = ["Poste de travail"] + kpi_list + [score_col]
        h = f'<table class="tw {table_class}"><thead><tr>'
        for c in cols:
            h += f'<th>{c}</th>'
        h += '</tr></thead><tbody>'

        for row in rows:
            is_total = (row["Poste de travail"] == "Total general")
            cls = ' class="cb"' if is_total else ''
            h += f'<tr{cls}>'
            h += f'<td style="text-align:left;font-weight:{"800" if is_total else "700"}">{row["Poste de travail"]}</td>'

            for kpi in kpi_list:
                if is_total:
                    # Total : afficher la valeur inversée pour LOWER_BETTER
                    val_display = row[kpi]  # déjà calculée par get_total_display_val
                    bg = get_cell_bg_total(kpi, total_p_raw.get(kpi, total_q_raw.get(kpi, 0)))
                    h += f'<td style="{bg}">{val_display:.1f}%</td>'
                else:
                    # Poste : afficher la valeur brute
                    val = row.get(kpi, 0)
                    bg = get_cell_bg_poste(kpi, val)
                    h += f'<td style="{bg}">{val:.1f}%</td>'

            sc = row.get(score_col, 0)
            if is_total:
                h += f'<td style="font-size:14px;font-weight:800">{sc:.1f}%</td>'
            else:
                h += f'<td style="{get_score_bg(sc)}">{sc:.1f}%</td>'
            h += '</tr>'

        # Ligne Cible
        h += '<tr style="background:#f0f9ff"><td style="text-align:left;font-weight:800;color:#1e40af">Cible</td>'
        for kpi in kpi_list:
            if is_total and kpi in LOWER_BETTER:
                cible_display = get_total_display_cible(kpi)
            else:
                cible_display = CIBLE.get(kpi, 100.0)
            h += f'<td style="color:#1e40af;font-weight:700">{cible_display:.0f}%</td>'
        h += '<td style="color:#1e40af;font-weight:700">100%</td></tr>'

        h += '</tbody></table>'
        return h

    def render_anomalies_table(ano_rows, ano_cols, table_class):
        if not ano_rows:
            return '<div class="es">✅ Aucune anomalie détectée</div>'
        h = f'<table class="tw {table_class}"><thead><tr>'
        for c in ano_cols:
            h += f'<th>{c}</th>'
        h += '</tr></thead><tbody>'
        for r in ano_rows:
            h += '<tr>'
            h += f'<td style="text-align:left;font-weight:700">{r["Poste de travail"]}</td>'
            h += f'<td style="text-align:left">{r["KPI"]}</td>'
            h += f'<td style="background:#fee2e2;color:#991b1b;font-weight:700">{r["Valeur"]:.1f}%</td>'
            h += f'<td>{r["Cible"]:.0f}%</td>'
            h += f'<td style="color:#dc2626;font-weight:700">{r["Ecart"]:.1f}</td>'
            h += f'<td>{r["Responsable"]}</td>'
            h += f'<td style="text-align:left;font-size:11px">{r["Action"]}</td>'
            h += '</tr>'
        h += '</tbody></table>'
        return h

    def render_bars_poste(poste, kpi_list, data_dict, is_total=False):
        """Barres de KPIs. Si is_total, utilise les valeurs inversées pour LOWER_BETTER."""
        html = f'<div class="ca"><div class="ct">📊 {poste}</div>'
        for kpi in kpi_list:
            cible = CIBLE.get(kpi, 100.0)
            if is_total:
                # Pour le Total, afficher la valeur inversée
                val_display = get_total_display_val(kpi, data_dict.get(kpi, 0))
                cible_display = get_total_display_cible(kpi)
                color = "#38a169" if val_display >= cible_display else "#f59e0b" if val_display >= cible_display - 5 else "#e53e3e"
                pct = min(max(val_display, 0), 120) / 120 * 100
                target_pct = min(cible_display, 120) / 120 * 100
                icon = "✅" if val_display >= cible_display else "❌"
                suffix = " *" if kpi in LOWER_BETTER else ""
            else:
                # Pour les postes, afficher la valeur brute
                val_display = data_dict.get(kpi, 0)
                cible_display = cible
                color = get_bar_color_poste(kpi, val_display)
                pct = min(max(val_display, 0), 120) / 120 * 100
                target_pct = min(cible_display, 120) / 120 * 100
                if kpi in LOWER_BETTER:
                    icon = "✅" if val_display <= cible_display else "❌"
                else:
                    icon = "✅" if val_display >= cible_display else "❌"
                suffix = ""
            html += f'''<div class="car">
                <div class="cal" title="{kpi}">{kpi}{suffix}</div>
                <div class="cab">
                    <div class="caf" style="width:{pct}%;background:{color}"></div>
                    <div class="target-mark" style="left:{target_pct}%"></div>
                </div>
                <div class="cav-out">{val_display:.1f}% {icon}</div>
                <div class="cav-tgt">C:{cible_display:.0f}%</div>
            </div>'''
        if is_total:
            html += '<div style="font-size:10px;color:#64748b;margin-top:6px">* Valeur affichée = 100 - valeur brute (inversion pour calcul uniforme)</div>'
        html += '</div>'
        return html

    # ================================================================
    # ONGLETS
    # ================================================================
    tab1, tab2, tab3 = st.tabs(["📈 Performance", "📐 Qualité", "📋 Plan d'Action"])

    # ================================================================
    # ONGLET 1 : PERFORMANCE
    # ================================================================
    with tab1:
        st.markdown('<div class="stl">Indicateurs de Performance</div>', unsafe_allow_html=True)

        sel_p = st.selectbox("Sélectionner un poste", options=posts + ["Total general"], index=len(posts), key="sel_perf")
        if sel_p == "Total general":
            st.markdown(render_bars_poste("Total general", QK, total_p_raw, is_total=True), unsafe_allow_html=True)
        else:
            poste_data = next((r for r in prows if r["Poste de travail"] == sel_p), None)
            if poste_data:
                st.markdown(render_bars_poste(sel_p, QK, poste_data, is_total=False), unsafe_allow_html=True)

        st.markdown(render_kpi_table(prows, QK, "Score Performance", "pt"), unsafe_allow_html=True)

        if ano_p_r:
            st.markdown('<div class="stl" style="color:#dc2626;border-left-color:#dc2626">⚠️ Anomalies Performance</div>', unsafe_allow_html=True)
            st.markdown(render_anomalies_table(ano_p_r, ano_p_c, "at"), unsafe_allow_html=True)

    # ================================================================
    # ONGLET 2 : QUALITÉ
    # ================================================================
    with tab2:
        st.markdown('<div class="stl">Indicateurs de Qualité</div>', unsafe_allow_html=True)

        sel_q = st.selectbox("Sélectionner un poste", options=posts + ["Total general"], index=len(posts), key="sel_qual")
        if sel_q == "Total general":
            st.markdown(render_bars_poste("Total general", PK, total_q_raw, is_total=True), unsafe_allow_html=True)
        else:
            poste_data = next((r for r in qrows if r["Poste de travail"] == sel_q), None)
            if poste_data:
                st.markdown(render_bars_poste(sel_q, PK, poste_data, is_total=False), unsafe_allow_html=True)

        st.markdown(render_kpi_table(qrows, PK, "Score Qualite", "qt"), unsafe_allow_html=True)

        if ano_q_r:
            st.markdown('<div class="stl" style="color:#dc2626;border-left-color:#dc2626">⚠️ Anomalies Qualité</div>', unsafe_allow_html=True)
            st.markdown(render_anomalies_table(ano_q_r, ano_q_c, "at"), unsafe_allow_html=True)

    # ================================================================
    # ONGLET 3 : PLAN D'ACTION
    # ================================================================
    with tab3:
        st.markdown('<div class="stl">Plan d\'Action Correctif</div>', unsafe_allow_html=True)

        all_anomalies = []
        for r in ano_p_r:
            all_anomalies.append({**r, "Type": "Performance"})
        for r in ano_q_r:
            all_anomalies.append({**r, "Type": "Qualité"})

        if not all_anomalies:
            st.markdown('<div class="es" style="font-size:18px;padding:40px">✅ Aucune action corrective requise — Tous les KPIs sont à leur cible !</div>', unsafe_allow_html=True)
        else:
            # Trier par poste puis par KPI
            all_anomalies.sort(key=lambda x: (x["Poste de travail"], x["Type"], x["KPI"]))

            h = '<table class="plan-action-table"><thead><tr>'
            h += '<th>Poste</th><th>Type</th><th>KPI</th><th>Valeur</th><th>Cible</th><th>Écart</th><th>Responsable</th><th>Action Corrective</th><th>Statut</th>'
            h += '</tr></thead><tbody>'

            for i, r in enumerate(all_anomalies):
                bg = "#fef2f2" if i % 2 == 0 else "#fff"
                h += f'<tr style="background:{bg}">'
                h += f'<td style="text-align:left;font-weight:800">{r["Poste de travail"]}</td>'
                h += f'<td>{r["Type"]}</td>'
                h += f'<td style="text-align:left">{r["KPI"]}</td>'
                h += f'<td style="color:#dc2626;font-weight:700">{r["Valeur"]:.1f}%</td>'
                h += f'<td>{r["Cible"]:.0f}%</td>'
                h += f'<td style="color:#dc2626;font-weight:700">{r["Ecart"]:.1f}</td>'
                h += f'<td>{r["Responsable"]}</td>'
                h += f'<td style="text-align:left;font-size:11px">{r["Action"]}</td>'
                h += f'<td style="color:#d97706;font-weight:700">⏳ En cours</td>'
                h += '</tr>'

            h += f'<tr style="background:#1e3a5f"><td colspan="9" style="color:#fff;font-weight:800;text-align:center">Total : {len(all_anomalies)} action(s) corrective(s)</td></tr>'
            h += '</tbody></table>'
            st.markdown(h, unsafe_allow_html=True)

            # Résumé par responsable
            st.markdown('<div class="stl">Résumé par Responsable</div>', unsafe_allow_html=True)
            resp_counts = {}
            for r in all_anomalies:
                resp = r["Responsable"]
                if resp:
                    resp_counts[resp] = resp_counts.get(resp, 0) + 1

            if resp_counts:
                h2 = '<div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:8px">'
                for resp, cnt in sorted(resp_counts.items(), key=lambda x: -x[1]):
                    color = "#dc2626" if cnt >= 5 else "#d97706" if cnt >= 3 else "#059669"
                    h2 += f'<div style="background:#fff;border:2px solid {color};border-radius:8px;padding:12px 20px;text-align:center;min-width:150px">'
                    h2 += f'<div style="font-size:28px;font-weight:900;color:{color}">{cnt}</div>'
                    h2 += f'<div style="font-size:12px;font-weight:700;color:#1e293b">{resp}</div>'
                    h2 += '</div>'
                h2 += '</div>'
                st.markdown(h2, unsafe_allow_html=True)

    # ---- FOOTER ----
    st.markdown('<div class="footer">Dashboard KPI Maintenance — Généré le %s</div>' % datetime.now().strftime("%d/%m/%Y %H:%M"), unsafe_allow_html=True)


# ============================================================
# FONCTION calc_kpis (était tronquée dans l'original)
# ============================================================
def calc_kpis(df_i, av_i, now_ts, posts):
    res={}; df=df_i.copy(); av=av_i.copy()
    res['dfp']=df

    filt_corr=(df["Nº appel pl.entret."].fillna(0)==0)&(df["Contient SOPL"]==1)
    an=cpiv(df,filt_corr,"Statut OT",posts)
    for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
    an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]
    an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
    an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"]))

    pr=cpiv(df,(df["Statut OT"]=="CRÉÉ")&(df["Statut utilisateur"].str.contains("CRPR",na=False)),"ap",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pr[c]=pr.get(c,0)
    pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"])
    pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0)
    pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)

    pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)),"alp",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
    pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"])
    pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0)
    pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)

    ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
    ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
    ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"])
    ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0)
    ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)

    la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: la[c]=la.get(c,0)
    la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])

    pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
    pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])

    plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
    plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])

    for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
        pv=pd.pivot_table(df,index="Poste travail princ.",columns=cn,values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: pv[c]=pv.get(c,0)
        pv["Total"]=pv["OUI"]+pv["NON"]; pv[cn]=ckpi(pv["OUI"],pv["Total"]); res[kn.lower().replace(" ","_")]=pv

    avf=av.copy(); res['avf']=avf
    tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
    tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])

    g_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
    g_den=df[(df["Contient SOPL"]==1)&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
    g_df=pd.DataFrame({"_n":g_num,"_d":g_den}).reindex(posts,fill_value=0)
    g_df["Performance Graissage"]=np.where(g_df["_d"]==0,100.0,(g_df["_n"]/g_df["_d"])*100)

    ins_types=[290,300,310]
    ins_base=(df["_tw_num"].isin(ins_types))&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
    ins_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&ins_base].groupby("Poste travail princ.")["Ordre"].count()
    ins_den=df[(df["Contient SOPL"]==1)&ins_base].groupby("Poste travail princ.")["Ordre"].count()
    ins_df=pd.DataFrame({"_n":ins_num,"_d":ins_den}).reindex(posts,fill_value=0)
    ins_df["Performance Inspection"]=np.where(ins_df["_d"]==0,100.0,(ins_df["_n"]/ins_df["_d"])*100)

    sys_base=(df["_tw_num"]==360)&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
    sys_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&sys_base].groupby("Poste travail princ.")["Ordre"].count()
    sys_den=df[(df["Contient SOPL"]==1)&sys_base].groupby("Poste travail princ.")["Ordre"].count()
    sys_df=pd.DataFrame({"_n":sys_num,"_d":sys_den}).reindex(posts,fill_value=0)
    sys_df["Performance Appels Systématiques"]=np.where(sys_df["_d"]==0,100.0,(sys_df["_n"]/sys_df["_d"])*100)

    fiab_s=pd.Series(100.0,index=posts); avpan_s=pd.Series(100.0,index=posts)

    res['ckdf']=pd.DataFrame({
        "TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],
        "OT préparation <1 mois":pr["OT préparation <1 mois"],
        "OT préparation >3 mois":pr["OT préparation >3 mois"],
        "OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
        "OT planification <1 mois":pl["OT planification <1 mois"],
        "OT planification >3 mois":pl["OT planification >3 mois"],
        "OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
        "OT exécution <1 mois":ex["OT exécution <1 mois"],
        "OT exécution >3 mois":ex["OT exécution >3 mois"],
        "OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
        "Performance Graissage":g_df["Performance Graissage"],
        "Performance Inspection":ins_df["Performance Inspection"],
        "Performance Appels Systématiques":sys_df["Performance Appels Systématiques"],
        "appel avis approuvé":tca["appel avis approuvé"],
        "OT LANC ESTIME":la["OT LANC ESTIME"],
        "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],
        "Backlog planification caractérisé":plc["Backlog planification caractérisé"],
        "OT CONFIME":res['ot_confime']["OT CONFIME"],
        "OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"],
        "OT Fiabilité":fiab_s,
        "Total Avis de Panne":avpan_s
    })
    return res

# ============================================================
if __name__ == "__main__":
    main()
