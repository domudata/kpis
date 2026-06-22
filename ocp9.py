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

# ============================================================
st.set_page_config(layout="wide", page_title="Dashboard KPI", initial_sidebar_state="expanded")
# ============================================================

QK = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois",
      "OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois",
      "OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois",
      "OT exécution 1mois< <3mois",
      "Performance Graissage","Performance Inspection","Performance Appels Systématiques"]
PK = ["Taux d'approbation des Avis","OT LANC ESTIME","Backlog préparation caractérisé",
      "Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL",
      "OT Fiabilité","Total Avis de Panne"]
ALL_KPI = QK + PK

CIBLE = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,
         "OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,
         "OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,
         "OT exécution 1mois< <3mois":15,"Taux d'approbation des Avis":95,"OT LANC ESTIME":100,
         "Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,
         "OT CONFIME":100,"OT_COR_EGAL":100,
         "Performance Graissage":95,"Performance Inspection":95,"Performance Appels Systématiques":95,
         "OT Fiabilité":100,"Total Avis de Panne":100}

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
           "Taux d'approbation des Avis":"Creer un OT pour les avis sans ordre.",
           "OT préparation 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT planification 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "Performance Graissage":"Ameliorer le taux de realisation des OT de graissage (Type 350).",
           "Performance Inspection":"Ameliorer le taux de realisation des OT d'inspection (Types 290,300,310).",
           "Performance Appels Systématiques":"Ameliorer le taux de realisation des appels systematiques (Type 360).",
           "OT Fiabilité":"Maintenir la fiabilite des OT a 100%.",
           "Total Avis de Panne":"Maintenir le suivi des avis de panne a 100%."}

KPI_RESP_MAP = {
    "TAUX_REALISATION_CORRECTIF/PT": "Chef d'atelier", "OT préparation <1 mois": "Préparateur BM",
    "OT préparation 1mois< <3mois": "Préparateur BM", "OT préparation >3 mois": "Préparateur BM",
    "OT planification <1 mois": "Planificateur BM", "OT planification 1mois< <3mois": "Planificateur BM",
    "OT planification >3 mois": "Planificateur BM", "OT exécution <1 mois": "Chef d'atelier",
    "OT exécution 1mois< <3mois": "Chef d'atelier", "OT exécution >3 mois": "Chef d'atelier",
    "Taux d'approbation des Avis": "Chef d'atelier", "OT LANC ESTIME": "Fiabilité",
    "Backlog préparation caractérisé": "Préparateur BM", "Backlog planification caractérisé": "Planificateur BM",
    "OT CONFIME": "Agent de saisie", "OT_COR_EGAL": "Agent de saisie",
    "Performance Graissage": "Chef d'atelier", "Performance Inspection": "Chef d'atelier",
    "Performance Appels Systématiques": "Chef d'atelier", "OT Fiabilité": "Fiabilité",
    "Total Avis de Panne": "Fiabilité"
}

LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois",
                "OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]

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
                with open(path, "rb") as f: return base64.b64encode(f.read()).decode()
            except Exception: pass
    return None

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return pd.Timestamp.today().strftime("%d/%m/%Y")

def contient_mot(t,lm):
    t=str(t); return any(m in t for l in lm for m in l.split())
def cat_age(a):
    if pd.isna(a): return "Inconnu"
    a = float(a)
    if a<=1: return "<1 mois"
    elif a>=3: return ">3 mois"
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
            except Exception: bio.seek(0); continue
    if header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        for engine in ['xlrd', 'calamine']:
            try: return pd.read_excel(bio, engine=engine)
            except Exception: bio.seek(0); continue
    for engine in ['openpyxl', 'xlrd', 'calamine']:
        try: bio.seek(0); return pd.read_excel(bio, engine=engine)
        except Exception: continue
    raise ValueError("Format de fichier non reconnu.")

@st.cache_data(show_spinner=False)
def prepare_data(ot_bytes, av_bytes, date_str):
    raw_ot = read_excel_safe(ot_bytes); raw_av = read_excel_safe(av_bytes)
    raw_ot = excr(raw_ot); raw_av = excr(raw_av)
    
    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in raw_ot.columns: raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in raw_av.columns: raw_av[c]=pd.to_datetime(raw_av[c],errors="coerce")
        
    now_ts = pd.Timestamp.today()
    df = raw_ot.copy()
    
    # Sécurisation anti-NaN ambigus
    df["Statut utilisateur"] = df["Statut utilisateur"].fillna("").astype(str)
    df["Statut système"] = df["Statut système"].fillna("").astype(str)
    
    df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
    df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
    df["Type Carac Prep"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MP_KW if kw in str(x)), "NON CARACTERISE"))
    df["Type Carac Plan"]=df["Statut utilisateur"].apply(lambda x: next((kw.split()[0] for kw in MPLAN_KW if kw in str(x)), "NON CARACTERISE"))
    
    for dc,am,ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
        if dc in df.columns:
            df[am]=((now_ts.year-df[dc].dt.year)*12+(now_ts.month-df[dc].dt.month)).round(2)
            df[ac]=df[am].apply(cat_age)
        else: df[am]=np.nan; df[ac]="Inconnu"
        
    df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLOT|TCLO",na=False) & df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
    df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    df["_tw_num"]=pd.to_numeric(df.get("Type de travail",pd.Series(dtype=float)),errors="coerce")
    
    if "Statut système" in df.columns: df["Statut OT"]=df["Statut système"].str.strip().str.split().str[0]
    else: df["Statut OT"] = ""
    
    avf = raw_av[(raw_av["Ordre"].isna() | (raw_av["Ordre"].astype(str).str.strip() == "")) & (raw_av["Type d'avis"].isin(["ZU","Z4","ZR","ZP"]))].copy()
    apm = sorted(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
    
    return df, avf, apm, now_ts

def save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,sheet_name):
    kpis_dir="kpis"; os.makedirs(kpis_dir,exist_ok=True)
    filepath=os.path.join(kpis_dir,"indicateurs_kpis.xlsx")
    sn=str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf=Font(bold=True,color="FFFFFF",size=10); hfl=PatternFill(start_color="1E3A5F",end_color="1E3A5F",fill_type="solid")
    tf=Font(bold=True,size=12,color="1E3A5F"); tb=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
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
                if section and headers is None and cell0: headers=[str(c).strip() if c else "" for c in row]; continue
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
    perf_df=hist_df[hist_df["_section"]=="perf"].copy(); qual_df=hist_df[hist_df["_section"]=="qual"].copy()
    variations=[]
    for i in range(1,len(dates)):
        prev_date,curr_date=dates[i-1],dates[i]
        prev_perf=perf_df[perf_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        curr_perf=perf_df[perf_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        prev_qual=qual_df[qual_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        curr_qual=qual_df[qual_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        for sec_name,prev_d,curr_d,kpi_list in [("Performance",prev_perf,curr_perf,PK+["Score Performance"]),("Qualite",prev_qual,curr_qual,QK+["Score Qualite"])]:
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
                        if (trend == "hausse" and kpi not in LOWER_BETTER) or (trend == "baisse" and kpi in LOWER_BETTER): sens = "Amelioration"
                        else: sens = "Degradation"
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,"Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),"Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend, "Sens":sens})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty: return pd.DataFrame()
    j=var_df.copy(); j["Significatif"]=j["Ecart %"].abs()>=5
    return j[j["Significatif"]].copy().sort_values(["Date actuelle","Ecart %"],ascending=[True,False])

def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    .main .block-container{max-width:100%!important;width:100%!important;padding-left:0.5rem!important;padding-right:0.5rem!important}
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--primary:#1e3a5f;--primary-light:#2c5282;--success:#10b981;--warning:#f59e0b;--danger:#ef4444;--info:#3b82f6;--border:#e2e8f0;--radius:10px;}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#f8fafc;font-family:'Inter',sans-serif}
    .main .block-container{padding-top:.8rem;padding-bottom:.8rem}
    .mh{background:linear-gradient(135deg,#1e3a5f 0%,#2563eb 100%);padding:16px 24px;border-radius:12px;margin-bottom:12px;box-shadow:0 8px 24px rgba(30,58,95,0.15);display:flex;align-items:center;gap:16px;}
    .mh h1{color:#fff;font-size:42px;font-weight:800;margin:0;flex:1}
    .mh .db{background:rgba(255,255,255,0.2);padding:6px 16px;border-radius:16px;color:#fff;font-size:20px;font-weight:600;border:1px solid rgba(255,255,255,0.3);}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px}
    .cc{background:#fff;border-radius:12px;padding:18px 16px;box-shadow:0 4px 12px rgba(0,0,0,0.06);border-left:4px solid;transition:transform 0.2s,box-shadow 0.2s;text-align:center;}
    .cc:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(0,0,0,0.1);}
    .cc .cv{font-size:32px;font-weight:900;line-height:1.1;}
    .cc .cl{font-size:16px;color:#1e293b;font-weight:800;text-transform:uppercase;letter-spacing:.5px;margin-top:8px;}
    .cc.c1{border-left-color:#3b82f6}.cc.c1 .cv{color:#2563eb}
    .cc.c2{border-left-color:#10b981}.cc.c2 .cv{color:#059669}
    .cc.c4{border-left-color:#ef4444}.cc.c4 .cv{color:#dc2626}
    .cc.c5{border-left-color:#3b82f6}.cc.c5 .cv{color:#2563eb}
    .stl{font-size:16px;font-weight:800;color:var(--primary);margin:10px 0 5px 0;padding-left:12px;border-left:4px solid var(--info);}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:13px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--primary);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:6px 8px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2563eb,#3b82f6)}
    .tw.pt thead th{background:linear-gradient(135deg,#059669,#10b981)}
    .tw.anom thead th{background:linear-gradient(135deg,#d97706,#f59e0b)}
    .tw thead th:first-child{z-index:11;left:0}
    .tw tbody td:first-child{position:sticky;left:0;background:#fff;z-index:5;border-right:1px solid var(--border);color:#1e293b !important}
    .tw tbody tr:nth-child(even) td:first-child{background:#f8fafc}
    .tw tbody tr:hover td:first-child{background:#eff6ff}
    .tw tbody td{padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;color:#1e293b !important}
    .tw tbody tr:nth-child(even) td{background:#f8fafc}
    .tw tbody tr:hover td{background:#eff6ff!important}
    .cb td{background:#2563eb!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .plan-action-table { width:100%; border-collapse:collapse; font-family:Inter,sans-serif; font-size:12px; border:1px solid #cbd5e1; }
    .plan-action-table th { background:#1e3a5f; color:#fff; font-weight:700; padding:8px 6px; border:1px solid #1e3a5f; }
    .plan-action-table td { padding:6px 8px; border:1px solid #cbd5e1; text-align:center; vertical-align:middle;}
    .plan-action-table td:first-child { text-align:left; font-weight:800; }
    .stTabs [data-baseweb="tab-list"]{gap:6px;background:#e2e8f0;padding:6px;border-radius:8px;margin-bottom:8px}
    .stTabs [data-baseweb="tab"]{border-radius:6px;padding:12px 22px;font-weight:700;font-size:20px;line-height:1.5;min-height:48px;}
    .stTabs [data-baseweb="tab"] span,.stTabs [data-baseweb="tab"] > div{font-size:22px !important;}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--primary)!important;box-shadow:0 3px 8px rgba(0,0,0,.1);font-size:21px;}
    .stTabs [data-baseweb="tab"] svg{width:22px;height:22px}
    .es{text-align:center;padding:14px;color:#64748b;font-size:14px}
    [data-testid="stHeaderActionElements"]{display:none !important;}
    [data-testid="stActionButtonContainer"]{display:none !important;}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh{padding:8px 10px;gap:8px}.mh h1{font-size:18px}.tw{font-size:10px}}
    @media print {section[data-testid="stSidebar"], header[data-testid="stHeader"], div[data-testid="stToolbar"], footer, .stDeployButton, #MainMenu { display: none !important; } .main .block-container { padding-top: 0 !important; max-width: 100% !important; } .stButton, .stDownloadButton { display: none !important; } * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
    </style>""",unsafe_allow_html=True)

# ============================================================
def main():
    try: locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception:
        try: locale.setlocale(locale.LC_ALL,'fr_FR')
        except Exception: pass
    inject_custom_css()
    fichier_date = get_date_from_file()

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

    def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)
    def cpiv(df,f,c,p): return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)
        
    def show_pie_pair(piv_df, title_prefix):
        global_counts = piv_df[["CRÉÉ","LANC","CLOT","TCLO"]].sum()
        global_counts = global_counts[global_counts > 0]
        realised = global_counts.get("CLOT", 0) + global_counts.get("TCLO", 0)
        not_realised = global_counts.sum() - realised
        if global_counts.empty: st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True); return
        colors = ["#8b5cf6", "#f59e0b", "#10b981", "#3b82f6"]
        fig = make_subplots(rows=1, cols=2, specs=[[{"type":"domain"},{"type":"domain"}]], subplot_titles=(f"{title_prefix} — Par Statut OT", f"{title_prefix} — Réalisés vs Non Réalisés"))
        fig.add_trace(go.Pie(labels=global_counts.index, values=global_counts.values, hole=0.4, textinfo='percent+label', texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', textposition='inside', insidetextorientation='radial', textfont=dict(size=14, color='white'), marker=dict(colors=colors, line=dict(color='#FFFFFF', width=3))), 1, 1)
        pie2_data = pd.Series([realised, not_realised], index=["Réalisés (CLOT+TCLO)", "Non Réalisés"])
        fig.add_trace(go.Pie(labels=pie2_data.index, values=pie2_data.values, hole=0.5, textinfo='percent+label', texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', textposition='inside', insidetextorientation='radial', textfont=dict(size=14, color='white'), marker=dict(colors=["#10b981", "#8b5cf6"], line=dict(color='#FFFFFF', width=3))), 1, 2)
        fig.update_layout(margin=dict(t=80, b=20, l=20, r=20), height=450, legend=dict(orientation="h", yanchor="bottom", y=-0.12, x=0.5, xanchor="center"))
        st.plotly_chart(fig, use_container_width=True)

    def show_simple_pie(piv_df, title, keep_non_carac=False):
        if not keep_non_carac and "NON CARACTERISE" in piv_df.columns: piv_df = piv_df.drop(columns=["NON CARACTERISE"])
        counts = piv_df.sum(); counts = counts[counts > 0]
        if counts.empty: st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True); return
        type_palette = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#14b8a6']
        colors = [type_palette[i % len(type_palette)] for i in range(len(counts))]
        total_sum = counts.sum()
        pull_list = [0.05 if (v/total_sum)*100 < 10 else 0 for v in counts.values]
        fig = go.Figure(go.Pie(labels=counts.index, values=counts.values, hole=0.4, sort=False, textinfo="percent", textposition="outside", pull=pull_list, marker=dict(colors=colors, line=dict(color="white", width=2))))
        fig.update_traces(hovertemplate="<b>%{label}</b><br>Nombre : %{value}<extra></extra>", textfont=dict(size=13))
        fig.update_layout(title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)), height=500, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"), margin=dict(t=80, b=80, l=40, r=40))
        st.plotly_chart(fig, use_container_width=True)

    def calc_kpis(df_i, av_i, now_ts, posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        res['dfp']=df
        anom = {}

        is_cree = df["Statut OT"] == "CRÉÉ"
        is_lanc = df["Statut OT"] == "LANC"
        is_clos = df["Statut OT"].isin(["CLOT","TCLO"])
        has_sopl = df["Contient SOPL"] == 1
        has_crpr = df["Statut utilisateur"].str.contains(r"\bCRPR\b",case=False,na=False)
        has_atpl = df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)

        filt_corr = (df["Nº appel pl.entret."].fillna(0)==0) & has_sopl
        an=cpiv(df,filt_corr,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
        an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]; an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"]))
        anom["TAUX_REALISATION_CORRECTIF/PT"] = an["CRÉÉ"] + an["LANC"]
        
        pr=cpiv(df, is_cree & has_crpr, "ap", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pr[c]=pr.get(c,0)
        pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"]); pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0); pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)
        anom["OT préparation <1 mois"] = pr["Total"] - pr["<1 mois"]; anom["OT préparation >3 mois"] = pr[">3 mois"]; anom["OT préparation 1mois< <3mois"] = pr["1 mois < <3 mois"]
        
        pl=cpiv(df, is_lanc & has_atpl, "alp", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"]); pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0); pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)
        anom["OT planification <1 mois"] = pl["Total"] - pl["<1 mois"]; anom["OT planification >3 mois"] = pl[">3 mois"]; anom["OT planification 1mois< <3mois"] = pl["1 mois < <3 mois"]
        
        ex=cpiv(df, is_lanc & has_sopl, "aex", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"]); ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0); ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)
        anom["OT exécution <1 mois"] = ex["Total"] - ex["<1 mois"]; anom["OT exécution >3 mois"] = ex[">3 mois"]; anom["OT exécution 1mois< <3mois"] = ex["1 mois < <3 mois"]
        
        la=pd.pivot_table(df[is_lanc],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: la[c]=la.get(c,0)
        la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"]); anom["OT LANC ESTIME"] = la["NON"]
        
        pc=pd.pivot_table(df[is_cree],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
        pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"]); anom["Backlog préparation caractérisé"] = pc["NON CARACTERISE"]
        
        plc=pd.pivot_table(df[is_lanc & ~has_sopl],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"]); anom["Backlog planification caractérisé"] = plc["NON CARACTERISE"]
        
        pv_conf=pd.pivot_table(df[is_clos],index="Poste travail princ.",columns="OT CONFIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: pv_conf[c]=pv_conf.get(c,0)
        pv_conf["Total"]=pv_conf["OUI"]+pv_conf["NON"]; pv_conf["OT CONFIME"]=ckpi(pv_conf["OUI"],pv_conf["Total"]); res['ot_confime']=pv_conf; anom["OT CONFIME"] = pv_conf["NON"]

        pv_cor=pd.pivot_table(df[is_clos],index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: pv_cor[c]=pv_cor.get(c,0)
        pv_cor["Total"]=pv_cor["OUI"]+pv_cor["NON"]; pv_cor["OT_COR_EGAL"]=ckpi(pv_cor["OUI"],pv_cor["Total"]); res['ot_cor_egal']=pv_cor; anom["OT_COR_EGAL"] = pv_cor["NON"]
        
        avf=av.copy(); res['avf']=avf
        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["Taux d'approbation des Avis"]=ckpi(tca["APRV"],tca["Total"]); anom["Taux d'approbation des Avis"] = tca["Total"] - tca["APRV"]
        
        g_num=df[is_clos & (df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_den=df[has_sopl & (df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_df=pd.DataFrame({"_n":g_num,"_d":g_den}).reindex(posts,fill_value=0)
        g_df["Performance Graissage"]=np.where(g_df["_d"]==0,100.0,(g_df["_n"]/g_df["_d"])*100); anom["Performance Graissage"] = (g_df["_d"] - g_df["_n"])

        ins_types=[290,300,310]
        ins_base=(df["_tw_num"].isin(ins_types))&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
        ins_num=df[is_clos & ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_den=df[has_sopl & ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_df=pd.DataFrame({"_n":ins_num,"_d":ins_den}).reindex(posts,fill_value=0)
        ins_df["Performance Inspection"]=np.where(ins_df["_d"]==0,100.0,(ins_df["_n"]/ins_df["_d"])*100); anom["Performance Inspection"] = (ins_df["_d"] - ins_df["_n"])

        sys_base=(df["_tw_num"]==360)&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
        sys_num=df[is_clos & sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_den=df[has_sopl & sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_df=pd.DataFrame({"_n":sys_num,"_d":sys_den}).reindex(posts,fill_value=0)
        sys_df["Performance Appels Systématiques"]=np.where(sys_df["_d"]==0,100.0,(sys_df["_n"]/sys_df["_d"])*100); anom["Performance Appels Systématiques"] = (sys_df["_d"] - sys_df["_n"])

        fiab_s=pd.Series(100.0,index=posts); avpan_s=pd.Series(100.0,index=posts)
        anom["OT Fiabilité"] = pd.Series(0, index=posts); anom["Total Avis de Panne"] = pd.Series(0, index=posts)

        res['anom'] = pd.DataFrame(anom, index=posts).fillna(0).astype(int)

        res['ckdf']=pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "Performance Graissage":g_df["Performance Graissage"],"Performance Inspection":ins_df["Performance Inspection"],"Performance Appels Systématiques":sys_df["Performance Appels Systématiques"],
        }, index=posts)

        res['pkdf']=pd.DataFrame({
            "Taux d'approbation des Avis":tca["Taux d'approbation des Avis"],"OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":pv_conf["OT CONFIME"],"OT_COR_EGAL":pv_cor["OT_COR_EGAL"],"OT Fiabilité":fiab_s,"Total Avis de Panne":avpan_s,
        }, index=posts)

        qvals=res['ckdf'][QK]; pvals=res['pkdf'][PK]
        qscore_list=[]
        for kpi in QK:
            cible=CIBLE[kpi]; val=qvals[kpi]
            if kpi in LOWER_BETTER: score=max(0,min(20,20*(1-max(0,val-cible)/max(cible,1))))
            else: score=max(0,min(20,20*val/max(cible,1)))
            qscore_list.append(score)
        res['ckdf']["Score Qualité"]=pd.Series(qscore_list,index=posts)
        
        pscore_list=[]
        for kpi in PK:
            cible=CIBLE[kpi]; val=pvals[kpi]
            if kpi in LOWER_BETTER: score=max(0,min(12.5,12.5*(1-max(0,val-cible)/max(cible,1))))
            else: score=max(0,min(12.5,12.5*val/max(cible,1)))
            pscore_list.append(score)
        res['pkdf']["Score Performance"]=pd.Series(pscore_list,index=posts)
        res['pkdf']["Score Qualité"]=res['ckdf'].pop("Score Qualité")

        res['an_pivot']=an; res['pr_pivot']=pr; res['pl_pivot']=pl; res['ex_pivot']=ex
        res['la_pivot']=la; res['pc_pivot']=pc; res['plc_pivot']=plc; res['tca_pivot']=tca
        return res

    def html_kpi_table(kpi_df, cibles, table_class):
        cols=["Poste de travail"]+list(kpi_df.columns)
        h=f'<table class="tw {table_class}"><thead><tr>'+''.join(f'<th>{c}</th>' for c in cols)+'</tr></thead><tbody>'
        for poste,row in kpi_df.iterrows():
            h+=f'<tr><td style="font-weight:700">{poste}</td>'
            for kpi in kpi_df.columns:
                val=row[kpi]; cible=cibles.get(kpi,100)
                if isinstance(val, (int, float, np.integer, np.floating)):
                    val = float(val)
                    if kpi in LOWER_BETTER: color="color:#dc2626;" if val>cible else "color:#10b981;"
                    else: color="color:#dc2626;" if val<cible else "color:#10b981;"
                    h+=f'<td style="text-align:center;{color}font-weight:600">{val:.1f}</td>'
                else: h+=f'<td style="text-align:center">{val}</td>'
            h+='</tr>'
        h+='<tr class="cb"><td>Total</td>'
        for kpi in kpi_df.columns:
            if kpi_df[kpi].dtype in ['float64','int64']: h+=f'<td style="text-align:center">{kpi_df[kpi].mean():.1f}</td>'
            else: h+=f'<td style="text-align:center"></td>'
        h+='</tr></tbody></table>'
        return h

    def html_anom_table(anom_df, kpi_list, table_class):
        cols=["Poste de travail"]+kpi_list+["Total"]
        h=f'<table class="tw {table_class}"><thead><tr>'+''.join(f'<th>{c}</th>' for c in cols)+'</tr></thead><tbody>'
        for poste,row in anom_df.iterrows():
            h+=f'<tr><td style="font-weight:700">{poste}</td>'
            total_anom=0
            for kpi in kpi_list:
                val=int(row.get(kpi,0)); total_anom+=val
                color="color:#dc2626;font-weight:800;" if val>0 else "color:#10b981;"
                h+=f'<td style="text-align:center;{color}">{val}</td>'
            h+=f'<td style="text-align:center;font-weight:800;color:#1e293b">{total_anom}</td></tr>'
        h+='<tr class="cb"><td>Total Anomalies</td>'
        grand_total=0
        for kpi in kpi_list:
            sum_val=int(anom_df[kpi].sum()); grand_total+=sum_val
            h+=f'<td style="text-align:center">{sum_val}</td>'
        h+=f'<td style="text-align:center">{grand_total}</td></tr></tbody></table>'
        return h

    def html_action_table(ano_rows):
        if not ano_rows: return '<div class="es">Aucune anomalie détectée. Tous les KPIs sont à leur cible.</div>'
        cols=["Poste de travail","KPI","Valeur (%)","Cible (%)","Ecart","Nombre d'anomalies","Action","Responsable"]
        h='<table class="plan-action-table"><thead><tr>'+''.join(f'<th>{c}</th>' for c in cols)+'</tr></thead><tbody>'
        for r in ano_rows:
            h+=f'<tr><td>{r["Poste de travail"]}</td><td>{r["KPI"]}</td><td>{r["Valeur (%)"]}</td><td>{r["Cible (%)"]}</td><td>{r["Ecart"]}</td><td style="font-weight:800;color:#dc2626">{r["Nombre d\'anomalies"]}</td><td>{r["Action"]}</td><td>{r["Responsable"]}</td></tr>'
        h+='</tbody></table>'
        return h

    # ================= RECHERCHE AUTOMATIQUE DES FICHIERS =================
    ot_path = next((f for f in os.listdir() if f.lower() == "ot.xlsx"), None)
    av_path = next((f for f in os.listdir() if f.lower() == "avis.xlsx"), None)

    with st.sidebar:
        logo_b64 = get_logo_base64()
        if logo_b64:
            st.markdown(f'<div style="text-align:center;padding:10px 0"><img src="data:image/png;base64,{logo_b64}" class="logo" style="height:60px;width:auto"></div>', unsafe_allow_html=True)
        
        if not ot_path or not av_path:
            st.error("Fichiers introuvables. Vérifiez que vos fichiers s'appellent bien **OT.xlsx** et **avis.xlsx**.")
            st.markdown(f"**Fichiers actuels dans `{os.getcwd()}` :**")
            st.code("\n".join(os.listdir()))
            st.stop()

        try:
            with open(ot_path, "rb") as f: ot_bytes = f.read()
            with open(av_path, "rb") as f: av_bytes = f.read()
            
            df_raw, av_raw, all_posts, now_ts = prepare_data(ot_bytes, av_bytes, fichier_date)
            st.success(f"Fichiers chargés : {len(df_raw)} OTs, {len(av_raw)} Avis")
            
            selected_posts = st.multiselect("Filtrer par Poste de travail", all_posts, default=all_posts)
            posts_filt = [p for p in all_posts if p in selected_posts]
            
            if not posts_filt:
                st.warning("Veuillez sélectionner au moins un poste.")
                st.stop()
                
            df_filt = df_raw[df_raw["Poste travail princ."].isin(posts_filt)].copy()
            av_filt = av_raw[av_raw["Poste travail princ."].isin(posts_filt)].copy()
            kdata = calc_kpis(df_filt, av_filt, now_ts, posts_filt)
            
            ckdf = kdata['ckdf']
            pkdf = kdata['pkdf']
            anom_df = kdata['anom']
            
            vue = st.radio("Affichage", ["Qualité (Performance)", "Performance (Qualité)"], index=0)
            is_quality_view = "Qualité" in vue
            
            tab1, tab2, tab3, tab4 = st.tabs(["Indicateurs", "Plans d'Actions", "Journal des Variations", "Historique & Export"])
            
            with tab1:
                st.markdown(f'<div class="mh"><h1>DASHBOARD KPI</h1><div class="db">{fichier_date}</div></div>', unsafe_allow_html=True)
                
                if is_quality_view:
                    kpi_list_show = [k for k in QK if k in ckdf.columns]
                    kpi_df_show = ckdf[kpi_list_show]
                    t_class = "qt"
                    c1_val, c1_lab = ckdf[kpi_list_show].mean().mean(), "Moyenne KPIs Qualité"
                    c2_val, c2_lab = anom_df[kpi_list_show].sum().sum(), "Total Anomalies Qualité"
                    c3_val, c3_lab = len(posts_filt), "Postes Sélectionnés"
                else:
                    kpi_list_show = [k for k in PK if k in pkdf.columns]
                    kpi_df_show = pkdf[kpi_list_show]
                    t_class = "pt"
                    c1_val, c1_lab = pkdf[kpi_list_show].mean().mean(), "Moyenne KPIs Performance"
                    c2_val, c2_lab = anom_df[kpi_list_show].sum().sum(), "Total Anomalies Perf"
                    c3_val, c3_lab = len(posts_filt), "Postes Sélectionnés"
                    
                st.markdown(f'<div class="cr" style="grid-template-columns: repeat(3, 1fr)"><div class="cc c1"><div class="cv">{c1_val:.1f}%</div><div class="cl">{c1_lab}</div></div><div class="cc c4"><div class="cv">{c2_val}</div><div class="cl">{c2_lab}</div></div><div class="cc c5"><div class="cv">{c3_val}</div><div class="cl">{c3_lab}</div></div></div>', unsafe_allow_html=True)
                
                st.markdown(f'<div class="stl">Tableau des Indicateurs</div>', unsafe_allow_html=True)
                st.markdown(html_kpi_table(kpi_df_show, CIBLE, t_class), unsafe_allow_html=True)
                
                st.markdown('<div class="stl">Nombre d\'anomalies par KPI et Poste</div>', unsafe_allow_html=True)
                st.markdown(html_anom_table(anom_df, kpi_list_show, "anom"), unsafe_allow_html=True)
                
                st.markdown('<div class="stl">Analyses Graphiques</div>', unsafe_allow_html=True)
                col_g1, col_g2 = st.columns(2)
                with col_g1: show_pie_pair(kdata['an_pivot'], "Correctifs (Global)")
                with col_g2: 
                    show_simple_pie(kdata['pr_pivot'][["<1 mois", ">3 mois", "1 mois < <3 mois"]], "Age Préparation (CRÉÉ)", keep_non_carac=False)
                    show_simple_pie(kdata['pc_pivot'][["CARACTERISE", "NON CARACTERISE"]], "Caractérisation Préparation", keep_non_carac=True)

            with tab2:
                st.markdown('<div class="stl">Plans d\'Actions (Source unique : Tableau des Anomalies)</div>', unsafe_allow_html=True)
                ano_rows = []
                for poste in posts_filt:
                    for kpi in ALL_KPI:
                        val = ckdf.loc[poste, kpi] if kpi in ckdf.columns else (pkdf.loc[poste, kpi] if kpi in pkdf.columns else None)
                        if val is None: continue
                        cible = CIBLE[kpi]
                        is_anom = (kpi in LOWER_BETTER and val > cible) or (kpi not in LOWER_BETTER and val < cible)
                        if is_anom:
                            ano_rows.append({
                                "Poste de travail": poste, "KPI": kpi, "Valeur (%)": round(float(val), 1),
                                "Cible (%)": cible, "Ecart": round(float(val) - cible, 1),
                                "Nombre d'anomalies": int(anom_df.loc[poste, kpi]),
                                "Action": ACT_MAP.get(kpi, ""), "Responsable": KPI_RESP_MAP.get(kpi, "")
                            })
                st.markdown(html_action_table(ano_rows), unsafe_allow_html=True)

            with tab3:
                st.markdown('<div class="stl">Journal des Variations</div>', unsafe_allow_html=True)
                hist_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
                if os.path.exists(hist_path):
                    hist_df = load_historical_kpis(hist_path)
                    var_df = calculate_variations(hist_df)
                    if not var_df.empty:
                        journal_df = generate_journal(var_df)
                        if not journal_df.empty: st.dataframe(journal_df, use_container_width=True, hide_index=True)
                        else: st.markdown('<div class="es">Aucune variation significative (>5%) détectée.</div>', unsafe_allow_html=True)
                    else: st.markdown('<div class="es">Pas assez d\'historique pour calculer les variations.</div>', unsafe_allow_html=True)
                else: st.markdown('<div class="es">Aucun historique trouvé. Exportez les données une première fois.</div>', unsafe_allow_html=True)

            with tab4:
                st.markdown('<div class="stl">Historique & Export Excel</div>', unsafe_allow_html=True)
                if st.button("Sauvegarder et Exporter les KPIs dans Excel", key="export_btn"):
                    def df_to_rows(df, poste_col="Poste de travail"):
                        rows = []
                        for idx, row in df.iterrows():
                            r = {poste_col: idx}
                            for c in df.columns: r[c] = round(row[c], 2) if isinstance(row[c], float) else row[c]
                            rows.append(r)
                        return rows
                    p_rows = df_to_rows(pkdf); p_cols = ["Poste de travail"] + list(pkdf.columns)
                    q_rows = df_to_rows(ckdf); q_cols = ["Poste de travail"] + list(ckdf.columns)
                    ano_p_r, ano_q_r = [], []
                    for r in ano_rows:
                        if r["KPI"] in PK: ano_p_r.append(r)
                        elif r["KPI"] in QK: ano_q_r.append(r)
                    ano_cols = ["Poste de travail","KPI","Valeur (%)","Cible (%)","Ecart","Nombre d'anomalies","Action","Responsable"]
                    save_kpis_to_excel(p_rows, p_cols, q_rows, q_cols, ano_p_r, ano_cols, ano_q_r, ano_cols, fichier_date)
                    st.success("Fichier Excel mis à jour avec succès dans le dossier /kpis !")
                if os.path.exists(os.path.join("kpis", "indicateurs_kpis.xlsx")):
                    with open(os.path.join("kpis", "indicateurs_kpis.xlsx"), "rb") as f:
                        st.download_button("Télécharger le fichier Historique", f, file_name="Historique_KPIs.xlsx", mime="application/vnd.ms-excel")

        except Exception as e:
            st.error(f"Erreur lors du traitement des données : {e}")

if __name__ == "__main__":
    main()
