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

def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    .main .block-container{max-width:100%!important;width:100%!important;padding-left:0.5rem!important;padding-right:0.5rem!important}
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{
        --primary:#1e3a5f;--primary-light:#2c5282;--success:#10b981;--success-dark:#059669;
        --warning:#f59e0b;--warning-dark:#d97706;--danger:#ef4444;--danger-dark:#dc2626;
        --info:#3b82f6;--border:#e2e8f0;--radius:10px;
    }
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
    .cc.c5{border-left-color:#3b82f6}.cc.c5 .cv{color:#2563eb}
    .cc.c6{border-left-color:#06b6d4}.cc.c6 .cv{color:#0891b2}
    .cc.c7{border-left-color:#f59e0b}.cc.c7 .cv{color:#d97706}
    .cc.c8{border-left-color:#f97316}.cc.c8 .cv{color:#ea580c}
    .stl{font-size:16px;font-weight:800;color:var(--primary);margin:10px 0 5px 0;padding-left:12px;border-left:4px solid var(--info)}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:13px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--primary);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:6px 8px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2563eb,#3b82f6)}
    .tw.pt thead th{background:linear-gradient(135deg,#059669,#10b981)}
    .tw.at thead th{background:linear-gradient(135deg,#dc2626,#ef4444)}
    .tw.st thead th{background:linear-gradient(135deg,#d97706,#f59e0b)}
    .tw thead th:first-child{z-index:11;left:0}
    .tw tbody td:first-child{position:sticky;left:0;background:#fff;z-index:5;border-right:1px solid var(--border);color:#1e293b !important}
    .tw tbody tr:nth-child(even) td:first-child{background:#f8fafc}
    .tw tbody tr:hover td:first-child{background:#eff6ff}
    .tw tbody td{padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;color:#1e293b !important}
    .tw tbody tr:nth-child(even) td{background:#f8fafc}
    .tw tbody tr:hover td{background:#eff6ff!important}
    .cb td{background:#2563eb!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .plan-action-table{width:100%;border-collapse:collapse;font-family:Inter,sans-serif;font-size:12px;border:1px solid #cbd5e1}
    .plan-action-table th{background:#1e3a5f;color:#fff;font-weight:700;padding:8px 6px;border:1px solid #1e3a5f}
    .plan-action-table td{padding:6px 8px;border:1px solid #cbd5e1;text-align:center;vertical-align:middle}
    .plan-action-table td:first-child{text-align:left;font-weight:800}
    .stTabs [data-baseweb="tab-list"]{gap:6px;background:#e2e8f0;padding:6px;border-radius:8px;margin-bottom:8px}
    .stTabs [data-baseweb="tab"]{border-radius:6px;padding:12px 22px;font-weight:700;font-size:20px;line-height:1.5;min-height:48px}
    .stTabs [data-baseweb="tab"] span,.stTabs [data-baseweb="tab"] > div{font-size:22px !important}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--primary)!important;box-shadow:0 3px 8px rgba(0,0,0,.1);font-size:21px}
    .stTabs [data-baseweb="tab"] svg{width:22px;height:22px}
    .ca{background:#fff;border-radius:var(--radius);padding:12px;margin-top:6px;border:1px solid var(--border);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:14px;font-weight:700;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid var(--border)}
    .car{display:flex;align-items:center;margin-bottom:6px;font-size:12px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:260px;font-weight:600;color:var(--primary);text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:26px;background:#edf2f7;border-radius:4px;overflow:visible;position:relative}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .target-mark{position:absolute;top:-4px;bottom:-4px;width:3px;background:var(--info)!important;z-index:20;transform:translateX(-50%);box-shadow:0 0 6px rgba(59,130,246,.9),0 0 2px rgba(0,0,0,.4);border-radius:2px}
    .car .cav-out{font-size:12px;font-weight:800;color:#1e293b;min-width:55px;text-align:right;padding-left:6px}
    .car .cav-tgt{font-size:10px;font-weight:700;color:#1e293b;min-width:42px;text-align:right;padding-left:4px;opacity:.7}
    .gbr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f1f5f9}
    .gbr:last-child{border:none}
    .gbr-l{width:160px;font-weight:600;color:#1e293b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;padding-right:10px}
    .gbr-g{display:flex;align-items:center;gap:4px;flex:1;position:relative}
    .gbr-target{position:absolute;left:90%;top:-4px;bottom:-4px;width:3px;background:var(--primary)!important;z-index:10;box-shadow:0 0 6px rgba(30,58,95,.8);border-radius:2px}
    .gbr-target-label{position:absolute;left:90%;top:-20px;transform:translateX(-50%);font-size:9px;font-weight:800;color:#fff;background:var(--primary)!important;padding:1px 5px;border-radius:3px;white-space:nowrap;z-index:11;box-shadow:0 1px 3px rgba(0,0,0,.2)}
    .gbr-w{flex:1;height:22px;background:#f1f5f9;border-radius:3px;overflow:hidden}
    .gbr-f{height:100%;border-radius:3px}
    .gbr-v{font-size:11px;font-weight:800;min-width:48px;text-align:right;color:#1e293b}
    .gbr-legend{display:flex;gap:14px;margin-bottom:10px;font-size:12px;font-weight:700;align-items:center}
    .gbr-legend span{display:flex;align-items:center;gap:5px}
    .gbr-legend i{display:inline-block;width:14px;height:14px;border-radius:2px}
    .gbr-legend .target-icon{display:inline-block;width:3px;height:14px;background:var(--primary)!important;border-radius:1px;box-shadow:0 0 3px rgba(30,58,95,.6)}
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg>div{background:#fff;border-radius:var(--radius);padding:8px 10px;border:1px solid var(--border)}
    .cg .ct{font-size:13px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--border)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f1f5f9}
    .cgr:last-child{border:none}
    .cgr .rk{width:18px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1e293b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:55px;text-align:right}
    .dgrid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--primary),var(--primary-light));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,#1e40af 0%,#1e3a8a 50%,#1e3a5f 100%)!important}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label,div[data-testid="stSidebar"] .stTextInput label{color:rgba(255,255,255,.9)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.1);border-radius:6px;padding:5px 10px;margin-bottom:5px;border:1px solid rgba(255,255,255,.15)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div,div[data-testid="stSidebar"] .stTextInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#64748b;font-size:14px}
    [data-testid="stHeaderActionElements"]{display:none !important}
    [data-testid="stActionButtonContainer"]{display:none !important}
    .footer{text-align:center;margin-top:30px;padding:15px;color:#64748b;font-size:13px;border-top:1px solid var(--border);font-weight:600}
    .anom-detail-box{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:13px}
    .anom-detail-box .anom-kpi{font-weight:800;color:#1e3a5f;font-size:14px}
    .anom-detail-box .anom-count{display:inline-block;background:#ef4444;color:#fff;font-weight:800;padding:2px 10px;border-radius:12px;font-size:12px;margin-left:8px}
    .anom-detail-box .anom-meta{color:#64748b;font-size:12px;margin-top:4px}
    @media(max-width:768px){
        .cr{grid-template-columns:repeat(2,1fr)}
        .mh{padding:8px 10px;gap:8px}
        .mh h1{font-size:18px}
        .mh .logo{height:35px;max-width:70px}
        .mh .db{font-size:11px;padding:2px 8px}
        .cg,.dgrid{grid-template-columns:1fr}
        .car{flex-wrap:wrap;gap:2px}
        .car .cal{width:100%;text-align:left;padding-right:0;margin-bottom:2px}
        .car .cab{flex:1 1 70%}
        .car .cav-out,.car .cav-tgt{flex:1 1 15%;min-width:40px}
        .gbr{flex-direction:column;align-items:flex-start;gap:4px}
        .gbr-l{width:100%;margin-bottom:2px}
        .gbr-g{width:100%;flex-wrap:wrap}
        .gbr-w{flex:1 1 45%}
        .gbr-v{flex:1 1 10%;min-width:40px}
        .tw{font-size:10px}
        .tw thead th,.tw tbody td{padding:3px 4px}
        .stl{font-size:13px}
        .stTabs [data-baseweb="tab"]{padding:8px 12px;font-size:15px}
        .stTabs [data-baseweb="tab"] span{font-size:16px !important}
    }
    @media print {
        section[data-testid="stSidebar"],header[data-testid="stHeader"],div[data-testid="stToolbar"],div[data-testid="stHeaderActionElements"],footer,.stDeployButton,#MainMenu{display:none !important}
        .main .block-container{padding-top:0!important;padding-left:0!important;padding-right:0!important;max-width:100%!important}
        .stButton,.stDownloadButton{display:none !important}
        *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
        .tw{page-break-inside:avoid;overflow:visible!important}
    }
    </style>""",unsafe_allow_html=True)

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

    def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)
    def cpiv(df,f,c,p):
        return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)

    def build_statut_pivot(df_sub, posts):
        if df_sub.empty:
            return pd.DataFrame(index=posts, columns=["CRÉÉ","LANC","CLOT","TCLO","Total"]).fillna(0).astype(int)
        piv=pd.pivot_table(df_sub, index="Poste travail princ.", columns="Statut OT", values="Ordre", aggfunc="count", fill_value=0)
        for s in ["CRÉÉ","LANC","CLOT","TCLO"]:
            if s not in piv.columns: piv[s]=0
        piv["Total"]=piv[["CRÉÉ","LANC","CLOT","TCLO"]].sum(axis=1)
        return piv.reindex(posts, fill_value=0).fillna(0).astype(int)

    def html_statut_pivot(piv_df, table_class):
        cols=["Poste de travail","CRÉÉ","LANC","CLOT","TCLO","Total"]
        sc={"CRÉÉ":"background:#fef3c7;color:#92400e;font-weight:600;","LANC":"background:#dbeafe;color:#1e40af;font-weight:600;","CLOT":"background:#d1fae5;color:#065f46;font-weight:600;","TCLO":"background:#a7f3d0;color:#064e3b;font-weight:600;","Total":"background:#ede9fe;color:#5b21b6;font-weight:700;"}
        h='<table class="tw %s"><thead><tr>'%table_class+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for poste,row in piv_df.iterrows():
            h+='<tr><td style="font-weight:700">%s</td>'%poste
            for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
                h+='<td style="text-align:center;%s">%d</td>'%(sc[c], int(row.get(c,0)))
            h+='<td style="text-align:center;%s">%d</td>'%(sc["Total"], int(row.get("Total",0)))
            h+='</tr>'
        h+='<tr class="tr"><td style="font-weight:800">Total</td>'
        for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
            h+='<td style="text-align:center;font-weight:800;%s">%d</td>'%(sc[c], int(piv_df[c].sum()))
        h+='<td style="text-align:center;font-weight:800;%s">%d</td>'%(sc["Total"], int(piv_df["Total"].sum()))
        h+='</tr></tbody></table>'
        return h

    def show_pie_pair(piv_df, title_prefix):
        global_counts = piv_df[["CRÉÉ","LANC","CLOT","TCLO"]].sum()
        global_counts = global_counts[global_counts > 0]
        realised = global_counts.get("CLOT", 0) + global_counts.get("TCLO", 0)
        not_realised = global_counts.sum() - realised
        if global_counts.empty:
            st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True); return
        colors = ["#8b5cf6", "#f59e0b", "#10b981", "#3b82f6"]
        fig = make_subplots(rows=1, cols=2, specs=[[{"type":"domain"},{"type":"domain"}]], subplot_titles=(f"{title_prefix} — Par Statut OT", f"{title_prefix} — Réalisés vs Non Réalisés"))
        fig.add_trace(go.Pie(labels=global_counts.index, values=global_counts.values, hole=0.4, textinfo='percent+label', texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', textposition='inside', insidetextorientation='radial', textfont=dict(size=14, color='white', family='Inter, sans-serif'), marker=dict(colors=colors, line=dict(color='#FFFFFF', width=3))), 1, 1)
        pie2_data = pd.Series([realised, not_realised], index=["Réalisés (CLOT+TCLO)", "Non Réalisés"])
        fig.add_trace(go.Pie(labels=pie2_data.index, values=pie2_data.values, hole=0.5, textinfo='percent+label', texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', textposition='inside', insidetextorientation='radial', textfont=dict(size=14, color='white', family='Inter, sans-serif'), marker=dict(colors=["#10b981", "#8b5cf6"], line=dict(color='#FFFFFF', width=3))), 1, 2)
        fig.update_layout(margin=dict(t=80, b=20, l=20, r=20), height=450, legend=dict(orientation="h", yanchor="bottom", y=-0.12, x=0.5, xanchor="center"))
        st.plotly_chart(fig, use_container_width=True)

    def show_simple_pie(piv_df, title, keep_non_carac=False):
        if not keep_non_carac and "NON CARACTERISE" in piv_df.columns:
            piv_df = piv_df.drop(columns=["NON CARACTERISE"])
        counts = piv_df.sum(); counts = counts[counts > 0]
        if counts.empty:
            st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True); return
        color_map = {"CARACTERISE": "#10b981", "NON CARACTERISE": "#f97316"}
        type_palette = ['#3b82f6','#10b981','#f59e0b','#8b5cf6','#06b6d4','#14b8a6','#6366f1','#0ea5e9','#d946ef','#a855f7']
        colors = []; pi = 0
        for c in counts.index:
            cs = str(c)
            if cs in color_map: colors.append(color_map[cs])
            else: colors.append(type_palette[pi % len(type_palette)]); pi += 1
        total_sum = counts.sum()
        pull_list = [0.05 if (v/total_sum)*100 < 10 else 0 for v in counts.values]
        fig = go.Figure(go.Pie(labels=counts.index, values=counts.values, hole=0.4, sort=False, textinfo="percent", textposition="outside", pull=pull_list, marker=dict(colors=colors, line=dict(color="white", width=2))))
        fig.update_traces(hovertemplate="<b>%{label}</b><br>Nombre : %{value}<br>Pourcentage : %{percent}<extra></extra>", textfont=dict(size=13, family='Inter, sans-serif'))
        fig.update_layout(title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)), height=500, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"), margin=dict(t=80, b=80, l=40, r=40))
        st.plotly_chart(fig, use_container_width=True)

    def is_green(kpi, val):
        try: v = float(val)
        except Exception: return False
        if kpi == "TAUX_REALISATION_CORRECTIF/PT": return v >= 95
        if kpi in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return v >= 80
        if kpi in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return v <= 5
        if kpi in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return v <= 15
        if kpi in ["Performance Graissage","Performance Inspection"]: return v >= 95
        if kpi == "Performance Appels Systématiques": return v >= 85
        if kpi == "appel avis approuvé": return v >= 95
        if kpi == "OT_COR_EGAL": return v >= 95
        if kpi in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT Fiabilité","Total Avis de Panne"]: return v >= 100
        return False

    def get_bar_color(kpi, val):
        return "#38a169" if is_green(kpi, val) else "#e53e3e"

    def get_anomalous_items(df_full, av_full, poste, kpi, now_ts):
        items = []
        df_p = df_full[df_full["Poste travail princ."] == poste]
        av_p = av_full[av_full["Poste travail princ."] == poste]
        if kpi == "TAUX_REALISATION_CORRECTIF/PT":
            filt = (df_p["Nº appel pl.entret."].fillna(0)==0) & (df_p["Contient SOPL"]==1)
            anom = df_p[filt & ~df_p["Statut OT"].isin(["CLOT","TCLO"])]
            items = anom["Ordre"].dropna().tolist()
        elif kpi.startswith("OT préparation"):
            sub = df_p[(df_p["Statut OT"]=="CRÉÉ") & (df_p["Statut utilisateur"].str.contains("CRPR",na=False))]
            if kpi == "OT préparation <1 mois": anom = sub[sub["ap"] > 1]
            elif kpi == "OT préparation >3 mois": anom = sub[sub["ap"] >= 3]
            elif kpi == "OT préparation 1mois< <3mois": anom = sub[(sub["ap"] > 1) & (sub["ap"] < 3)]
            else: anom = pd.DataFrame()
            items = anom["Ordre"].dropna().tolist()
        elif kpi.startswith("OT planification"):
            sub = df_p[(df_p["Statut OT"]=="LANC") & (df_p["Statut utilisateur"].str.contains("ATPL",case=False,na=False))]
            if kpi == "OT planification <1 mois": anom = sub[sub["alp"] > 1]
            elif kpi == "OT planification >3 mois": anom = sub[sub["alp"] >= 3]
            elif kpi == "OT planification 1mois< <3mois": anom = sub[(sub["alp"] > 1) & (sub["alp"] < 3)]
            else: anom = pd.DataFrame()
            items = anom["Ordre"].dropna().tolist()
        elif kpi.startswith("OT exécution"):
            sub = df_p[(df_p["Statut OT"]=="LANC") & (df_p["Contient SOPL"]==1)]
            if kpi == "OT exécution <1 mois": anom = sub[sub["aex"] > 1]
            elif kpi == "OT exécution >3 mois": anom = sub[sub["aex"] >= 3]
            elif kpi == "OT exécution 1mois< <3mois": anom = sub[(sub["aex"] > 1) & (sub["aex"] < 3)]
            else: anom = pd.DataFrame()
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "Performance Graissage":
            base = (df_p["_tw_num"]==350) & (df_p["Contient SOPL"]==1) & (df_p["Date de début planifiée"].notna()) & (df_p["Date de début planifiée"]<=now_ts)
            anom = df_p[base & ~df_p["Statut OT"].isin(["CLOT","TCLO"])]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "Performance Inspection":
            base = (df_p["_tw_num"].isin([290,300,310])) & (df_p["Contient SOPL"]==1) & (df_p["Date de début planifiée"].notna()) & (df_p["Date de début planifiée"]<=now_ts)
            anom = df_p[base & ~df_p["Statut OT"].isin(["CLOT","TCLO"])]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "Performance Appels Systématiques":
            base = (df_p["_tw_num"]==360) & (df_p["Contient SOPL"]==1) & (df_p["Date de début planifiée"].not_na() if hasattr(df_p["Date de début planifiée"], 'not_na') else df_p["Date de début planifiée"].notna()) & (df_p["Date de début planifiée"]<=now_ts)
            anom = df_p[base & ~df_p["Statut OT"].isin(["CLOT","TCLO"])]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "appel avis approuvé":
            anom = av_p[~av_p["Statut utilisateur"].str.contains("APRV",na=False)]
            items = anom["Avis"].dropna().tolist()
        elif kpi == "OT LANC ESTIME":
            anom = df_p[(df_p["Statut OT"]=="LANC") & (df_p["Total coûts budgétés"].fillna(0)==0)]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "Backlog préparation caractérisé":
            anom = df_p[(df_p["Statut OT"]=="CRÉÉ") & (df_p["Backlog preparation"]=="NON CARACTERISE")]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "Backlog planification caractérisé":
            anom = df_p[(df_p["Statut OT"]=="LANC") & (df_p["Backlog planification"]=="NON CARACTERISE")]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "OT CONFIME":
            anom = df_p[(df_p["Statut système"].str.contains("CLOT|TCLO",na=False)) & (~df_p["Statut système"].str.contains("CONF",na=False))]
            items = anom["Ordre"].dropna().tolist()
        elif kpi == "OT_COR_EGAL":
            anom = df_p[(df_p["Total coûts budgétés"].fillna(0) - df_p["Total coûts réels"].fillna(0)) != 0]
            items = anom["Ordre"].dropna().tolist()
        elif kpi in ["OT Fiabilité", "Total Avis de Panne"]:
            items = []
        return [str(int(i)) if isinstance(i, (int,float)) and not pd.isna(i) and float(i).is_integer() else str(i) for i in items if not pd.isna(i)]

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
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"]); pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0); pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)
        pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)),"alp",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"]); pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0); pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)
        ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"]); ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0); ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)
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
        ins_base=(df["_tw_num"].isin([290,300,310]))&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
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
            "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "Performance Graissage":g_df["Performance Graissage"],"Performance Inspection":ins_df["Performance Inspection"],"Performance Appels Systématiques":sys_df["Performance Appels Systématiques"],
            "appel avis approuvé":tca["appel avis approuvé"],"OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"],
            "OT Fiabilité":fiab_s,"Total Avis de Panne":avpan_s
        })
        return res

    def html_kpi_table(kpi_list, ckdf, posts, table_class):
        cols=["Poste de travail"]+kpi_list+["Total Général"]
        h='<table class="tw %s"><thead><tr>'%table_class+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        col_green_counts = {k: 0 for k in kpi_list}
        for p in posts:
            h+='<tr><td>%s</td>'%p
            row_green = 0
            for k in kpi_list:
                v=ckdf.loc[p,k] if p in ckdf.index else 0
                v=round(float(v),1) if not pd.isna(v) else 0
                green = is_green(k, v)
                if green: row_green += 1; col_green_counts[k] += 1
                bc="#38a169" if green else "#e53e3e"
                h+='<td style="text-align:center;background:%s;color:#fff;font-weight:700">%s%%</td>'%(bc,str(v))
            tg = round(row_green / len(kpi_list) * 100, 1) if kpi_list else 0
            tg_c = "#38a169" if tg >= 90 else "#e53e3e"
            h+='<td style="text-align:center;background:%s;color:#fff;font-weight:800;font-size:13px">%s%%</td>'%(tg_c, str(tg))
            h+='</tr>'
        h+='<tr class="cb"><td>MOYENNE</td>'
        for k in kpi_list:
            avg=round(ckdf[k].mean(),1) if not ckdf.empty else 0
            h+='<td>%s%%</td>'%str(avg)
        tg_col = round(sum(col_green_counts.values()) / (len(kpi_list) * len(posts)) * 100, 1) if posts and kpi_list else 0
        tg_col_c = "#38a169" if tg_col >= 90 else "#e53e3e"
        h+='<td style="background:%s;color:#fff;font-weight:800">%s%%</td>'%(tg_col_c, str(tg_col))
        h+='</tr></tbody></table>'
        return h

    def build_anomalies(kpi_list, ckdf, posts, df_full, av_full, now_ts):
        rows=[]; cols=["Poste de travail","KPI","Valeur","Cible","Ecart","Nb Anomalies","Statut"]
        for p in posts:
            if p not in ckdf.index: continue
            for k in kpi_list:
                v=float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
                cible=CIBLE.get(k,100)
                ecart=round(v-cible,1)
                if k in LOWER_BETTER:
                    statut="OK" if v<=cible else "ANOMALIE"
                else:
                    statut="OK" if v>=cible else "ANOMALIE"
                if statut=="ANOMALIE":
                    items = get_anomalous_items(df_full, av_full, p, k, now_ts)
                    nb_anom = len(items)
                    rows.append({"Poste de travail":p,"KPI":k,"Valeur":round(v,1),"Cible":cible,"Ecart":ecart,"Nb Anomalies":nb_anom,"Statut":"ANOMALIE","_items":items})
        return rows, cols

    def html_anomalies_table(rows, cols):
        if not rows: return '<div class="es">Aucune anomalie detectee</div>'
        h='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            nb = r.get("Nb Anomalies", 0)
            h+='<tr>'
            for c in cols:
                if c == "Nb Anomalies":
                    h+='<td style="text-align:center;font-weight:800;background:#fef2f2;color:#dc2626">%d</td>'%nb
                else:
                    h+='<td style="text-align:center">%s</td>'%r.get(c,"")
            h+='</tr>'
        h+='</tbody></table>'
        return h

    def html_bar_charts(kpi_list, ckdf, posts, section_title):
        h='<div class="ca"><div class="ct">%s</div>'%section_title
        for k in kpi_list:
            cible=CIBLE.get(k,100)
            max_val=max(120, ckdf[k].max()*1.1) if not ckdf.empty else 120
            tgt_pct=min((cible/max_val)*100, 100)
            h+='<div class="gbr-legend"><span><i style="background:#38a169"></i> Atteint</span><span><i style="background:#e53e3e"></i> Non atteint</span><span><span class="target-icon"></span> Cible (%s%%)</span></div>'%str(cible)
            h+='<div style="margin-bottom:16px">'
            for p in posts:
                if p not in ckdf.index: continue
                v=float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
                v=round(v,1)
                pct=min((v/max_val)*100, 100)
                bc=get_bar_color(k,v)
                h+='<div class="gbr"><div class="gbr-l" title="%s">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f" style="width:%s%%;background:%s"></div></div><div class="gbr-v">%s%%</div><div class="gbr-target" style="left:%s%%"></div><div class="gbr-target-label">%s%%</div></div></div>'%(p,p,str(pct),bc,str(v),str(tgt_pct),str(cible))
            h+='</div>'
        h+='</div>'
        return h

    def generate_plan_action_excel(all_plan_rows):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if all_plan_rows:
                df_plan = pd.DataFrame(all_plan_rows)
                df_plan.to_excel(writer, sheet_name='Plan Action', index=False)
                ws = writer.sheets['Plan Action']
                hf=Font(bold=True,color="FFFFFF",size=10); hfl=PatternFill(start_color="1E3A5F",end_color="1E3A5F",fill_type="solid")
                tb=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
                for col_idx in range(1, len(df_plan.columns)+1):
                    cell=ws.cell(row=1,column=col_idx); cell.font=hf; cell.fill=hfl; cell.alignment=Alignment(horizontal='center'); cell.border=tb
                for row_idx in range(2, len(df_plan)+2):
                    for col_idx in range(1, len(df_plan.columns)+1):
                        ws.cell(row=row_idx,column=col_idx).border=tb; ws.cell(row=row_idx,column=col_idx).alignment=Alignment(horizontal='center')
                for col in ws.columns:
                    ml=0; cl=col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value: ml=max(ml,len(str(cell.value)))
                        except Exception: pass
                    ws.column_dimensions[cl].width=min(max(ml+4,12),50)
            else:
                pd.DataFrame({"Info": ["Aucun plan d'action requis"]}).to_excel(writer, sheet_name='Plan Action', index=False)
        output.seek(0)
        return output.getvalue()

    # ============================================================
    # CHARGEMENT AUTOMATIQUE DES FICHIERS OT.XLSX, AVIS.XLSX, DATE.TXT
    # ============================================================
    ot_path = "OT.xlsx"
    av_path = "AVIS.xlsx"
    auto_loaded = False
    ot_bytes_auto = None
    av_bytes_auto = None

    if os.path.exists(ot_path) and os.path.exists(av_path):
        try:
            with open(ot_path, "rb") as f: ot_bytes_auto = f.read()
            with open(av_path, "rb") as f: av_bytes_auto = f.read()
            auto_loaded = True
        except Exception:
            auto_loaded = False

    # ---- SIDEBAR ----
    with st.sidebar:
        st.markdown('<div style="text-align:center;margin-bottom:15px"><div style="font-size:28px;font-weight:900;color:#fff">⚙️ PARAMETRES</div></div>',unsafe_allow_html=True)

        if auto_loaded:
            st.markdown('<div style="background:rgba(16,185,129,0.2);border:1px solid rgba(16,185,129,0.5);border-radius:8px;padding:10px 12px;margin-bottom:8px;text-align:center"><div style="font-size:18px;margin-bottom:4px">✅ Fichiers charges automatiquement</div><div style="font-size:11px;opacity:0.9">OT.xlsx • AVIS.xlsx • date.txt</div></div>', unsafe_allow_html=True)
            ot_file = None
            av_file = None
            st.markdown("**Fichiers locaux actifs :**")
            st.markdown("- 📄 `OT.xlsx`")
            st.markdown("- 📄 `AVIS.xlsx`")
            st.markdown("- 📄 `date.txt` → `%s`" % fichier_date)
            use_auto = st.checkbox("Utiliser les fichiers locaux", value=True, key="use_auto")
            if not use_auto:
                st.markdown("---")
                st.markdown("**Ou uploader manuellement :**")
                ot_file=st.file_uploader("Fichier OT (Excel)",type=["xlsx","xls"],key="ot_up")
                av_file=st.file_uploader("Fichier Avis (Excel)",type=["xlsx","xls"],key="av_up")
        else:
            st.markdown('<div style="background:rgba(239,68,68,0.2);border:1px solid rgba(239,68,68,0.5);border-radius:8px;padding:10px 12px;margin-bottom:8px;text-align:center"><div style="font-size:14px">⚠️ Fichiers locaux non trouvés</div><div style="font-size:11px;opacity:0.9">Placez OT.xlsx et AVIS.xlsx à côté du script</div></div>', unsafe_allow_html=True)
            ot_file=st.file_uploader("Fichier OT (Excel)",type=["xlsx","xls"],key="ot_up")
            av_file=st.file_uploader("Fichier Avis (Excel)",type=["xlsx","xls"],key="av_up")

        st.markdown("---")
        # Determiner les bytes a utiliser
        if auto_loaded and (not hasattr(st, 'session_state') or st.session_state.get("use_auto", True)):
            final_ot_bytes = ot_bytes_auto
            final_av_bytes = av_bytes_auto
        elif ot_file and av_file:
            final_ot_bytes = ot_file.read()
            final_av_bytes = av_file.read()
        else:
            final_ot_bytes = None
            final_av_bytes = None

        if final_ot_bytes and final_av_bytes:
            df_r, av_r, posts, now_ts = prepare_data(final_ot_bytes, final_av_bytes, fichier_date)
            if posts:
                sel_posts=st.multiselect("Postes de travail",posts,default=posts,key="sp")
            else:
                sel_posts=[]; st.warning("Aucun poste SF1/SF2 trouve")
        else:
            st.info("Veuillez charger les deux fichiers Excel pour commencer.")
            sel_posts=[]; df_r=None; av_r=None; posts=[]; now_ts=pd.Timestamp.now()

        st.markdown("---")
        show_hist=st.checkbox("Analyse historique",value=False,key="sh")
        hist_df=pd.DataFrame(); var_df=pd.DataFrame(); journal_df=pd.DataFrame(); top5=pd.DataFrame(); bottom5=pd.DataFrame()
        if show_hist:
            hist_file=st.file_uploader("Fichier historique KPIs",type=["xlsx","xls"],key="hf")
            if hist_file:
                hist_df=load_historical_kpis(io.BytesIO(hist_file.read()))
                var_df=calculate_variations(hist_df); journal_df=generate_journal(var_df); top5,bottom5=calculate_rankings(var_df)
        st.markdown("---")
        st.markdown('<div style="text-align:center;font-size:11px;color:rgba(255,255,255,.5);margin-top:10px">Dashboard KPI v2.0<br>Maintenance BM</div>',unsafe_allow_html=True)

    # ---- MAIN CONTENT ----
    if not final_ot_bytes or not final_av_bytes:
        st.markdown('<div style="min-height:60vh;display:flex;align-items:center;justify-content:center"><div style="text-align:center"><div style="font-size:80px;margin-bottom:20px">📊</div><h2 style="color:#1e3a5f;font-size:32px;font-weight:800">Dashboard KPI Maintenance</h2><p style="color:#64748b;font-size:18px;margin-top:10px">Chargez les fichiers OT et Avis dans la barre laterale pour demarrer</p></div></div>',unsafe_allow_html=True)
        st.stop()
    if not sel_posts:
        st.markdown('<div class="es">Aucun poste selectionne. Veuillez selectionner au moins un poste.</div>',unsafe_allow_html=True)
        st.stop()

    df_sel = df_r[df_r["Poste travail princ."].isin(sel_posts)].copy()
    av_sel = av_r[av_r["Poste travail princ."].isin(sel_posts)].copy()
    res = calc_kpis(df_sel, av_sel, now_ts, sel_posts)
    ckdf = res['ckdf']
    perf_score = round(ckdf[QK].mean().mean(), 1)
    qual_score = round(ckdf[PK].mean().mean(), 1)
    global_score = round((perf_score + qual_score) / 2, 1)
    total_ot = len(df_sel); total_av = len(av_sel)

    nb_anomalies_perf = 0; nb_anomalies_qual = 0
    for p in sel_posts:
        if p not in ckdf.index: continue
        for k in QK:
            v = float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
            cible = CIBLE.get(k,100)
            if k in LOWER_BETTER:
                if v > cible: nb_anomalies_perf += len(get_anomalous_items(df_sel, av_sel, p, k, now_ts))
            else:
                if v < cible: nb_anomalies_perf += len(get_anomalous_items(df_sel, av_sel, p, k, now_ts))
        for k in PK:
            v = float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
            cible = CIBLE.get(k,100)
            if k in LOWER_BETTER:
                if v > cible: nb_anomalies_qual += len(get_anomalous_items(df_sel, av_sel, p, k, now_ts))
            else:
                if v < cible: nb_anomalies_qual += len(get_anomalous_items(df_sel, av_sel, p, k, now_ts))

    logo_b64 = get_logo_base64()
    logo_html = '<img src="data:image/png;base64,%s" class="logo" alt="Logo">'%logo_b64 if logo_b64 else ''
    st.markdown('<div class="mh">%s<h1>📊 DASHBOARD KPI MAINTENANCE</h1><div class="db">📅 %s</div></div>'%(logo_html,fichier_date),unsafe_allow_html=True)

    perf_c = "#38a169" if perf_score>=90 else "#e53e3e"
    qual_c = "#38a169" if qual_score>=90 else "#e53e3e"
    glob_c = "#38a169" if global_score>=90 else "#e53e3e"
    anom_c = "#e53e3e" if (nb_anomalies_perf+nb_anomalies_qual)>0 else "#38a169"

    st.markdown('<div class="cr"><div class="cc c1"><div class="cv">%s</div><div class="cl">Total OT</div></div><div class="cc c2"><div class="cv" style="color:%s">%s%%</div><div class="cl">Score Performance</div></div><div class="cc c3"><div class="cv" style="color:%s">%s%%</div><div class="cl">Score Qualite</div></div><div class="cc c4"><div class="cv" style="color:%s">%s%%</div><div class="cl">Score Global</div></div></div>'%(str(total_ot),perf_c,str(perf_score),qual_c,str(qual_score),glob_c,str(global_score)),unsafe_allow_html=True)
    st.markdown('<div class="cr"><div class="cc c5"><div class="cv">%s</div><div class="cl">Total Avis</div></div><div class="cc c6"><div class="cv">%s</div><div class="cl">Postes Selectionnes</div></div><div class="cc c7"><div class="cv" style="color:%s">%s</div><div class="cl">Anomalies Perf.</div></div><div class="cc c8"><div class="cv" style="color:%s">%s</div><div class="cl">Anomalies Qual.</div></div></div>'%(str(total_av),str(len(sel_posts)),anom_c,str(nb_anomalies_perf),anom_c,str(nb_anomalies_qual)),unsafe_allow_html=True)

    ano_p_rows, ano_p_cols = build_anomalies(QK, ckdf, sel_posts, df_sel, av_sel, now_ts)
    ano_q_rows, ano_q_cols = build_anomalies(PK, ckdf, sel_posts, df_sel, av_sel, now_ts)
    p_rows = [{"Poste de travail":p} | {k:round(float(ckdf.loc[p,k]),1) if p in ckdf.index and not pd.isna(ckdf.loc[p,k]) else 0 for k in QK} | {"Score Performance":round(float(ckdf.loc[p,QK].mean()),1) if p in ckdf.index else 0} for p in sel_posts]
    q_rows = [{"Poste de travail":p} | {k:round(float(ckdf.loc[p,k]),1) if p in ckdf.index and not pd.isna(ckdf.loc[p,k]) else 0 for k in PK} | {"Score Qualite":round(float(ckdf.loc[p,PK].mean()),1) if p in ckdf.index else 0} for p in sel_posts]
    p_cols_list = ["Poste de travail"]+QK+["Score Performance"]
    q_cols_list = ["Poste de travail"]+PK+["Score Qualite"]
    save_kpis_to_excel(p_rows, p_cols_list, q_rows, q_cols_list, ano_p_rows, ano_p_cols, ano_q_rows, ano_q_cols, fichier_date)

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Performance","✅ Qualite","⚠️ Anomalies","📋 Plan d'Action","📈 Historique"])

    with tab1:
        st.markdown('<div class="stl">INDICATEURS DE PERFORMANCE</div>',unsafe_allow_html=True)
        st.markdown(html_kpi_table(QK, ckdf, sel_posts, "pt"), unsafe_allow_html=True)
        st.markdown('<div class="stl">GRAPHIQUES PERFORMANCE</div>',unsafe_allow_html=True)
        st.markdown(html_bar_charts(QK, ckdf, sel_posts, "Barres de progression - Performance"), unsafe_allow_html=True)
        st.markdown('<div class="stl">REPARTITION PAR STATUT OT (Correctifs)</div>',unsafe_allow_html=True)
        filt_corr = (df_sel["Nº appel pl.entret."].fillna(0)==0)&(df_sel["Contient SOPL"]==1)
        piv_statut = build_statut_pivot(df_sel[filt_corr], sel_posts)
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown(html_statut_pivot(piv_statut, "pt"), unsafe_allow_html=True)
        with col_s2:
            show_pie_pair(piv_statut, "Correctifs")

    with tab2:
        st.markdown('<div class="stl">INDICATEURS DE QUALITE</div>',unsafe_allow_html=True)
        st.markdown(html_kpi_table(PK, ckdf, sel_posts, "qt"), unsafe_allow_html=True)
        st.markdown('<div class="stl">GRAPHIQUES QUALITE</div>',unsafe_allow_html=True)
        st.markdown(html_bar_charts(PK, ckdf, sel_posts, "Barres de progression - Qualite"), unsafe_allow_html=True)
        st.markdown('<div class="stl">BACKLOG PREPARATION</div>',unsafe_allow_html=True)
        pc_piv = pd.pivot_table(df_sel[df_sel["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(sel_posts, fill_value=0)
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            show_simple_pie(pc_piv, "Caracterisation Backlog Preparation", keep_non_carac=True)
        with col_b2:
            tc_piv = pd.pivot_table(df_sel[df_sel["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Type Carac Prep", values="Ordre", aggfunc="count", fill_value=0).reindex(sel_posts, fill_value=0)
            show_simple_pie(tc_piv, "Par Type de Caracterisation")
        st.markdown('<div class="stl">BACKLOG PLANIFICATION</div>',unsafe_allow_html=True)
        plc_piv = pd.pivot_table(df_sel[df_sel["Statut OT"]=="LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(sel_posts, fill_value=0)
        col_b3, col_b4 = st.columns(2)
        with col_b3:
            show_simple_pie(plc_piv, "Caracterisation Backlog Planification", keep_non_carac=True)
        with col_b4:
            tcl_piv = pd.pivot_table(df_sel[df_sel["Statut OT"]=="LANC"], index="Poste travail princ.", columns="Type Carac Plan", values="Ordre", aggfunc="count", fill_value=0).reindex(sel_posts, fill_value=0)
            show_simple_pie(tcl_piv, "Par Type de Caracterisation")

    with tab3:
        st.markdown('<div class="stl">ANOMALIES PERFORMANCE</div>',unsafe_allow_html=True)
        st.markdown(html_anomalies_table(ano_p_rows, ano_p_cols), unsafe_allow_html=True)
        if ano_p_rows:
            st.markdown('<div class="stl">DETAIL DES ANOMALIES PERFORMANCE</div>',unsafe_allow_html=True)
            for r in ano_p_rows:
                p=r["Poste de travail"]; k=r["KPI"]; nb=r["Nb Anomalies"]; items=r.get("_items",[])
                with st.expander("🔍 %s — %s (%d anomalie(s))" % (p, k, nb)):
                    st.markdown('**Valeur :** %s%% | **Cible :** %s%% | **Ecart :** %s' % (r["Valeur"], r["Cible"], r["Ecart"]))
                    st.markdown('**Responsable :** %s' % KPI_RESP_MAP.get(k, "N/A"))
                    st.markdown('**Action :** %s' % ACT_MAP.get(k, "N/A"))
                    if items:
                        st.markdown('**Liste des %d élément(s) à corriger :**' % len(items))
                        st.code(", ".join(items[:5000]))
                    else:
                        st.markdown('*Aucun élément identifié dans le périmètre filtré*')
        st.markdown('<div class="stl">ANOMALIES QUALITE</div>',unsafe_allow_html=True)
        st.markdown(html_anomalies_table(ano_q_rows, ano_q_cols), unsafe_allow_html=True)
        if ano_q_rows:
            st.markdown('<div class="stl">DETAIL DES ANOMALIES QUALITE</div>',unsafe_allow_html=True)
            for r in ano_q_rows:
                p=r["Poste de travail"]; k=r["KPI"]; nb=r["Nb Anomalies"]; items=r.get("_items",[])
                with st.expander("🔍 %s — %s (%d anomalie(s))" % (p, k, nb)):
                    st.markdown('**Valeur :** %s%% | **Cible :** %s%% | **Ecart :** %s' % (r["Valeur"], r["Cible"], r["Ecart"]))
                    st.markdown('**Responsable :** %s' % KPI_RESP_MAP.get(k, "N/A"))
                    st.markdown('**Action :** %s' % ACT_MAP.get(k, "N/A"))
                    if items:
                        st.markdown('**Liste des %d élément(s) à corriger :**' % len(items))
                        st.code(", ".join(items[:5000]))
                    else:
                        st.markdown('*Aucun élément identifié dans le périmètre filtré*')
        total_ano = len(ano_p_rows) + len(ano_q_rows)
        ano_cc = "#e53e3e" if total_ano > 0 else "#38a169"
        st.markdown('<div style="text-align:center;padding:20px;background:#fff;border-radius:12px;border:2px solid %s;margin-top:10px"><div style="font-size:48px;font-weight:900;color:%s">%d</div><div style="font-size:16px;font-weight:700;color:#1e293b">Indicateurs en anomalie</div></div>'%(ano_cc,ano_cc,total_ano),unsafe_allow_html=True)

    with tab4:
        all_plan_rows = []
        st.markdown('<div class="stl">PLAN D\'ACTION PERFORMANCE</div>',unsafe_allow_html=True)
        has_perf_plan = False
        for p in sel_posts:
            if p not in ckdf.index: continue
            poste_anomalies = []
            for k in QK:
                v = float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
                cible = CIBLE.get(k,100)
                is_anom = (v > cible) if k in LOWER_BETTER else (v < cible)
                if is_anom:
                    items = get_anomalous_items(df_sel, av_sel, p, k, now_ts)
                    nb_anom = len(items)
                    poste_anomalies.append({"Poste de travail":p,"KPI":k,"Valeur (%)":round(v,1),"Cible (%)":cible,"Nb Anomalies":nb_anom,"Responsable":KPI_RESP_MAP.get(k,"N/A"),"Action":ACT_MAP.get(k,"N/A"),"_items":items})
                    all_plan_rows.append({"Poste de travail":p,"KPI":k,"Valeur (%)":round(v,1),"Cible (%)":cible,"Nb Anomalies":nb_anom,"Responsable":KPI_RESP_MAP.get(k,"N/A"),"Action":ACT_MAP.get(k,"N/A")})
            if poste_anomalies:
                has_perf_plan = True
                total_nb = sum(a["Nb Anomalies"] for a in poste_anomalies)
                with st.expander("🔧 %s — %d indicateur(s) en anomalie — %d élément(s) à corriger" % (p, len(poste_anomalies), total_nb)):
                    for a in poste_anomalies:
                        st.markdown('<div class="anom-detail-box"><span class="anom-kpi">%s</span><span class="anom-count">%d à corriger</span><div class="anom-meta">Valeur: %s%% | Cible: %s%% | 👤 %s</div><div class="anom-meta">📋 %s</div></div>' % (a["KPI"], a["Nb Anomalies"], a["Valeur (%)"], a["Cible (%)"], a["Responsable"], a["Action"]), unsafe_allow_html=True)
                        if a["_items"]:
                            with st.expander("Voir les %d élément(s)" % len(a["_items"])):
                                st.code(", ".join(a["_items"][:2000]))
        if not has_perf_plan:
            st.markdown('<div class="es">✅ Aucun plan d\'action requis en Performance — Tous les KPI sont atteints</div>',unsafe_allow_html=True)

        st.markdown('<div class="stl">PLAN D\'ACTION QUALITE</div>',unsafe_allow_html=True)
        has_qual_plan = False
        for p in sel_posts:
            if p not in ckdf.index: continue
            poste_anomalies = []
            for k in PK:
                v = float(ckdf.loc[p,k]) if not pd.isna(ckdf.loc[p,k]) else 0
                cible = CIBLE.get(k,100)
                is_anom = (v > cible) if k in LOWER_BETTER else (v < cible)
                if is_anom:
                    items = get_anomalous_items(df_sel, av_sel, p, k, now_ts)
                    nb_anom = len(items)
                    poste_anomalies.append({"Poste de travail":p,"KPI":k,"Valeur (%)":round(v,1),"Cible (%)":cible,"Nb Anomalies":nb_anom,"Responsable":KPI_RESP_MAP.get(k,"N/A"),"Action":ACT_MAP.get(k,"N/A"),"_items":items})
                    all_plan_rows.append({"Poste de travail":p,"KPI":k,"Valeur (%)":round(v,1),"Cible (%)":cible,"Nb Anomalies":nb_anom,"Responsable":KPI_RESP_MAP.get(k,"N/A"),"Action":ACT_MAP.get(k,"N/A")})
            if poste_anomalies:
                has_qual_plan = True
                total_nb = sum(a["Nb Anomalies"] for a in poste_anomalies)
                with st.expander("🔧 %s — %d indicateur(s) en anomalie — %d élément(s) à corriger" % (p, len(poste_anomalies), total_nb)):
                    for a in poste_anomalies:
                        st.markdown('<div class="anom-detail-box"><span class="anom-kpi">%s</span><span class="anom-count">%d à corriger</span><div class="anom-meta">Valeur: %s%% | Cible: %s%% | 👤 %s</div><div class="anom-meta">📋 %s</div></div>' % (a["KPI"], a["Nb Anomalies"], a["Valeur (%)"], a["Cible (%)"], a["Responsable"], a["Action"]), unsafe_allow_html=True)
                        if a["_items"]:
                            with st.expander("Voir les %d élément(s)" % len(a["_items"])):
                                st.code(", ".join(a["_items"][:2000]))
        if not has_qual_plan:
            st.markdown('<div class="es">✅ Aucun plan d\'action requis en Qualite — Tous les KPI sont atteints</div>',unsafe_allow_html=True)

        st.markdown("---")
        plan_bytes = generate_plan_action_excel(all_plan_rows)
        safe_date = fichier_date.replace("/","-")
        st.download_button(
            label="📥 Télécharger le Plan d'Action complet (Excel)",
            data=plan_bytes,
            file_name="plan_action_%s.xlsx" % safe_date,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with tab5:
        if show_hist and not hist_df.empty:
            st.markdown('<div class="stl">JOURNAL DES VARIATIONS SIGNIFICATIVES</div>',unsafe_allow_html=True)
            if journal_df.empty:
                st.markdown('<div class="es">Aucune variation significative detectee</div>',unsafe_allow_html=True)
            else:
                jcols = ["Date precedente","Date actuelle","Poste","Type","KPI","Valeur precedente","Valeur actuelle","Ecart","Ecart %","Sens"]
                st.dataframe(journal_df[jcols].copy(), use_container_width=True, height=400)
            st.markdown('<div class="stl">CLASSEMENT DES POSTES</div>',unsafe_allow_html=True)
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                st.markdown('<div style="text-align:center;font-weight:800;color:#059669;font-size:14px;margin-bottom:8px">🏆 TOP 5 AMELIORATION</div>',unsafe_allow_html=True)
                if not top5.empty: st.dataframe(top5, use_container_width=True, height=300)
                else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
            with col_r2:
                st.markdown('<div style="text-align:center;font-weight:800;color:#dc2626;font-size:14px;margin-bottom:8px">⚠️ TOP 5 DEGRADATION</div>',unsafe_allow_html=True)
                if not bottom5.empty: st.dataframe(bottom5, use_container_width=True, height=300)
                else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
            if not hist_df.empty:
                st.markdown('<div class="stl">SYNTHESE HISTORIQUE</div>',unsafe_allow_html=True)
                perf_hist = hist_df[hist_df["_section"]=="perf"].copy()
                qual_hist = hist_df[hist_df["_section"]=="qual"].copy()
                if not perf_hist.empty and "Poste de travail" in perf_hist.columns:
                    dl = sorted(perf_hist["Date"].unique())
                    if len(dl) > 1:
                        st.markdown('<div style="font-weight:700;color:#059669;margin-bottom:5px">Performance</div>',unsafe_allow_html=True)
                        st.markdown('<div style="font-size:12px;color:#64748b;margin-bottom:8px">Periode: %s → %s</div>'%(dl[0],dl[-1]),unsafe_allow_html=True)
                if not qual_hist.empty and "Poste de travail" in qual_hist.columns:
                    dlq = sorted(qual_hist["Date"].unique())
                    if len(dlq) > 1:
                        st.markdown('<div style="font-weight:700;color:#2563eb;margin-bottom:5px">Qualite</div>',unsafe_allow_html=True)
                        st.markdown('<div style="font-size:12px;color:#64748b;margin-bottom:8px">Periode: %s → %s</div>'%(dlq[0],dlq[-1]),unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Activez l\'analyse historique dans la barre laterale et chargez le fichier historique.</div>',unsafe_allow_html=True)

    st.markdown('<div class="footer">Dashboard KPI Maintenance BM — %s | Genere automatiquement | Tous droits reserves</div>'%fichier_date,unsafe_allow_html=True)

if __name__ == "__main__":
    main()
