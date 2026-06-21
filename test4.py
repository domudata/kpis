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

# === CIBLES MISES A JOUR ===
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
        --primary:#1e3a5f;
        --primary-light:#2c5282;
        --success:#10b981;
        --success-dark:#059669;
        --warning:#f59e0b;
        --warning-dark:#d97706;
        --danger:#ef4444;
        --danger-dark:#dc2626;
        --info:#3b82f6;
        --border:#e2e8f0;
        --radius:10px;
    }
    
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#f8fafc;font-family:'Inter',sans-serif}
    .main .block-container{padding-top:.8rem;padding-bottom:.8rem}
    
    .mh{
        background:linear-gradient(135deg,#1e3a5f 0%,#2563eb 100%);
        padding:16px 24px;
        border-radius:12px;
        margin-bottom:12px;
        box-shadow:0 8px 24px rgba(30,58,95,0.15);
        display:flex;
        align-items:center;
        gap:16px;
    }
    .mh h1{color:#fff;font-size:42px;font-weight:800;margin:0;flex:1}
    .mh .logo{height:50px;width:auto;max-width:150px;object-fit:contain;border-radius:6px}
    .mh .db{
        background:rgba(255,255,255,0.2);
        padding:6px 16px;
        border-radius:16px;
        color:#fff;
        font-size:20px;
        font-weight:600;
        border:1px solid rgba(255,255,255,0.3);
        backdrop-filter:blur(10px);
    }
    
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px}
    .cc{
        background:#fff;
        border-radius:12px;
        padding:18px 16px;
        box-shadow:0 4px 12px rgba(0,0,0,0.06);
        border-left:4px solid;
        transition:transform 0.2s,box-shadow 0.2s;
        text-align:center;
    }
    .cc:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(0,0,0,0.1);}
    .cc .cv{font-size:32px;font-weight:900;line-height:1.1;}
    .cc .cl{font-size:16px;color:#1e293b;font-weight:800;text-transform:uppercase;letter-spacing:.5px;margin-top:8px;}
    .cc.c1{border-left-color:#3b82f6}.cc.c1 .cv{color:#2563eb}
    .cc.c2{border-left-color:#10b981}.cc.c2 .cv{color:#059669}
    .cc.c3{border-left-color:#8b5cf6}.cc.c3 .cv{color:#7c3aed}
    .cc.c4{border-left-color:#ef4444}.cc.c4 .cv{color:#dc2626}
    .cc.c5{border-left-color:#3b82f6}.cc.c5 .cv{color:#2563eb}
    .cc.c6{border-left-color:#06b6d4}.cc.c6 .cv{color:#0891b2}
    .cc.c7{border-left-color:#f59e0b}.cc.c7 .cv{color:#d97706}
    .cc.c8{border-left-color:#f97316}.cc.c8 .cv{color:#ea580c}
    
    .stl{font-size:16px;font-weight:800;color:var(--primary);margin:10px 0 5px 0;padding-left:12px;border-left:4px solid var(--info);}
    
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
    .cb-green td{background:#059669!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .cb-red td{background:#dc2626!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    
    .plan-action-table { width:100%; border-collapse:collapse; font-family:Inter,sans-serif; font-size:12px; border:1px solid #cbd5e1; }
    .plan-action-table th { background:#1e3a5f; color:#fff; font-weight:700; padding:8px 6px; border:1px solid #1e3a5f; }
    .plan-action-table td { padding:6px 8px; border:1px solid #cbd5e1; text-align:center; vertical-align:middle;}
    .plan-action-table td:first-child { text-align:left; font-weight:800; }
    
    .stTabs [data-baseweb="tab-list"]{gap:6px;background:#e2e8f0;padding:6px;border-radius:8px;margin-bottom:8px}
    .stTabs [data-baseweb="tab"]{
        border-radius:6px;
        padding:12px 22px;
        font-weight:700;
        font-size:20px;
        line-height:1.5;
        min-height:48px;
    }
    .stTabs [data-baseweb="tab"] span,
    .stTabs [data-baseweb="tab"] > div{
        font-size:22px !important;
    }
    .stTabs [aria-selected="true"]{
        background:#fff!important;
        color:var(--primary)!important;
        box-shadow:0 3px 8px rgba(0,0,0,.1);
        font-size:21px;
    }
    .stTabs [data-baseweb="tab"] svg{width:22px;height:22px}
    
    .ca{background:#fff;border-radius:var(--radius);padding:12px;margin-top:6px;border:1px solid var(--border);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:14px;font-weight:700;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid var(--border)}
    .car{display:flex;align-items:center;margin-bottom:6px;font-size:12px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:260px;font-weight:600;color:var(--primary);text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:26px;background:#edf2f7;border-radius:4px;overflow:visible;position:relative}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    
    .car .target-mark{
        position:absolute;
        top:-4px;
        bottom:-4px;
        width:3px;
        background:var(--info) !important; 
        z-index:20;
        transform:translateX(-50%);
        box-shadow:0 0 6px rgba(59,130,246,.9),0 0 2px rgba(0,0,0,.4);
        border-radius:2px;
    }
    .car .cav-out{font-size:12px;font-weight:800;color:#1e293b;min-width:55px;text-align:right;padding-left:6px}
    .car .cav-tgt{font-size:10px;font-weight:700;color:#1e293b;min-width:42px;text-align:right;padding-left:4px;opacity:.7}
    .gbr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f1f5f9}
    .gbr:last-child{border:none}
    .gbr-l{width:160px;font-weight:600;color:#1e293b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;padding-right:10px;}
    .gbr-g{display:flex;align-items:center;gap:4px;flex:1;position:relative}
    .gbr-target{position:absolute;left:90%;top:-4px;bottom:-4px;width:3px;background:var(--primary) !important;z-index:10;box-shadow:0 0 6px rgba(30,58,95,.8);border-radius:2px}
    .gbr-target-label{position:absolute;left:90%;top:-20px;transform:translateX(-50%);font-size:9px;font-weight:800;color:#fff;background:var(--primary) !important;padding:1px 5px;border-radius:3px;white-space:nowrap;z-index:11;box-shadow:0 1px 3px rgba(0,0,0,.2)}
    .gbr-w{flex:1;height:22px;background:#f1f5f9;border-radius:3px;overflow:hidden}
    .gbr-f{height:100%;border-radius:3px}
    .gbr-v{font-size:11px;font-weight:800;min-width:48px;text-align:right;color:#1e293b}
    .gbr-legend{display:flex;gap:14px;margin-bottom:10px;font-size:12px;font-weight:700;align-items:center}
    .gbr-legend span{display:flex;align-items:center;gap:5px}
    .gbr-legend i{display:inline-block;width:14px;height:14px;border-radius:2px}
    .gbr-legend .target-icon{display:inline-block;width:3px;height:14px;background:var(--primary) !important;border-radius:1px;box-shadow:0 0 3px rgba(30,58,95,.6)}
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
    
    div[data-testid="stSidebar"]{
        background:linear-gradient(180deg,#1e40af 0%,#1e3a8a 50%,#1e3a5f 100%)!important;
    }
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label,div[data-testid="stSidebar"] .stTextInput label{color:rgba(255,255,255,.9)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.1);border-radius:6px;padding:5px 10px;margin-bottom:5px;border:1px solid rgba(255,255,255,.15)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div,div[data-testid="stSidebar"] .stTextInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    
    .es{text-align:center;padding:14px;color:#64748b;font-size:14px}
    .synth-tbl{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px}
    .synth-tbl thead th{background:var(--primary);color:#fff;font-weight:700;font-size:11px;padding:5px 8px;border:none;white-space:nowrap;position:sticky;top:0}
    .synth-tbl tbody td{padding:4px 8px;border-bottom:1px solid var(--border);text-align:center;color:#1e293b !important}
    .synth-tbl tbody tr:nth-child(even) td{background:#f8fafc}
    .synth-tbl tbody tr:hover td{background:#eff6ff!important}
    .synth-tbl .poste-cell{text-align:left;font-weight:700;white-space:nowrap;min-width:140px;color:#1e293b !important}
    
    [data-testid="stHeaderActionElements"]{display:none !important;}
    [data-testid="stActionButtonContainer"]{display:none !important;}
    
    .footer {
        text-align: center;
        margin-top: 30px;
        padding: 15px;
        color: #64748b;
        font-size: 13px;
        border-top: 1px solid var(--border);
        font-weight: 600;
    }

    div[data-testid="stDataEditor"] table, 
    div[data-testid="stDataEditor"] th, 
    div[data-testid="stDataEditor"] td {
        font-size: 18px !important;
        line-height: 1.4 !important;
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    div[data-testid="stDataEditor"] [data-testid="stMarkdownContainer"] {
        font-size: 18px !important;
    }
    div[data-testid="stDataEditor"] [data-testid="stTable"] {
        overflow-x: hidden !important;
        width: 100% !important;
    }

    .ano-detail-box{background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:12px}
    .ano-detail-box .ano-ot-list{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
    .ano-detail-box .ano-ot-tag{background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:4px;font-weight:600;font-size:11px;border:1px solid #fca5a5}

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
        div[data-testid="stDataEditor"] table, div[data-testid="stDataEditor"] th, div[data-testid="stDataEditor"] td { font-size: 14px !important; }
    }
    
    @media print {
        section[data-testid="stSidebar"], 
        header[data-testid="stHeader"], 
        div[data-testid="stToolbar"], 
        div[data-testid="stHeaderActionElements"],
        footer, 
        .stDeployButton, 
        #MainMenu { display: none !important; }
        .main .block-container { 
            padding-top: 0 !important; 
            padding-left: 0 !important; 
            padding-right: 0 !important; 
            max-width: 100% !important; 
        }
        .stButton, .stDownloadButton { display: none !important; }
        * {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }
        .tw, .synth-tbl { page-break-inside: avoid; overflow: visible !important; }
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
        
    def get_text_col(df):
        for c in ["Désignation","Designation","Désignation OT","Texte ordre","Texte","Description","Libellé","Libelle"]:
            if c in df.columns: return c
        for c in df.columns:
            if df[c].dtype=='object' and any(kw in str(c).lower() for kw in ['sign','text','desc','libell']):
                return c
        return None
        
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
        statut_colors = {
            "CRÉÉ": "background:#fef3c7;color:#92400e;font-weight:600;",
            "LANC": "background:#dbeafe;color:#1e40af;font-weight:600;",
            "CLOT": "background:#d1fae5;color:#065f46;font-weight:600;",
            "TCLO": "background:#a7f3d0;color:#064e3b;font-weight:600;",
            "Total": "background:#ede9fe;color:#5b21b6;font-weight:700;"
        }
        h='<table class="tw %s"><thead><tr>'%table_class+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for poste,row in piv_df.iterrows():
            h+='<tr><td style="font-weight:700">%s</td>'%poste
            for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
                h+='<td style="text-align:center;%s">%d</td>'%(statut_colors[c], int(row.get(c,0)))
            h+='<td style="text-align:center;%s">%d</td>'%(statut_colors["Total"], int(row.get("Total",0)))
            h+='</tr>'
        h+='<tr class="tr"><td style="font-weight:800">Total</td>'
        for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
            h+='<td style="text-align:center;font-weight:800;%s">%d</td>'%(statut_colors[c], int(piv_df[c].sum()))
        h+='<td style="text-align:center;font-weight:800;%s">%d</td>'%(statut_colors["Total"], int(piv_df["Total"].sum()))
        h+='</tr></tbody></table>'
        return h
        
    def show_pie_pair(piv_df, title_prefix):
        global_counts = piv_df[["CRÉÉ","LANC","CLOT","TCLO"]].sum()
        global_counts = global_counts[global_counts > 0]
        realised = global_counts.get("CLOT", 0) + global_counts.get("TCLO", 0)
        not_realised = global_counts.sum() - realised
        
        if global_counts.empty:
            st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)
            return
            
        colors = ["#8b5cf6", "#f59e0b", "#10b981", "#3b82f6"]
        fig = make_subplots(rows=1, cols=2, specs=[[{"type":"domain"},{"type":"domain"}]], 
                            subplot_titles=(f"{title_prefix} — Par Statut OT", f"{title_prefix} — Réalisés vs Non Réalisés"))
        
        fig.add_trace(go.Pie(labels=global_counts.index, values=global_counts.values, hole=0.4, 
                             textinfo='percent+label', 
                             texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', 
                             textposition='inside',
                             insidetextorientation='radial',
                             textfont=dict(size=14, color='white', family='Inter, sans-serif'),
                             marker=dict(colors=colors, line=dict(color='#FFFFFF', width=3))), 1, 1)
                             
        pie2_data = pd.Series([realised, not_realised], index=["Réalisés (CLOT+TCLO)", "Non Réalisés"])
        
        fig.add_trace(go.Pie(labels=pie2_data.index, values=pie2_data.values, hole=0.5, 
                             textinfo='percent+label', 
                             texttemplate='%{label}<br>%{percent:.1%}<br>(%{value})', 
                             textposition='inside',
                             insidetextorientation='radial',
                             textfont=dict(size=14, color='white', family='Inter, sans-serif'),
                             marker=dict(colors=["#10b981", "#8b5cf6"], line=dict(color='#FFFFFF', width=3))), 1, 2)
                             
        fig.update_layout(margin=dict(t=80, b=20, l=20, r=20), height=450, 
                          legend=dict(orientation="h", yanchor="bottom", y=-0.12, x=0.5, xanchor="center"))
                          
        st.plotly_chart(fig, use_container_width=True)

    def show_simple_pie(piv_df, title, keep_non_carac=False):
        if not keep_non_carac and "NON CARACTERISE" in piv_df.columns:
            piv_df = piv_df.drop(columns=["NON CARACTERISE"])
            
        counts = piv_df.sum()
        counts = counts[counts > 0]
        
        if counts.empty:
            st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)
            return
            
        color_map = {"CARACTERISE": "#10b981", "NON CARACTERISE": "#f97316"}
        type_palette = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#14b8a6', '#6366f1', '#0ea5e9', '#d946ef', '#a855f7']
        
        colors = []
        palette_idx = 0
        for c in counts.index:
            c_str = str(c)
            if c_str in color_map:
                colors.append(color_map[c_str])
            else:
                colors.append(type_palette[palette_idx % len(type_palette)])
                palette_idx += 1
        
        total_sum = counts.sum()
        pull_list = [0.05 if (v/total_sum)*100 < 10 else 0 for v in counts.values]
        
        fig = go.Figure(
            go.Pie(
                labels=counts.index,
                values=counts.values,
                hole=0.4,
                sort=False,
                textinfo="percent",
                textposition="outside",
                pull=pull_list,
                marker=dict(
                    colors=colors,
                    line=dict(color="white", width=2)
                )
            )
        )

        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>Nombre : %{value}<br>Pourcentage : %{percent}<extra></extra>",
            textfont=dict(size=13, family='Inter, sans-serif')
        )

        fig.update_layout(
            title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)), 
            height=500, 
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"),
            margin=dict(t=80, b=80, l=40, r=40)
        )
                          
        st.plotly_chart(fig, use_container_width=True)

    # ============================================================
    # === MODIFIE : calc_kpis stocke aussi numerateurs/denominateurs ===
    # === et les listes d'OT/Avis anormaux pour le Plan d'action ===
    # ============================================================
    def calc_kpis(df_i, av_i, now_ts, posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        res['dfp']=df
        
        # Dictionnaires pour stocker les details d'anomalies
        # {kpi_name: {poste: {"num": X, "den": Y, "anomaly_items": [ordres/avis]}}}
        raw_details = {}
        for kpi in ALL_KPI:
            raw_details[kpi] = {}
            for p in posts:
                raw_details[kpi][p] = {"num": 0, "den": 0, "anomaly_items": []}
        
        # --- TAUX_REALISATION_CORRECTIF/PT ---
        filt_corr=(df["Nº appel pl.entret."].fillna(0)==0)&(df["Contient SOPL"]==1)
        an=cpiv(df,filt_corr,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
        an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]
        an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"]))
        
        for p in posts:
            num = int(an.loc[p,"OT_CLOTURES"]) if p in an.index else 0
            den = int(an.loc[p,"TOTAL_OT"]) if p in an.index else 0
            raw_details["TAUX_REALISATION_CORRECTIF/PT"][p] = {"num": num, "den": den, "anomaly_items": []}
            if den > 0 and num < den:
                anom_ots = df[filt_corr & (df["Poste travail princ."]==p) & (~df["Statut OT"].isin(["CLOT","TCLO"]))]["Ordre"].tolist()
                raw_details["TAUX_REALISATION_CORRECTIF/PT"][p]["anomaly_items"] = anom_ots
        
        # --- OT préparation ---
        pr=cpiv(df,(df["Statut OT"]=="CRÉÉ")&(df["Statut utilisateur"].str.contains("CRPR",na=False)),"ap",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pr[c]=pr.get(c,0)
        pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"]); pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0); pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)
        
        prep_base = df[(df["Statut OT"]=="CRÉÉ")&(df["Statut utilisateur"].str.contains("CRPR",na=False))]
        for p in posts:
            total_p = int(pr.loc[p,"Total"]) if p in pr.index else 0
            lt1 = int(pr.loc[p,"<1 mois"]) if p in pr.index else 0
            gt3 = int(pr.loc[p,">3 mois"]) if p in pr.index else 0
            bt13 = int(pr.loc[p,"1 mois < <3 mois"]) if p in pr.index else 0
            raw_details["OT préparation <1 mois"][p] = {"num": lt1, "den": total_p, "anomaly_items": []}
            raw_details["OT préparation >3 mois"][p] = {"num": gt3, "den": total_p, "anomaly_items": []}
            raw_details["OT préparation 1mois< <3mois"][p] = {"num": bt13, "den": total_p, "anomaly_items": []}
            if total_p > 0:
                if lt1 < total_p:
                    anom_ots = prep_base[(prep_base["Poste travail princ."]==p)&(prep_base["ap"]!="<1 mois")]["Ordre"].tolist()
                    raw_details["OT préparation <1 mois"][p]["anomaly_items"] = anom_ots
                if gt3 > 0:
                    anom_ots = prep_base[(prep_base["Poste travail princ."]==p)&(prep_base["ap"]==">3 mois")]["Ordre"].tolist()
                    raw_details["OT préparation >3 mois"][p]["anomaly_items"] = anom_ots
                if bt13 > 0:
                    anom_ots = prep_base[(prep_base["Poste travail princ."]==p)&(prep_base["ap"]=="1 mois < <3 mois")]["Ordre"].tolist()
                    raw_details["OT préparation 1mois< <3mois"][p]["anomaly_items"] = anom_ots
        
        # --- OT planification ---
        pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)),"alp",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"]); pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0); pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)
        
        plan_base = df[(df["Statut OT"]=="LANC")&(df["Statut utilisateur"].str.contains("ATPL",case=False,na=False))]
        for p in posts:
            total_p = int(pl.loc[p,"Total"]) if p in pl.index else 0
            lt1 = int(pl.loc[p,"<1 mois"]) if p in pl.index else 0
            gt3 = int(pl.loc[p,">3 mois"]) if p in pl.index else 0
            bt13 = int(pl.loc[p,"1 mois < <3 mois"]) if p in pl.index else 0
            raw_details["OT planification <1 mois"][p] = {"num": lt1, "den": total_p, "anomaly_items": []}
            raw_details["OT planification >3 mois"][p] = {"num": gt3, "den": total_p, "anomaly_items": []}
            raw_details["OT planification 1mois< <3mois"][p] = {"num": bt13, "den": total_p, "anomaly_items": []}
            if total_p > 0:
                if lt1 < total_p:
                    anom_ots = plan_base[(plan_base["Poste travail princ."]==p)&(plan_base["alp"]!="<1 mois")]["Ordre"].tolist()
                    raw_details["OT planification <1 mois"][p]["anomaly_items"] = anom_ots
                if gt3 > 0:
                    anom_ots = plan_base[(plan_base["Poste travail princ."]==p)&(plan_base["alp"]==">3 mois")]["Ordre"].tolist()
                    raw_details["OT planification >3 mois"][p]["anomaly_items"] = anom_ots
                if bt13 > 0:
                    anom_ots = plan_base[(plan_base["Poste travail princ."]==p)&(plan_base["alp"]=="1 mois < <3 mois")]["Ordre"].tolist()
                    raw_details["OT planification 1mois< <3mois"][p]["anomaly_items"] = anom_ots
        
        # --- OT exécution ---
        ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"]); ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0); ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)
        
        ex_base = df[(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1)]
        for p in posts:
            total_p = int(ex.loc[p,"Total"]) if p in ex.index else 0
            lt1 = int(ex.loc[p,"<1 mois"]) if p in ex.index else 0
            gt3 = int(ex.loc[p,">3 mois"]) if p in ex.index else 0
            bt13 = int(ex.loc[p,"1 mois < <3 mois"]) if p in ex.index else 0
            raw_details["OT exécution <1 mois"][p] = {"num": lt1, "den": total_p, "anomaly_items": []}
            raw_details["OT exécution >3 mois"][p] = {"num": gt3, "den": total_p, "anomaly_items": []}
            raw_details["OT exécution 1mois< <3mois"][p] = {"num": bt13, "den": total_p, "anomaly_items": []}
            if total_p > 0:
                if lt1 < total_p:
                    anom_ots = ex_base[(ex_base["Poste travail princ."]==p)&(ex_base["aex"]!="<1 mois")]["Ordre"].tolist()
                    raw_details["OT exécution <1 mois"][p]["anomaly_items"] = anom_ots
                if gt3 > 0:
                    anom_ots = ex_base[(ex_base["Poste travail princ."]==p)&(ex_base["aex"]==">3 mois")]["Ordre"].tolist()
                    raw_details["OT exécution >3 mois"][p]["anomaly_items"] = anom_ots
                if bt13 > 0:
                    anom_ots = ex_base[(ex_base["Poste travail princ."]==p)&(ex_base["aex"]=="1 mois < <3 mois")]["Ordre"].tolist()
                    raw_details["OT exécution 1mois< <3mois"][p]["anomaly_items"] = anom_ots
        
        # --- OT LANC ESTIME ---
        la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: la[c]=la.get(c,0)
        la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])
        
        lanc_base = df[df["Statut OT"]=="LANC"]
        for p in posts:
            oui = int(la.loc[p,"OUI"]) if p in la.index else 0
            non = int(la.loc[p,"NON"]) if p in la.index else 0
            total_p = oui + non
            raw_details["OT LANC ESTIME"][p] = {"num": oui, "den": total_p, "anomaly_items": []}
            if non > 0:
                anom_ots = lanc_base[(lanc_base["Poste travail princ."]==p)&(lanc_base["OT LANC ESTIME"]=="NON")]["Ordre"].tolist()
                raw_details["OT LANC ESTIME"][p]["anomaly_items"] = anom_ots
        
        # --- Backlog préparation caractérisé ---
        pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
        pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])
        
        crea_base = df[df["Statut OT"]=="CRÉÉ"]
        for p in posts:
            carac = int(pc.loc[p,"CARACTERISE"]) if p in pc.index else 0
            ncarac = int(pc.loc[p,"NON CARACTERISE"]) if p in pc.index else 0
            total_p = carac + ncarac
            raw_details["Backlog préparation caractérisé"][p] = {"num": carac, "den": total_p, "anomaly_items": []}
            if ncarac > 0:
                anom_ots = crea_base[(crea_base["Poste travail princ."]==p)&(crea_base["Backlog preparation"]=="NON CARACTERISE")]["Ordre"].tolist()
                raw_details["Backlog préparation caractérisé"][p]["anomaly_items"] = anom_ots
        
        # --- Backlog planification caractérisé ---
        plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])
        
        for p in posts:
            carac = int(plc.loc[p,"CARACTERISE"]) if p in plc.index else 0
            ncarac = int(plc.loc[p,"NON CARACTERISE"]) if p in plc.index else 0
            total_p = carac + ncarac
            raw_details["Backlog planification caractérisé"][p] = {"num": carac, "den": total_p, "anomaly_items": []}
            if ncarac > 0:
                anom_ots = lanc_base[(lanc_base["Poste travail princ."]==p)&(lanc_base["Backlog planification"]=="NON CARACTERISE")]["Ordre"].tolist()
                raw_details["Backlog planification caractérisé"][p]["anomaly_items"] = anom_ots
        
        # --- OT CONFIME ---
        for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv=pd.pivot_table(df,index="Poste travail princ.",columns=cn,values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            for c in ["OUI","NON"]: pv[c]=pv.get(c,0)
            pv["Total"]=pv["OUI"]+pv["NON"]; pv[cn]=ckpi(pv["OUI"],pv["Total"]); res[kn.lower().replace(" ","_")]=pv
            
            for p in posts:
                oui = int(pv.loc[p,"OUI"]) if p in pv.index else 0
                non = int(pv.loc[p,"NON"]) if p in pv.index else 0
                total_p = oui + non
                raw_details[kn][p] = {"num": oui, "den": total_p, "anomaly_items": []}
                if non > 0:
                    anom_ots = df[(df["Poste travail princ."]==p)&(df[cn]=="NON")]["Ordre"].tolist()
                    raw_details[kn][p]["anomaly_items"] = anom_ots
        
        # --- appel avis approuvé ---
        avf=av.copy(); res['avf']=avf
        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])
        
        for p in posts:
            apprv = int(tca.loc[p,"APRV"]) if p in tca.index else 0
            total_p = int(tca.loc[p,"Total"]) if p in tca.index else 0
            raw_details["appel avis approuvé"][p] = {"num": apprv, "den": total_p, "anomaly_items": []}
            if total_p > 0 and apprv < total_p:
                anom_avs = avf[(avf["Poste travail princ."]==p)&(avf["Statut utilisateur"]!="APRV")]["Avis"].tolist()
                raw_details["appel avis approuvé"][p]["anomaly_items"] = anom_avs
        
        # --- Performance Graissage ---
        g_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_den=df[(df["Contient SOPL"]==1)&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_df=pd.DataFrame({"_n":g_num,"_d":g_den}).reindex(posts,fill_value=0)
        g_df["Performance Graissage"]=np.where(g_df["_d"]==0,100.0,(g_df["_n"]/g_df["_d"])*100)
        
        for p in posts:
            num = int(g_df.loc[p,"_n"]) if p in g_df.index else 0
            den = int(g_df.loc[p,"_d"]) if p in g_df.index else 0
            raw_details["Performance Graissage"][p] = {"num": num, "den": den, "anomaly_items": []}
            if den > 0 and num < den:
                anom_ots = df[(df["Contient SOPL"]==1)&(df["_tw_num"]==350)&(df["Poste travail princ."]==p)&(~df["Statut OT"].isin(["CLOT","TCLO"]))]["Ordre"].tolist()
                raw_details["Performance Graissage"][p]["anomaly_items"] = anom_ots
        
        # --- Performance Inspection ---
        ins_types=[290,300,310]
        ins_base=(df["_tw_num"].isin(ins_types))&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
        ins_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_den=df[(df["Contient SOPL"]==1)&ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_df=pd.DataFrame({"_n":ins_num,"_d":ins_den}).reindex(posts,fill_value=0)
        ins_df["Performance Inspection"]=np.where(ins_df["_d"]==0,100.0,(ins_df["_n"]/ins_df["_d"])*100)
        
        for p in posts:
            num = int(ins_df.loc[p,"_n"]) if p in ins_df.index else 0
            den = int(ins_df.loc[p,"_d"]) if p in ins_df.index else 0
            raw_details["Performance Inspection"][p] = {"num": num, "den": den, "anomaly_items": []}
            if den > 0 and num < den:
                anom_ots = df[(df["Contient SOPL"]==1)&ins_base&(df["Poste travail princ."]==p)&(~df["Statut OT"].isin(["CLOT","TCLO"]))]["Ordre"].tolist()
                raw_details["Performance Inspection"][p]["anomaly_items"] = anom_ots
        
        # --- Performance Appels Systématiques ---
        sys_base=(df["_tw_num"]==360)&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now_ts)
        sys_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_den=df[(df["Contient SOPL"]==1)&sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_df=pd.DataFrame({"_n":sys_num,"_d":sys_den}).reindex(posts,fill_value=0)
        sys_df["Performance Appels Systématiques"]=np.where(sys_df["_d"]==0,100.0,(sys_df["_n"]/sys_df["_d"])*100)
        
        for p in posts:
            num = int(sys_df.loc[p,"_n"]) if p in sys_df.index else 0
            den = int(sys_df.loc[p,"_d"]) if p in sys_df.index else 0
            raw_details["Performance Appels Systématiques"][p] = {"num": num, "den": den, "anomaly_items": []}
            if den > 0 and num < den:
                anom_ots = df[(df["Contient SOPL"]==1)&sys_base&(df["Poste travail princ."]==p)&(~df["Statut OT"].isin(["CLOT","TCLO"]))]["Ordre"].tolist()
                raw_details["Performance Appels Systématiques"][p]["anomaly_items"] = anom_ots
        
        # --- OT Fiabilité & Total Avis de Panne (toujours 100%) ---
        for kpi_fix in ["OT Fiabilité", "Total Avis de Panne"]:
            for p in posts:
                raw_details[kpi_fix][p] = {"num": 1, "den": 1, "anomaly_items": []}
        
        fiab_s=pd.Series(100.0,index=posts); avpan_s=pd.Series(100.0,index=posts)
        res['raw_details']=raw_details
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

    # ============================================================
    # === MODIFIE : coloration vert/rouge uniquement (plus de jaune) ===
    # ============================================================
    def get_bar_color(kpi, val):
        try: v = float(val)
        except: return "#cbd5e0"
        if kpi == "TAUX_REALISATION_CORRECTIF/PT":
            return "#38a169" if v >= 95 else "#e53e3e"
        if kpi in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            return "#38a169" if v >= 80 else "#e53e3e"
        if kpi in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return "#38a169" if v <= 15 else "#e53e3e"
        if kpi in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return "#38a169" if v <= 5 else "#e53e3e"
        if kpi in ["Performance Graissage","Performance Inspection"]:
            return "#38a169" if v >= 95 else "#e53e3e"
        if kpi == "Performance Appels Systématiques":
            return "#38a169" if v >= 85 else "#e53e3e"
        if kpi == "appel avis approuvé":
            return "#38a169" if v >= 95 else "#e53e3e"
        if kpi == "OT_COR_EGAL":
            return "#38a169" if v >= 95 else "#e53e3e"
        if kpi in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT Fiabilité","Total Avis de Panne"]:
            return "#38a169" if v >= 100 else "#e53e3e"
        return "#cbd5e0"

    # ============================================================
    # === NOUVEAU : détermine si une cellule est verte (pour Total Général) ===
    # ============================================================
    def is_cell_green(kpi, val):
        try: v = float(val)
        except: return False
        if kpi == "TAUX_REALISATION_CORRECTIF/PT": return v >= 95
        if kpi in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return v >= 80
        if kpi in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return v <= 15
        if kpi in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return v <= 5
        if kpi in ["Performance Graissage","Performance Inspection"]: return v >= 95
        if kpi == "Performance Appels Systématiques": return v >= 85
        if kpi == "appel avis approuvé": return v >= 95
        if kpi == "OT_COR_EGAL": return v >= 95
        if kpi in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT Fiabilité","Total Avis de Panne"]: return v >= 100
        return False

    # ============================================================
    # === MODIFIE : Total Général = (nb cellules vertes / nb total) * 100 ===
    # === vert si >=90%, rouge sinon ===
    # ============================================================
    def render_kpi_table(ckdf, kpi_list, posts, table_class, show_responsable=False):
        cols = ["Poste de travail"] + kpi_list
        if show_responsable:
            cols.append("Responsable")
        cols.append("Total Général")
        h = '<table class="tw %s"><thead><tr>' % table_class + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for poste in posts:
            if poste not in ckdf.index: continue
            row = ckdf.loc[poste]
            h += '<tr><td style="font-weight:700">%s</td>' % poste
            green_count = 0
            total_count = len(kpi_list)
            for kpi in kpi_list:
                val = row.get(kpi, 0)
                try: v = float(val)
                except: v = 0
                color = get_bar_color(kpi, v)
                bg = "background:%s;color:#fff;font-weight:700;" % color if color not in ["#cbd5e0"] else "color:#64748b;font-weight:600;"
                h += '<td style="text-align:center;%s">%.1f%%</td>' % (bg, v)
                if is_cell_green(kpi, v):
                    green_count += 1
            # Total Général corrigé
            tg = (green_count / total_count * 100) if total_count > 0 else 100
            tg_color = "#059669" if tg >= 90 else "#dc2626"
            tg_bg = "background:%s;color:#fff;font-weight:800;" % tg_color
            h += '<td style="text-align:center;%s">%.1f%%</td>' % (tg_bg, tg)
            if show_responsable:
                first_kpi = kpi_list[0] if kpi_list else ""
                resp = KPI_RESP_MAP.get(first_kpi, "")
                h += '<td style="text-align:center;font-weight:600;color:#1e293b">%s</td>' % resp
            h += '</tr>'
        # Ligne moyenne / Total Général global
        h += '<tr class="cb"><td style="font-weight:800">MOYENNE</td>'
        green_count_global = 0
        total_cells = 0
        for kpi in kpi_list:
            vals = []
            for poste in posts:
                if poste in ckdf.index:
                    try: vals.append(float(ckdf.loc[poste, kpi]))
                    except: pass
            avg = sum(vals) / len(vals) if vals else 0
            h += '<td style="text-align:center">%.1f%%</td>' % avg
            if is_cell_green(kpi, avg):
                green_count_global += 1
            total_cells += 1
        tg_global = (green_count_global / total_cells * 100) if total_cells > 0 else 100
        tg_global_color = "#059669" if tg_global >= 90 else "#dc2626"
        h += '<td style="text-align:center;background:%s;color:#fff;font-weight:800">%.1f%%</td>' % (tg_global_color, tg_global)
        if show_responsable:
            h += '<td style="text-align:center">—</td>'
        h += '</tr></tbody></table>'
        return h

    # ============================================================
    # === MODIFIE : nombre d'anomalies = dénominateur - numérateur ===
    # ============================================================
    def detect_anomalies(ckdf, kpi_list, posts, raw_details):
        ano_rows = []
        for poste in posts:
            if poste not in ckdf.index: continue
            row = ckdf.loc[poste]
            for kpi in kpi_list:
                val = row.get(kpi, 0)
                try: v = float(val)
                except: v = 0
                cible = CIBLE.get(kpi, 100)
                is_anomaly = False
                if kpi in LOWER_BETTER:
                    if v > cible: is_anomaly = True
                else:
                    if v < cible: is_anomaly = True
                if is_anomaly:
                    ecart = v - cible
                    # Nombre d'anomalies précis = denom - num
                    rd = raw_details.get(kpi, {}).get(poste, {})
                    nb_anomalies = rd.get("den", 0) - rd.get("num", 0)
                    if nb_anomalies < 0: nb_anomalies = 0
                    ano_rows.append({
                        "Poste de travail": poste,
                        "KPI": kpi,
                        "Valeur": round(v, 1),
                        "Cible": cible,
                        "Ecart": round(ecart, 1),
                        "Nb Anomalies": int(nb_anomalies),
                        "Action": ACT_MAP.get(kpi, ""),
                        "Responsable": KPI_RESP_MAP.get(kpi, "")
                    })
        return ano_rows

    def render_anomalies_table(ano_rows, table_class):
        if not ano_rows: return '<div class="es">Aucune anomalie detectee</div>'
        cols = ["Poste de travail", "KPI", "Valeur", "Cible", "Ecart", "Nb Anomalies", "Action", "Responsable"]
        h = '<table class="tw %s"><thead><tr>' % table_class + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for r in ano_rows:
            ecart_color = "#e53e3e" if r["Ecart"] < 0 else "#f59e0b"
            h += '<tr>'
            h += '<td style="font-weight:700">%s</td>' % r["Poste de travail"]
            h += '<td style="font-weight:600">%s</td>' % r["KPI"]
            h += '<td style="text-align:center;font-weight:700">%.1f%%</td>' % r["Valeur"]
            h += '<td style="text-align:center">%.0f%%</td>' % r["Cible"]
            h += '<td style="text-align:center;font-weight:800;color:%s">%+.1f</td>' % (ecart_color, r["Ecart"])
            h += '<td style="text-align:center;font-weight:800;color:#e53e3e">%d</td>' % r["Nb Anomalies"]
            h += '<td style="font-size:11px;max-width:280px">%s</td>' % r["Action"]
            h += '<td style="text-align:center;font-weight:600">%s</td>' % r["Responsable"]
            h += '</tr>'
        h += '</tbody></table>'
        return h

    def render_bar_charts(ckdf, kpi_list, posts, section_title):
        h = '<div class="ca"><div class="ct">%s</div>' % section_title
        for kpi in kpi_list:
            vals = []
            for p in posts:
                if p in ckdf.index:
                    try: vals.append(float(ckdf.loc[p, kpi]))
                    except: pass
            if not vals: continue
            avg = sum(vals)/len(vals)
            cible = CIBLE.get(kpi, 100)
            color = get_bar_color(kpi, avg)
            pct = min(avg, 100)
            h += '<div class="car">'
            h += '<div class="cal" title="%s">%s</div>' % (kpi, kpi)
            h += '<div class="cab">'
            h += '<div class="caf" style="width:%.1f%%;background:%s"></div>' % (pct, color)
            h += '<div class="target-mark" style="left:%.0f%%"></div>' % cible
            h += '</div>'
            h += '<div class="cav-out">%.1f%%</div>' % avg
            h += '<div class="cav-tgt">(%.0f%%)</div>' % cible
            h += '</div>'
        h += '</div>'
        return h

    # ============================================================
    # SIDEBAR
    # ============================================================
    with st.sidebar:
        st.markdown('<div style="text-align:center;padding:10px 0"><h2 style="color:#fff;font-size:18px;margin:0">📊 NAVIGATEUR</h2></div>', unsafe_allow_html=True)
        
        ot_file = st.file_uploader("Fichier OT (ot.xlsx)", type=["xlsx"], key="ot_up")
        av_file = st.file_uploader("Fichier Avis (avis.xlsx)", type=["xlsx"], key="av_up")
        
        st.markdown("---")
        sel_posts = st.multiselect("Postes de travail", [], format_func=lambda x: x)
        sel_search = st.text_input("Recherche poste", "")
        
        st.markdown("---")
        show_cibles = st.checkbox("Afficher cibles", value=True)
        show_bars = st.checkbox("Afficher barres", value=True)
        
        hist_file = st.file_uploader("Historique KPIs (optionnel)", type=["xlsx"], key="hist_up")

    # ============================================================
    # CHARGEMENT DES DONNEES
    # ============================================================
    if not ot_file or not av_file:
        st.markdown('<div style="min-height:60vh;display:flex;align-items:center;justify-content:center"><div style="text-align:center"><div style="font-size:80px;margin-bottom:20px">📂</div><h2 style="color:#1e3a5f;font-weight:800">Chargement des donnees</h2><p style="color:#64748b;font-size:16px;margin-top:8px">Veuillez charger les fichiers <b>ot.xlsx</b> et <b>avis.xlsx</b> depuis la barre laterale.</p></div></div>', unsafe_allow_html=True)
        st.stop()

    with st.spinner("Traitement des donnees..."):
        df, avf, all_posts, now_ts = prepare_data(ot_file.read(), av_file.read(), fichier_date)
    
    if not all_posts:
        st.markdown('<div class="es">Aucun poste de travail trouve dans les donnees.</div>', unsafe_allow_html=True)
        st.stop()

    posts = sel_posts if sel_posts else all_posts
    if sel_search:
        posts = [p for p in posts if sel_search.lower() in p.lower()]

    with st.spinner("Calcul des KPIs..."):
        res = calc_kpis(df, avf, now_ts, posts)
    
    ckdf = res['ckdf']
    raw_details = res.get('raw_details', {})

    # ============================================================
    # HEADER
    # ============================================================
    logo_b64 = get_logo_base64()
    logo_html = '<img class="logo" src="data:image/png;base64,%s" alt="Logo">' % logo_b64 if logo_b64 else ''
    st.markdown('<div class="mh">%s<h1>Dashboard KPI Maintenance</h1><div class="db">📅 %s</div></div>' % (logo_html, fichier_date), unsafe_allow_html=True)

    # ============================================================
    # CARTES RESUME
    # ============================================================
    perf_vals = []
    for p in posts:
        if p in ckdf.index:
            for kpi in QK:
                try: perf_vals.append(float(ckdf.loc[p, kpi]))
                except: pass
    qual_vals = []
    for p in posts:
        if p in ckdf.index:
            for kpi in PK:
                try: qual_vals.append(float(ckdf.loc[p, kpi]))
                except: pass
    
    avg_perf = sum(perf_vals)/len(perf_vals) if perf_vals else 0
    avg_qual = sum(qual_vals)/len(qual_vals) if qual_vals else 0
    
    # Score Performance basé sur cellules vertes
    green_perf = sum(1 for p in posts if p in ckdf.index for kpi in QK if is_cell_green(kpi, float(ckdf.loc[p, kpi])))
    total_perf_cells = len(posts) * len(QK)
    score_perf = (green_perf / total_perf_cells * 100) if total_perf_cells > 0 else 100
    
    green_qual = sum(1 for p in posts if p in ckdf.index for kpi in PK if is_cell_green(kpi, float(ckdf.loc[p, kpi])))
    total_qual_cells = len(posts) * len(PK)
    score_qual = (green_qual / total_qual_cells * 100) if total_qual_cells > 0 else 100

    total_ot = len(res['dfp'])
    total_av = len(avf)

    st.markdown('<div class="cr">'
        '<div class="cc c1"><div class="cv">%d</div><div class="cl">Total OT</div></div>'
        '<div class="cc c2"><div class="cv">%d</div><div class="cl">Total Avis</div></div>'
        '<div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div>'
        '<div class="cc c4"><div class="cv">%.1f%%</div><div class="cl">Score Qualite</div></div>'
        '</div>' % (total_ot, total_av, score_perf, score_qual), unsafe_allow_html=True)

    # ============================================================
    # CALCUL ANOMALIES
    # ============================================================
    ano_perf = detect_anomalies(ckdf, QK, posts, raw_details)
    ano_qual = detect_anomalies(ckdf, PK, posts, raw_details)
    total_ano = len(ano_perf) + len(ano_qual)
    total_ano_count = sum(r["Nb Anomalies"] for r in ano_perf + ano_qual)

    # ============================================================
    # ONGLETS
    # ============================================================
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["⚡ Performance", "✅ Qualité", "📋 Synthèse", "📰 Journal", "🎯 Plan d'action", "⚠️ Anomalies"])

    # --- ONGLET PERFORMANCE ---
    with tab1:
        st.markdown('<div class="stl">Indicateurs de Performance</div>', unsafe_allow_html=True)
        st.markdown(render_kpi_table(ckdf, QK, posts, "pt"), unsafe_allow_html=True)
        
        if show_bars:
            st.markdown(render_bar_charts(ckdf, QK, posts, "Barres de progression — Performance"), unsafe_allow_html=True)
        
        with st.expander("📊 Repartition par statut OT (correctifs)", expanded=False):
            corr_df = res['dfp'][(res['dfp']["Nº appel pl.entret."].fillna(0)==0)&(res['dfp']["Contient SOPL"]==1)]
            piv_statut = build_statut_pivot(corr_df, posts)
            st.markdown(html_statut_pivot(piv_statut, "pt"), unsafe_allow_html=True)
            show_pie_pair(piv_statut, "Correctifs")
        
        with st.expander("📊 Repartition par age — Preparation", expanded=False):
            pr_piv = cpiv(res['dfp'],(res['dfp']["Statut OT"]=="CRÉÉ")&(res['dfp']["Statut utilisateur"].str.contains("CRPR",na=False)),"ap",posts)
            show_simple_pie(pr_piv, "Age Preparation — Repartition globale")

    # --- ONGLET QUALITE ---
    with tab2:
        st.markdown('<div class="stl">Indicateurs de Qualite</div>', unsafe_allow_html=True)
        st.markdown(render_kpi_table(ckdf, PK, posts, "qt"), unsafe_allow_html=True)
        
        if show_bars:
            st.markdown(render_bar_charts(ckdf, PK, posts, "Barres de progression — Qualite"), unsafe_allow_html=True)
        
        with st.expander("📊 Backlog Preparation", expanded=False):
            bp_piv = pd.pivot_table(res['dfp'][res['dfp']["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            show_simple_pie(bp_piv, "Backlog Preparation — Caracterisation")
        
        with st.expander("📊 Backlog Planification", expanded=False):
            bl_piv = pd.pivot_table(res['dfp'][res['dfp']["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            show_simple_pie(bl_piv, "Backlog Planification — Caracterisation")

    # --- ONGLET SYNTHESE ---
    with tab3:
        st.markdown('<div class="stl">Synthese Generale</div>', unsafe_allow_html=True)
        
        synth_kpis = QK + PK
        synth_html = '<table class="synth-tbl"><thead><tr><th>Poste de travail</th>'
        for kpi in synth_kpis:
            synth_html += '<th>%s</th>' % kpi
        synth_html += '<th>Score Perf.</th><th>Score Qual.</th><th>Total General</th></tr></thead><tbody>'
        
        for poste in posts:
            if poste not in ckdf.index: continue
            row = ckdf.loc[poste]
            synth_html += '<tr><td class="poste-cell">%s</td>' % poste
            
            green_p = sum(1 for kpi in QK if is_cell_green(kpi, float(row.get(kpi, 0))))
            green_q = sum(1 for kpi in PK if is_cell_green(kpi, float(row.get(kpi, 0))))
            green_all = green_p + green_q
            total_all = len(QK) + len(PK)
            
            for kpi in synth_kpis:
                val = row.get(kpi, 0)
                try: v = float(val)
                except: v = 0
                color = get_bar_color(kpi, v)
                if color in ["#38a169"]:
                    synth_html += '<td style="background:#d1fae5;color:#065f46;font-weight:700">%.1f%%</td>' % v
                elif color in ["#e53e3e"]:
                    synth_html += '<td style="background:#fee2e2;color:#991b1b;font-weight:700">%.1f%%</td>' % v
                else:
                    synth_html += '<td style="font-weight:600">%.1f%%</td>' % v
            
            sp = (green_p/len(QK)*100) if len(QK) > 0 else 100
            sq = (green_q/len(PK)*100) if len(PK) > 0 else 100
            stg = (green_all/total_all*100) if total_all > 0 else 100
            
            sp_c = "#059669" if sp >= 90 else "#dc2626"
            sq_c = "#059669" if sq >= 90 else "#dc2626"
            stg_c = "#059669" if stg >= 90 else "#dc2626"
            
            synth_html += '<td style="background:%s;color:#fff;font-weight:800">%.1f%%</td>' % (sp_c, sp)
            synth_html += '<td style="background:%s;color:#fff;font-weight:800">%.1f%%</td>' % (sq_c, sq)
            synth_html += '<td style="background:%s;color:#fff;font-weight:800">%.1f%%</td>' % (stg_c, stg)
            synth_html += '</tr>'
        
        synth_html += '</tbody></table>'
        st.markdown(synth_html, unsafe_allow_html=True)

    # --- ONGLET JOURNAL ---
    with tab4:
        st.markdown('<div class="stl">Journal des Variations</div>', unsafe_allow_html=True)
        if hist_file:
            try:
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(hist_file.read())
                    tmp_path = tmp.name
                hist_df = load_historical_kpis(tmp_path)
                os.unlink(tmp_path)
                var_df = calculate_variations(hist_df)
                journal_df = generate_journal(var_df)
                if journal_df.empty:
                    st.markdown('<div class="es">Aucune variation significative detectee dans l\'historique.</div>', unsafe_allow_html=True)
                else:
                    st.dataframe(journal_df, use_container_width=True, height=500)
                    with st.expander("📊 Classements", expanded=False):
                        top5, bot5 = calculate_rankings(var_df)
                        if not top5.empty:
                            st.markdown("**🟢 Top 5 Ameliorations**")
                            st.dataframe(top5, use_container_width=True)
                        if not bot5.empty:
                            st.markdown("**🔴 Top 5 Degradations**")
                            st.dataframe(bot5, use_container_width=True)
            except Exception as e:
                st.markdown('<div class="es">Erreur de lecture de l\'historique: %s</div>' % str(e), unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Charger un fichier historique KPIs dans la barre laterale pour voir le journal des variations.</div>', unsafe_allow_html=True)

    # --- ONGLET PLAN D'ACTION (cliquable) ---
    with tab5:
        st.markdown('<div class="stl">Plan d\'Action — Detail des Anomalies</div>', unsafe_allow_html=True)
        st.markdown('<p style="color:#64748b;font-size:13px;margin-bottom:12px">Cliquez sur le nombre d\'anomalies pour afficher la liste detaillee des OT/Avis a corriger.</p>', unsafe_allow_html=True)
        
        # Grouper par poste
        all_anomalies = ano_perf + ano_qual
        if not all_anomalies:
            st.markdown('<div class="es">Aucune anomalie — Tous les indicateurs sont conformes.</div>', unsafe_allow_html=True)
        else:
            # Regrouper par poste
            from collections import OrderedDict
            by_poste = OrderedDict()
            for a in all_anomalies:
                p = a["Poste de travail"]
                if p not in by_poste:
                    by_poste[p] = []
                by_poste[p].append(a)
            
            plan_html = '<table class="plan-action-table"><thead><tr>'
            plan_html += '<th>Poste de travail</th><th>KPI</th><th>Valeur</th><th>Cible</th><th>Nb Anomalies</th><th>Action</th><th>Responsable</th>'
            plan_html += '</tr></thead><tbody>'
            
            for poste, ano_list in by_poste.items():
                total_nb = sum(a["Nb Anomalies"] for a in ano_list)
                first_row = True
                for a in ano_list:
                    plan_html += '<tr>'
                    if first_row:
                        plan_html += '<td rowspan="%d" style="font-weight:800;vertical-align:middle">%s</td>' % (len(ano_list), poste)
                        first_row = False
                    plan_html += '<td style="font-weight:600;text-align:left">%s</td>' % a["KPI"]
                    plan_html += '<td>%.1f%%</td>' % a["Valeur"]
                    plan_html += '<td>%.0f%%</td>' % a["Cible"]
                    plan_html += '<td style="font-weight:800;color:#dc2626">%d</td>' % a["Nb Anomalies"]
                    plan_html += '<td style="text-align:left;font-size:11px">%s</td>' % a["Action"]
                    plan_html += '<td>%s</td>' % a["Responsable"]
                    plan_html += '</tr>'
            
            plan_html += '</tbody></table>'
            st.markdown(plan_html, unsafe_allow_html=True)
            
            # Expanders cliquables par poste avec detail OT/Avis
            st.markdown('<div style="margin-top:16px"></div>', unsafe_allow_html=True)
            for poste, ano_list in by_poste.items():
                total_nb = sum(a["Nb Anomalies"] for a in ano_list)
                with st.expander("🔧 **%s** — %d anomalie(s) a corriger" % (poste, total_nb), expanded=False):
                    for a in ano_list:
                        kpi = a["KPI"]
                        nb = a["Nb Anomalies"]
                        if nb == 0:
                            continue
                        rd = raw_details.get(kpi, {}).get(poste, {})
                        items = rd.get("anomaly_items", [])
                        item_type = "Avis" if kpi == "appel avis approuvé" else "OT"
                        
                        st.markdown('**%s** — %d %s(s) a corriger (Valeur: %.1f%% / Cible: %.0f%%)' % (
                            kpi, nb, item_type, a["Valeur"], a["Cible"]
                        ))
                        
                        if items:
                            items_str = ", ".join(str(x) for x in sorted(items))
                            st.markdown('<div class="ano-detail-box"><div style="font-weight:700;color:#991b1b;margin-bottom:4px">Liste des %s :</div><div class="ano-ot-list">%s</div></div>' % (
                                item_type,
                                "".join('<span class="ano-ot-tag">%s</span>' % x for x in sorted(items))
                            ), unsafe_allow_html=True)
                        else:
                            st.markdown('<div style="color:#64748b;font-size:12px;padding-left:10px">Aucun detail disponible.</div>')
                        st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)

    # --- ONGLET ANOMALIES (tableau complet cliquable) ---
    with tab6:
        st.markdown('<div class="stl">Tableau Detaille des Anomalies</div>', unsafe_allow_html=True)
        
        if not all_anomalies:
            st.markdown('<div class="es">Aucune anomalie detectee.</div>', unsafe_allow_html=True)
        else:
            # Tableau HTML complet
            ano_html = '<table class="tw at"><thead><tr>'
            ano_html += '<th>Poste de travail</th><th>KPI</th><th>Valeur</th><th>Cible</th><th>Ecart</th><th>Nb Anomalies</th><th>Action</th><th>Responsable</th>'
            ano_html += '</tr></thead><tbody>'
            for r in all_anomalies:
                ecart_color = "#e53e3e" if r["Ecart"] < 0 else "#f59e0b"
                ano_html += '<tr>'
                ano_html += '<td style="font-weight:700">%s</td>' % r["Poste de travail"]
                ano_html += '<td style="font-weight:600">%s</td>' % r["KPI"]
                ano_html += '<td style="text-align:center;font-weight:700">%.1f%%</td>' % r["Valeur"]
                ano_html += '<td style="text-align:center">%.0f%%</td>' % r["Cible"]
                ano_html += '<td style="text-align:center;font-weight:800;color:%s">%+.1f</td>' % (ecart_color, r["Ecart"])
                ano_html += '<td style="text-align:center;font-weight:800;color:#dc2626">%d</td>' % r["Nb Anomalies"]
                ano_html += '<td style="font-size:11px;max-width:280px">%s</td>' % r["Action"]
                ano_html += '<td style="text-align:center;font-weight:600">%s</td>' % r["Responsable"]
                ano_html += '</tr>'
            ano_html += '</tbody></table>'
            st.markdown(ano_html, unsafe_allow_html=True)
            
            # Expanders detailles par poste et par indicateur
            st.markdown('<div style="margin-top:16px"></div>', unsafe_allow_html=True)
            st.markdown('<p style="color:#64748b;font-size:13px;margin-bottom:8px">Cliquez sur un poste pour voir le detail des OT/Avis anormaux par indicateur.</p>', unsafe_allow_html=True)
            
            for poste, ano_list in by_poste.items():
                with st.expander("⚠️ **%s** — %d indicateur(s) en anomalie" % (poste, len(ano_list)), expanded=False):
                    for a in ano_list:
                        kpi = a["KPI"]
                        nb = a["Nb Anomalies"]
                        rd = raw_details.get(kpi, {}).get(poste, {})
                        items = rd.get("anomaly_items", [])
                        item_type = "Avis" if kpi == "appel avis approuvé" else "OT"
                        
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.markdown('**%s**' % kpi)
                            st.markdown('<span style="color:#64748b;font-size:12px">Valeur: %.1f%% | Cible: %.0f%% | Ecart: %+.1f</span>' % (
                                a["Valeur"], a["Cible"], a["Ecart"]
                            ), unsafe_allow_html=True)
                        with col2:
                            st.markdown('<div style="text-align:right"><span style="background:#fee2e2;color:#dc2626;padding:4px 12px;border-radius:6px;font-weight:800;font-size:16px">%d %s(s)</span></div>' % (nb, item_type), unsafe_allow_html=True)
                        
                        if items:
                            st.markdown('<div class="ano-detail-box"><div style="font-weight:700;color:#991b1b;margin-bottom:6px">Detail — %s a corriger :</div><div class="ano-ot-list">%s</div></div>' % (
                                item_type,
                                "".join('<span class="ano-ot-tag">%s</span>' % x for x in sorted(items))
                            ), unsafe_allow_html=True)
                        else:
                            st.markdown('<div style="color:#64748b;font-size:12px;padding-left:10px">Aucun detail d\'element disponible.</div>')
                        st.markdown('<div style="height:8px;border-bottom:1px solid #e2e8f0"></div>', unsafe_allow_html=True)

    # ============================================================
    # EXPORT EXCEL
    # ============================================================
    def build_excel_rows(ckdf, kpi_list, posts, raw_details):
        rows = []
        for poste in posts:
            if poste not in ckdf.index: continue
            row_dict = {"Poste de travail": poste}
            row_ckdf = ckdf.loc[poste]
            green_count = 0
            for kpi in kpi_list:
                val = row_ckdf.get(kpi, 0)
                row_dict[kpi] = round(float(val), 1)
                if is_cell_green(kpi, float(val)):
                    green_count += 1
            tg = (green_count / len(kpi_list) * 100) if len(kpi_list) > 0 else 100
            row_dict["Total Général"] = round(tg, 1)
            rows.append(row_dict)
        return rows

    def build_ano_excel_rows(ano_rows):
        rows = []
        for r in ano_rows:
            rows.append({
                "Poste de travail": r["Poste de travail"],
                "KPI": r["KPI"],
                "Valeur": r["Valeur"],
                "Cible": r["Cible"],
                "Ecart": r["Ecart"],
                "Nb Anomalies": r["Nb Anomalies"],
                "Action": r["Action"],
                "Responsable": r["Responsable"]
            })
        return rows

    p_rows = build_excel_rows(ckdf, QK, posts, raw_details)
    p_cols = ["Poste de travail"] + QK + ["Total Général"]
    q_rows = build_excel_rows(ckdf, PK, posts, raw_details)
    q_cols = ["Poste de travail"] + PK + ["Total Général"]
    ano_p_rows = build_ano_excel_rows(ano_perf)
    ano_p_cols = ["Poste de travail","KPI","Valeur","Cible","Ecart","Nb Anomalies","Action","Responsable"] if ano_p_rows else []
    ano_q_rows = build_ano_excel_rows(ano_qual)
    ano_q_cols = ["Poste de travail","KPI","Valeur","Cible","Ecart","Nb Anomalies","Action","Responsable"] if ano_q_rows else []

    save_kpis_to_excel(p_rows, p_cols, q_rows, q_cols, ano_p_rows, ano_p_cols, ano_q_rows, ano_q_cols, fichier_date)

    # ============================================================
    # FOOTER
    # ============================================================
    st.markdown('<div class="footer">Dashboard KPI Maintenance — %s — %d postes | %d OT | %d Avis | %d anomalie(s) (%d elements a corriger)</div>' % (
        fichier_date, len(posts), total_ot, total_av, total_ano, total_ano_count
    ), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
