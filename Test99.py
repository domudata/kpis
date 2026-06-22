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
    "Taux d'approbation des Avis": "Chef d'atelier",
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
        
    df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLOT|TCLO",na=False) & df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
    
    df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    df["_tw_num"]=pd.to_numeric(df.get("Type de travail",pd.Series(dtype=float)),errors="coerce")
    
    if "Statut système" in df.columns: df["Statut OT"]=df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
    
    avf = raw_av[
    (
        raw_av["Ordre"].isna() |
        (raw_av["Ordre"].astype(str).str.strip() == "")
    )
    &
    (raw_av["Type d'avis"].isin(["ZU","Z4","ZR","ZP"]))
].copy()
    
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

    # === GENERALIZED PIE CHART FUNCTION ===
    def show_pie(data, title=None, hole=0.4, keep_non_carac=False, color_map=None, legend_y=-0.15, height=500):
        default_palette = ['#3b82f6','#10b981','#f59e0b','#8b5cf6','#06b6d4','#14b8a6','#6366f1','#0ea5e9','#d946ef','#a855f7']
        default_color_map = {
            "CARACTERISE":"#10b981","NON CARACTERISE":"#f97316",
            "Réalisés (CLOT+TCLO)":"#10b981","Non Réalisés":"#8b5cf6",
            "CRÉÉ":"#8b5cf6","LANC":"#f59e0b","CLOT":"#10b981","TCLO":"#3b82f6"
        }
        cm = color_map if color_map is not None else default_color_map

        def _build_series(s):
            s = s.copy()
            if not keep_non_carac and "NON CARACTERISE" in s.index:
                s = s.drop("NON CARACTERISE")
            s = s[s > 0]
            return s

        def _get_colors(s):
            colors = []
            for c in s.index:
                c_str = str(c)
                if c_str in cm:
                    colors.append(cm[c_str])
                else:
                    colors.append(default_palette[len(colors) % len(default_palette)])
            return colors

        def _get_pull(s):
            total_sum = s.sum()
            if total_sum == 0: return [0]*len(s)
            return [0.05 if (v/total_sum)*100 < 10 else 0 for v in s.values]

        def _make_trace(s, colors, pull):
            return go.Pie(
                labels=s.index, values=s.values, hole=hole, sort=False,
                textinfo="percent", textposition="outside", pull=pull,
                marker=dict(colors=colors, line=dict(color="white", width=2)),
                hovertemplate="<b>%%{label}</b><br>Nombre: %%{value}<br>Pourcentage: %%{percent}<extra></extra>",
                textfont=dict(size=13, family='Inter, sans-serif')
            )

        if isinstance(data, list):
            n = len(data)
            specs = [[{"type":"domain"}]*n]
            subplot_titles = [t for _, t in data]
            fig = make_subplots(rows=1, cols=n, specs=specs, subplot_titles=subplot_titles)
            has_data = False
            for i, (series, _) in enumerate(data):
                s = _build_series(series)
                if s.empty: continue
                has_data = True
                fig.add_trace(_make_trace(s, _get_colors(s), _get_pull(s)), 1, i+1)
            if not has_data:
                st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)
                return
            fig.update_layout(
                title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)) if title else None,
                height=height, showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=legend_y, x=0.5, xanchor="center"),
                margin=dict(t=80, b=80, l=40, r=40)
            )
        else:
            s = _build_series(data)
            if s.empty:
                st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)
                return
            fig = go.Figure(_make_trace(s, _get_colors(s), _get_pull(s)))
            fig.update_layout(
                title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)) if title else None,
                height=height, showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=legend_y, x=0.5, xanchor="center"),
                margin=dict(t=80, b=80, l=40, r=40)
            )
        st.plotly_chart(fig, use_container_width=True)

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
        
        plc=pd.pivot_table(df[(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0)],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])
        
        for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv=pd.pivot_table(df[df["Statut OT"].isin(["CLOT","TCLO"])],index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            for c in ["OUI","NON"]: pv[c]=pv.get(c,0)
            pv["Total"]=pv["OUI"]+pv["NON"]; pv[cn]=ckpi(pv["OUI"],pv["Total"]); res[kn.lower().replace(" ","_")]=pv
            
        avf=av.copy(); res['avf']=avf
        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["Taux d'approbation des Avis"] = ckpi(tca["APRV"] ,tca["Total"])
        
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
            "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "Performance Graissage":g_df["Performance Graissage"],"Performance Inspection":ins_df["Performance Inspection"],"Performance Appels Systématiques":sys_df["Performance Appels Systématiques"],
            "Taux d'approbation des Avis":tca["Taux d'approbation des Avis"],"OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"],
            "OT Fiabilité":fiab_s,"Total Avis de Panne":avpan_s
        })
        return res

    def get_bar_color(kpi, val):
        try: v = float(val)
        except: return "#cbd5e0"
        if kpi in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            if v>=80: return "#38a169"
            elif v>=75: return "#f59e0b"
            else: return "#e53e3e"
        if kpi in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return "#38a169" if v<=15 else "#e53e3e"
        if kpi in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return "#38a169" if v<=5 else "#e53e3e"
        if kpi=="TAUX_REALISATION_CORRECTIF/PT":
            if v>=85: return "#38a169"
            elif v>=80: return "#f59e0b"
            else: return "#e53e3e"
        if kpi=="Taux d'approbation des Avis":
            if v>=95: return "#38a169"
            elif v>=90: return "#f59e0b"
            else: return "#e53e3e"
        if kpi in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            if v>=100: return "#38a169"
            elif v>=95: return "#f59e0b"
            else: return "#e53e3e"
        if kpi in ["Performance Graissage","Performance Inspection","Performance Appels Systématiques"]:
            if v>=95: return "#38a169"
            elif v>90: return "#f59e0b"
            else: return "#e53e3e"
        if kpi in ["OT Fiabilité","Total Avis de Panne"]:
            return "#38a169" if v>=100 else "#f59e0b"
        if v>=90: return "#38a169"
        elif v>=80: return "#f59e0b"
        else: return "#e53e3e"

    def ks(v,c):
        try: val=float(v)
        except Exception: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            return "✅" if val>=80 else ("⚠️" if val>=75 else "❌")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return "✅" if val<=15 else "❌"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return "✅" if val<=5 else "❌"
        if c=="TAUX_REALISATION_CORRECTIF/PT":
            return "✅" if val>=85 else ("⚠️" if val>=80 else "❌")
        if c=="Taux d'approbation des Avis":
            return "✅" if val>=95 else ("⚠️" if val>=90 else "❌")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            return "✅" if val>=100 else ("⚠️" if val>=95 else "❌")
        if c in ["Performance Graissage","Performance Inspection","Performance Appels Systématiques"]:
            return "✅" if val>=95 else ("⚠️" if val>90 else "❌")
        if c in ["OT Fiabilité","Total Avis de Panne"]:
            return "✅" if val>=100 else "⚠️"
        if val>=90: return "✅"
        elif val>=80: return "⚠️"
        else: return "❌"

    # === ANOMALY & ACTION PLAN FUNCTIONS ===
    def build_anomalies_df(ckdf):
        anomalies = []
        for poste in ckdf.index:
            for kpi in ALL_KPI:
                if kpi not in ckdf.columns: continue
                try: val_f = float(ckdf.loc[poste, kpi])
                except Exception: continue
                cible = CIBLE.get(kpi, 100)
                if kpi in LOWER_BETTER:
                    is_anomaly = val_f > cible
                    ecart = val_f - cible
                else:
                    is_anomaly = val_f < cible
                    ecart = val_f - cible
                if is_anomaly:
                    anomalies.append({
                        "Poste de travail": poste,
                        "KPI": kpi,
                        "Valeur": round(val_f, 2),
                        "Cible": cible,
                        "Écart": round(ecart, 2),
                        "Type": "Performance" if kpi in QK else "Qualité"
                    })
        if not anomalies:
            return pd.DataFrame(columns=["Poste de travail","KPI","Valeur","Cible","Écart","Type"])
        return pd.DataFrame(anomalies)

    def build_action_plan_df(anom_df):
        if anom_df.empty:
            return pd.DataFrame(columns=["Poste de travail","KPI","Anomalie","Action Corrective","Responsable","Délai","Priorité","Type"])
        actions = []
        for _, row in anom_df.iterrows():
            kpi = row["KPI"]
            ecart = abs(row["Écart"])
            if ecart >= 20:
                priorite = "Critique"; delai = "Immédiat"
            elif ecart >= 10:
                priorite = "Haute"; delai = "Sous 7 jours"
            elif ecart >= 5:
                priorite = "Moyenne"; delai = "Sous 15 jours"
            else:
                priorite = "Basse"; delai = "Sous 30 jours"
            actions.append({
                "Poste de travail": row["Poste de travail"],
                "KPI": kpi,
                "Anomalie": "%s: %s%% vs Cible %s%% (Écart: %s%%)" % (kpi, row["Valeur"], row["Cible"], row["Écart"]),
                "Action Corrective": ACT_MAP.get(kpi, "Analyser et corriger l'écart constaté."),
                "Responsable": KPI_RESP_MAP.get(kpi, "Non défini"),
                "Délai": delai,
                "Priorité": priorite,
                "Type": row["Type"]
            })
        return pd.DataFrame(actions)

    def render_anomalies_table(anom_df):
        if anom_df.empty:
            return '<div class="es">✅ Aucune anomalie détectée — Tous les KPIs sont conformes aux cibles.</div>'
        cols = ["Poste de travail","KPI","Valeur (%)","Cible (%)","Écart (pts)","Type"]
        h = '<table class="tw at"><thead><tr>' + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for _, row in anom_df.iterrows():
            ec = "#dc2626"
            h += '<tr>'
            h += '<td style="font-weight:700">%s</td>' % row["Poste de travail"]
            h += '<td>%s</td>' % row["KPI"]
            h += '<td style="text-align:center;font-weight:700">%s</td>' % row["Valeur"]
            h += '<td style="text-align:center">%s</td>' % row["Cible"]
            h += '<td style="text-align:center;font-weight:700;color:%s">%s</td>' % (ec, row["Écart"])
            h += '<td style="text-align:center">%s</td>' % row["Type"]
            h += '</tr>'
        h += '</tbody></table>'
        return h

    def render_action_plan_table(ap_df):
        if ap_df.empty:
            return '<div class="es">✅ Aucun plan d\'action requis — Tous les KPIs sont conformes.</div>'
        cols = ["Poste de travail","KPI","Anomalie constatée","Action Corrective","Responsable","Délai","Priorité"]
        pc = {
            "Critique": "background:#fee2e2;color:#991b1b;font-weight:700;",
            "Haute": "background:#ffedd5;color:#9a3412;font-weight:700;",
            "Moyenne": "background:#fef9c3;color:#854d0e;font-weight:700;",
            "Basse": "background:#f0fdf4;color:#166534;font-weight:600;"
        }
        h = '<table class="plan-action-table"><thead><tr>' + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for _, row in ap_df.iterrows():
            ps = pc.get(row["Priorité"], "")
            h += '<tr>'
            h += '<td>%s</td>' % row["Poste de travail"]
            h += '<td style="font-weight:600">%s</td>' % row["KPI"]
            h += '<td style="text-align:left;font-size:11px">%s</td>' % row["Anomalie"]
            h += '<td style="text-align:left;font-size:11px">%s</td>' % row["Action Corrective"]
            h += '<td style="font-weight:600">%s</td>' % row["Responsable"]
            h += '<td>%s</td>' % row["Délai"]
            h += '<td style="%s">%s</td>' % (ps, row["Priorité"])
            h += '</tr>'
        h += '</tbody></table>'
        return h

    def build_anomaly_excel_rows(anom_df, type_kpi):
        if anom_df.empty: return [], ["Poste de travail","KPI","Valeur","Cible","Écart"]
        f = anom_df[anom_df["Type"] == type_kpi]
        if f.empty: return [], ["Poste de travail","KPI","Valeur","Cible","Écart"]
        return f[["Poste de travail","KPI","Valeur","Cible","Écart"]].to_dict("records"), ["Poste de travail","KPI","Valeur","Cible","Écart"]

    def render_kpi_bars(ckdf, kpi_list, posts, card_class="ca"):
        html_parts = []
        for kpi in kpi_list:
            if kpi not in ckdf.columns: continue
            cible = CIBLE.get(kpi, 100)
            target_pct = min(cible, 100)
            h = '<div class="%s"><div class="ct">%s <span style="float:right;font-size:12px;color:#64748b;font-weight:600">Cible: %s%%</span></div>' % (card_class, kpi, cible)
            for poste in posts:
                try: val = float(ckdf.loc[poste, kpi])
                except Exception: continue
                color = get_bar_color(kpi, val)
                pct = min(max(val, 0), 100)
                h += '<div class="car">'
                h += '<div class="cal">%s</div>' % poste
                h += '<div class="cab">'
                h += '<div class="caf" style="width:%.1f%%;background:%s"></div>' % (pct, color)
                h += '<div class="target-mark" style="left:%.1f%%"></div>' % target_pct
                h += '</div>'
                h += '<div class="cav-out">%s %s</div>' % (round(val,1), ks(val, kpi))
                h += '<div class="cav-tgt">Cible: %s%%</div>' % cible
                h += '</div>'
            h += '</div>'
            html_parts.append(h)
        return html_parts

    # === SIDEBAR ===
    logo_b64 = get_logo_base64()
    if logo_b64:
        st.sidebar.markdown('<img src="data:image/png;base64,%s" style="width:100%%;border-radius:8px;margin-bottom:10px">' % logo_b64, unsafe_allow_html=True)
    st.sidebar.title("📋 Dashboard KPI")
    st.sidebar.markdown("---")

    ot_file = st.sidebar.file_uploader("📁 Fichier OT (Excel)", type=["xlsx","xls"], key="ot_up")
    av_file = st.sidebar.file_uploader("📁 Fichier Avis (Excel)", type=["xlsx","xls"], key="av_up")

    st.sidebar.markdown("---")
    date_input = st.sidebar.text_input("📅 Date de référence", value=fichier_date, key="date_inp")

    if ot_file and av_file:
        with st.spinner("Chargement et traitement des données..."):
            df, avf, apm, now_ts = prepare_data(ot_file.read(), av_file.read(), date_input)

        if apm:
            sel_posts = st.sidebar.multiselect("🔧 Postes de travail", apm, default=apm, key="posts_sel")
        else:
            sel_posts = []
            st.sidebar.warning("Aucun poste SF1/SF2 détecté.")

        sel_kpi_ano = st.sidebar.multiselect("🔍 Filtrer anomalies par KPI", ALL_KPI, default=ALL_KPI, key="kpi_ano_sel")

        if not sel_posts:
            st.warning("⚠️ Veuillez sélectionner au moins un poste de travail.")
            st.stop()

        # === CALCULATE KPIs ===
        res = calc_kpis(df, avf, now_ts, sel_posts)
        ckdf = res['ckdf']

        # === ANOMALIES & ACTION PLAN ===
        anom_all = build_anomalies_df(ckdf)
        ap_all = build_action_plan_df(anom_all)

        # Apply KPI filter
        if sel_kpi_ano:
            anom_filtered = anom_all[anom_all["KPI"].isin(sel_kpi_ano)] if not anom_all.empty else pd.DataFrame()
            ap_filtered = ap_all[ap_all["KPI"].isin(sel_kpi_ano)] if not ap_all.empty else pd.DataFrame()
        else:
            anom_filtered = pd.DataFrame()
            ap_filtered = pd.DataFrame()

        anom_perf = anom_filtered[anom_filtered["Type"]=="Performance"] if not anom_filtered.empty else pd.DataFrame()
        anom_qual = anom_filtered[anom_filtered["Type"]=="Qualité"] if not anom_filtered.empty else pd.DataFrame()
        ap_perf = ap_filtered[ap_filtered["Type"]=="Performance"] if not ap_filtered.empty else pd.DataFrame()
        ap_qual = ap_filtered[ap_filtered["Type"]=="Qualité"] if not ap_filtered.empty else pd.DataFrame()

        # === HEADER ===
        total_ot = int(ckdf.shape[0])
        header_html = '<div class="mh"><h1>📊 Dashboard KPI Maintenance</h1>'
        header_html += '<div class="db">📅 %s</div>' % date_input
        header_html += '<div class="db">🔧 %d postes</div>' % len(sel_posts)
        header_html += '</div>'
        st.markdown(header_html, unsafe_allow_html=True)

        # === SCORE CALCULATIONS ===
        def calc_score(ckdf_sub, kpi_list):
            if ckdf_sub.empty: return 0.0
            vals = []
            for kpi in kpi_list:
                if kpi not in ckdf_sub.columns: continue
                cible = CIBLE.get(kpi, 100)
                for poste in ckdf_sub.index:
                    try: v = float(ckdf_sub.loc[poste, kpi])
                    except: continue
                    if kpi in LOWER_BETTER:
                        vals.append(min(v / cible * 100, 100) if cible > 0 else 100)
                    else:
                        vals.append(min(v / cible * 100, 100) if cible > 0 else 100)
            return round(sum(vals)/len(vals), 1) if vals else 100.0

        score_p = calc_score(ckdf, QK)
        score_q = calc_score(ckdf, PK)

        # === TABS ===
        tab1, tab2, tab3 = st.tabs(["📈 Performance", "🎯 Qualité", "📊 Synthèse Globale"])

        # =============================================
        # TAB 1 : PERFORMANCE
        # =============================================
        with tab1:
            # Cards
            total_ot_p = int(res.get('dfp', pd.DataFrame()).shape[0])
            clot_p = int(ckdf["TAUX_REALISATION_CORRECTIF/PT"].mean())
            nb_anom_p = len(anom_perf)
            nb_act_p = len(ap_perf)
            cards_html = '<div class="cr">'
            cards_html += '<div class="cc c1"><div class="cv">%d</div><div class="cl">Total OT</div></div>' % total_ot_p
            cards_html += '<div class="cc c2"><div class="cv">%s%%</div><div class="cl">Taux Réalisation</div></div>' % round(clot_p,1)
            cards_html += '<div class="cc c5"><div class="cv">%s%%</div><div class="cl">Score Performance</div></div>' % score_p
            cards_html += '<div class="cc c4"><div class="cv">%d</div><div class="cl">Anomalies</div></div>' % nb_anom_p
            cards_html += '</div>'
            st.markdown(cards_html, unsafe_allow_html=True)

            # KPI Bars
            st.markdown('<p class="stl">📈 Indicateurs de Performance</p>', unsafe_allow_html=True)
            bars = render_kpi_bars(ckdf, QK, sel_posts)
            for b in bars:
                st.markdown(b, unsafe_allow_html=True)

            # Statut OT Table & Pie
            st.markdown('<p class="stl">📋 Statut OT Correctif</p>', unsafe_allow_html=True)
            filt_corr_df = res['dfp'][(res['dfp']["Nº appel pl.entret."].fillna(0)==0)&(res['dfp']["Contient SOPL"]==1)]
            filt_corr_df = filt_corr_df[filt_corr_df["Poste travail princ."].isin(sel_posts)]
            piv_statut = build_statut_pivot(filt_corr_df, sel_posts)
            st.markdown(html_statut_pivot(piv_statut, "pt"), unsafe_allow_html=True)

            global_statut_counts = piv_statut[["CRÉÉ","LANC","CLOT","TCLO"]].sum()
            global_statut_counts = global_statut_counts[global_statut_counts > 0]
            realised = global_statut_counts.get("CLOT",0) + global_statut_counts.get("TCLO",0)
            not_realised = global_statut_counts.sum() - realised
            show_pie([
                (global_statut_counts, "Par Statut OT"),
                (pd.Series([realised, not_realised], index=["Réalisés (CLOT+TCLO)","Non Réalisés"]), "Réalisés vs Non Réalisés")
            ], "Distribution OT Correctif", hole=0.4, height=450)

            # Pie : Backlog Préparation
            pc_piv = pd.pivot_table(
                res['dfp'][(res['dfp']["Statut OT"]=="CRÉÉ")&(res['dfp']["Poste travail princ."].isin(sel_posts))],
                index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0
            )
            show_pie(pc_piv.sum(), "Caractérisation Backlog Préparation", hole=0.4)

            # Pie : Backlog Planification
            plc_piv = pd.pivot_table(
                res['dfp'][(res['dfp']["Statut OT"]=="LANC")&(res['dfp']["Contient SOPL"]==0)&(res['dfp']["Poste travail princ."].isin(sel_posts))],
                index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0
            )
            show_pie(plc_piv.sum(), "Caractérisation Backlog Planification", hole=0.4)

            # Anomalies Table
            st.markdown('<p class="stl">🔴 Tableau des Anomalies — Performance</p>', unsafe_allow_html=True)
            st.markdown(render_anomalies_table(anom_perf), unsafe_allow_html=True)

            # Action Plan Table
            st.markdown('<p class="stl">📋 Plan d\'Action — Performance</p>', unsafe_allow_html=True)
            st.markdown(render_action_plan_table(ap_perf), unsafe_allow_html=True)

        # =============================================
        # TAB 2 : QUALITÉ
        # =============================================
        with tab2:
            nb_anom_q = len(anom_qual)
            nb_act_q = len(ap_qual)
            cards_html_q = '<div class="cr">'
            cards_html_q += '<div class="cc c3"><div class="cv">%s%%</div><div class="cl">Score Qualité</div></div>' % score_q
            cards_html_q += '<div class="cc c6"><div class="cv">%d</div><div class="cl">KPI Qualité</div></div>' % len(PK)
            cards_html_q += '<div class="cc c4"><div class="cv">%d</div><div class="cl">Anomalies</div></div>' % nb_anom_q
            cards_html_q += '<div class="cc c7"><div class="cv">%d</div><div class="cl">Actions</div></div>' % nb_act_q
            cards_html_q += '</div>'
            st.markdown(cards_html_q, unsafe_allow_html=True)

            # KPI Bars
            st.markdown('<p class="stl">🎯 Indicateurs de Qualité</p>', unsafe_allow_html=True)
            bars_q = render_kpi_bars(ckdf, PK, sel_posts)
            for b in bars_q:
                st.markdown(b, unsafe_allow_html=True)

            # Pie : OT LANC ESTIME
            la_piv = pd.pivot_table(
                res['dfp'][(res['dfp']["Statut OT"]=="LANC")&(res['dfp']["Poste travail princ."].isin(sel_posts))],
                index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0
            )
            for c in ["OUI","NON"]: 
                if c not in la_piv.columns: la_piv[c]=0
            show_pie(la_piv.sum(), "OT Lancés Estimés", hole=0.4,
                     color_map={"OUI":"#10b981","NON":"#ef4444"})

            # Pie : OT CONFIME
            conf_piv = pd.pivot_table(
                res['dfp'][(res['dfp']["Statut OT"].isin(["CLOT","TCLO"]))&(res['dfp']["Poste travail princ."].isin(sel_posts))],
                index="Poste travail princ.", columns="OT CONFIME", values="Ordre", aggfunc="count", fill_value=0
            )
            for c in ["OUI","NON"]: 
                if c not in conf_piv.columns: conf_piv[c]=0
            show_pie(conf_piv.sum(), "OT Confirmés", hole=0.4,
                     color_map={"OUI":"#10b981","NON":"#ef4444"})

            # Pie : OT_COR_EGAL
            ce_piv = pd.pivot_table(
                res['dfp'][(res['dfp']["Statut OT"].isin(["CLOT","TCLO"]))&(res['dfp']["Poste travail princ."].isin(sel_posts))],
                index="Poste travail princ.", columns="OT_COR_EGAL", values="Ordre", aggfunc="count", fill_value=0
            )
            for c in ["OUI","NON"]: 
                if c not in ce_piv.columns: ce_piv[c]=0
            show_pie(ce_piv.sum(), "Coûts Réels = Coûts Budgétés", hole=0.4,
                     color_map={"OUI":"#10b981","NON":"#ef4444"})

            # Pie : Taux approbation Avis
            avf_filt = res['avf'][res['avf']["Poste travail princ."].isin(sel_posts)]
            av_statut_piv = pd.pivot_table(
                avf_filt, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0
            )
            for c in ["APRQ","APRV","APRV AVAU","REJT"]: 
                if c not in av_statut_piv.columns: av_statut_piv[c]=0
            av_sum = av_statut_piv[["APRQ","APRV","APRV AVAU","REJT"]].sum()
            av_sum = av_sum[av_sum > 0]
            show_pie(av_sum, "Statut des Avis sans OT", hole=0.4,
                     color_map={"APRV":"#10b981","APRV AVAU":"#3b82f6","APRQ":"#f59e0b","REJT":"#ef4444"})

            # Anomalies Table
            st.markdown('<p class="stl">🔴 Tableau des Anomalies — Qualité</p>', unsafe_allow_html=True)
            st.markdown(render_anomalies_table(anom_qual), unsafe_allow_html=True)

            # Action Plan Table
            st.markdown('<p class="stl">📋 Plan d\'Action — Qualité</p>', unsafe_allow_html=True)
            st.markdown(render_action_plan_table(ap_qual), unsafe_allow_html=True)

        # =============================================
        # TAB 3 : SYNTHÈSE GLOBALE
        # =============================================
        with tab3:
            cards_html_s = '<div class="cr">'
            cards_html_s += '<div class="cc c5"><div class="cv">%s%%</div><div class="cl">Score Performance</div></div>' % score_p
            cards_html_s += '<div class="cc c3"><div class="cv">%s%%</div><div class="cl">Score Qualité</div></div>' % score_q
            cards_html_s += '<div class="cc c6"><div class="cv">%s%%</div><div class="cl">Score Global</div></div>' % round((score_p+score_q)/2, 1)
            cards_html_s += '<div class="cc c4"><div class="cv">%d</div><div class="cl">Total Anomalies</div></div>' % len(anom_filtered)
            cards_html_s += '</div>'
            st.markdown(cards_html_s, unsafe_allow_html=True)

            # Synthesis table
            st.markdown('<p class="stl">📊 Synthèse par Poste</p>', unsafe_allow_html=True)
            synth_cols = ["Poste de travail"] + ALL_KPI + ["Score Performance","Score Qualité"]
            h = '<table class="synth-tbl"><thead><tr>' + ''.join('<th>%s</th>' % c for c in synth_cols) + '</tr></thead><tbody>'
            for poste in sel_posts:
                h += '<tr><td class="poste-cell">%s</td>' % poste
                for kpi in ALL_KPI:
                    if kpi in ckdf.columns:
                        try: val = float(ckdf.loc[poste, kpi])
                        except: val = 0
                        color = get_bar_color(kpi, val)
                        h += '<td style="color:%s;font-weight:700">%s</td>' % (color, round(val,1))
                    else:
                        h += '<td>-</td>'
                # Score Performance
                sp = calc_score(ckdf.loc[[poste]], QK)
                sp_color = "#38a169" if sp >= 90 else ("#f59e0b" if sp >= 80 else "#e53e3e")
                h += '<td style="color:%s;font-weight:800">%s%%</td>' % (sp_color, sp)
                # Score Qualité
                sq = calc_score(ckdf.loc[[poste]], PK)
                sq_color = "#38a169" if sq >= 90 else ("#f59e0b" if sq >= 80 else "#e53e3e")
                h += '<td style="color:%s;font-weight:800">%s%%</td>' % (sq_color, sq)
                h += '</tr>'
            h += '</tbody></table>'
            st.markdown(h, unsafe_allow_html=True)

            # Global pie : Performance vs Qualité
            show_pie(
                pd.Series([score_p, score_q], index=["Score Performance","Score Qualité"]),
                "Répartition des Scores", hole=0.5,
                color_map={"Score Performance":"#10b981","Score Qualité":"#3b82f6"}
            )

            # All anomalies
            st.markdown('<p class="stl">🔴 Toutes les Anomalies</p>', unsafe_allow_html=True)
            st.markdown(render_anomalies_table(anom_filtered), unsafe_allow_html=True)

            # All action plans
            st.markdown('<p class="stl">📋 Plan d\'Action Complet</p>', unsafe_allow_html=True)
            st.markdown(render_action_plan_table(ap_filtered), unsafe_allow_html=True)

        # === SAVE TO EXCEL ===
        pcols = ["Poste de travail"] + QK + ["Score Performance"]
        prows = []
        for poste in sel_posts:
            row = {"Poste de travail": poste}
            for kpi in QK:
                if kpi in ckdf.columns:
                    try: row[kpi] = round(float(ckdf.loc[poste, kpi]), 2)
                    except: row[kpi] = 0
            row["Score Performance"] = calc_score(ckdf.loc[[poste]], QK)
            prows.append(row)

        qcols = ["Poste de travail"] + PK + ["Score Qualité"]
        qrows = []
        for poste in sel_posts:
            row = {"Poste de travail": poste}
            for kpi in PK:
                if kpi in ckdf.columns:
                    try: row[kpi] = round(float(ckdf.loc[poste, kpi]), 2)
                    except: row[kpi] = 0
            row["Score Qualité"] = calc_score(ckdf.loc[[poste]], PK)
            qrows.append(row)

        ano_p_r, ano_p_c = build_anomaly_excel_rows(anom_all, "Performance")
        ano_q_r, ano_q_c = build_anomaly_excel_rows(anom_all, "Qualité")

        save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, date_input)

        # Footer
        st.markdown('<div class="footer">Dashboard KPI Maintenance — Généré le %s — Tous droits réservés</div>' % datetime.now().strftime("%d/%m/%Y %H:%M"), unsafe_allow_html=True)

    else:
        st.markdown("""<div style="min-height:60vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#f8fafc,#e2e8f0);padding:40px;border-radius:16px;margin:20px">
        <div style="font-size:80px;margin-bottom:20px">📂</div>
        <h2 style="text-align:center;font-size:28px;color:#1e3a5f;font-weight:800;margin:0">Chargement des Données</h2>
        <p style="text-align:center;color:#64748b;font-size:18px;margin-top:12px;max-width:500px">Veuillez charger les fichiers Excel <b>OT</b> et <b>Avis</b> via la barre latérale pour afficher le dashboard.</p>
        <div style="margin-top:24px;display:flex;gap:20px">
            <div style="background:#fff;padding:16px 24px;border-radius:10px;box-shadow:0 4px 12px rgba(0,0,0,.08);text-align:center">
                <div style="font-size:32px">📁</div>
                <div style="font-weight:700;color:#1e3a5f;margin-top:6px">Fichier OT</div>
                <div style="font-size:12px;color:#64748b">.xlsx / .xls</div>
            </div>
            <div style="background:#fff;padding:16px 24px;border-radius:10px;box-shadow:0 4px 12px rgba(0,0,0,.08);text-align:center">
                <div style="font-size:32px">📁</div>
                <div style="font-weight:700;color:#1e3a5f;margin-top:6px">Fichier Avis</div>
                <div style="font-size:12px;color:#64748b">.xlsx / .xls</div>
            </div>
        </div></div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
