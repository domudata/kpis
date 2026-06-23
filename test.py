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
    return pd.Timestamp.today().strftime("%d/%m/%Y")

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
def read_excel_safe(bytes_data):
    bio = io.BytesIO(bytes_data)
    header = bytes_data[:8]
    if header[:4] in (b'PK\x03\x04', b'PK\x05\x06'):
        for engine in ['openpyxl', 'calamine']:
            try:
                return pd.read_excel(bio, engine=engine)
            except Exception:
                bio.seek(0)
                continue
    if header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        for engine in ['xlrd', 'calamine']:
            try:
                return pd.read_excel(bio, engine=engine)
            except Exception:
                bio.seek(0)
                continue
    for engine in ['openpyxl', 'xlrd', 'calamine']:
        try:
            bio.seek(0)
            return pd.read_excel(bio, engine=engine)
        except Exception:
            continue
    raise ValueError(
        "Format de fichier non reconnu. Le fichier n'est ni un .xlsx ni un .xls valide.\n"
        "Vérifiez que le fichier n'est pas corrompu ou protégé par mot de passe."
    )


@st.cache_data(show_spinner=False)
def prepare_data(ot_bytes, av_bytes, date_str):
    raw_ot = read_excel_safe(ot_bytes)
    raw_av = read_excel_safe(av_bytes)
    raw_ot = excr(raw_ot)
    raw_av = excr(raw_av)

    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in raw_ot.columns: raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in raw_av.columns: raw_av[c]=pd.to_datetime(raw_av[c],errors="coerce")

    now_ts = pd.Timestamp.today()
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
    def ws_sec(title,cols,rows,sr,force_int=False):
        ws.cell(row=sr,column=1,value=title).font=tf; sr+=1
        for j,c in enumerate(cols,1):
            cl=ws.cell(row=sr,column=j,value=c); cl.font=hf; cl.fill=hfl; cl.alignment=Alignment(horizontal='center'); cl.border=tb
        sr+=1
        for r in rows:
            for j,c in enumerate(cols,1):
                val=r.get(c,"")
                if force_int and isinstance(val,(int,float,np.integer,np.floating)):
                    val=int(val)
                cl=ws.cell(row=sr,column=j,value=val); cl.border=tb; cl.alignment=Alignment(horizontal='center')
            sr+=1
        return sr+1
    rn=ws_sec("INDICATEURS DE PERFORMANCE",pcols,prows,rn)
    if ano_p_c and ano_p_r: rn=ws_sec("ANOMALIES PERFORMANCE",ano_p_c,ano_p_r,rn,force_int=True)
    rn=ws_sec("INDICATEURS DE QUALITE",qcols,qrows,rn)
    if ano_q_c and ano_q_r: rn=ws_sec("ANOMALIES QUALITE",ano_q_c,ano_q_r,rn,force_int=True)
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

    .tw.at tbody td{color:#000!important;font-weight:700}
    .tw.st tbody td{color:#000!important;font-weight:700}

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

    def calc_kpis(df_i, av_i, now_ts, posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        res['dfp']=df
        filt_corr=(df["Nº appel pl.entret."].fillna(0)==0)&(df["Contient SOPL"]==1)
        an=cpiv(df,filt_corr,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
        an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]
        an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"]))

        pr=cpiv(df,(df["Statut OT"]=="CRÉÉ")&(df["Statut utilisateur"].str.contains(r"\bCRPR\b",case=False,na=False)),"ap",posts)
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
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["Taux d'approbation des Avis"]=ckpi(tca["APRV"],tca["Total"])

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

        # ==========================
        # TABLEAU DES ANOMALIES
        # ==========================

        res["anomalies_perf"] = pd.DataFrame(index=posts)

        res["anomalies_perf"]["TAUX_REALISATION_CORRECTIF/PT"] = (an["TOTAL_OT"] - an["OT_CLOTURES"]).astype(int)
        res["anomalies_perf"]["OT préparation <1 mois"] = (pr["Total"] - pr["<1 mois"]).astype(int)
        res["anomalies_perf"]["OT préparation >3 mois"] = pr[">3 mois"].astype(int)
        res["anomalies_perf"]["OT préparation 1mois< <3mois"] = pr["1 mois < <3 mois"].astype(int)
        res["anomalies_perf"]["OT planification <1 mois"] = (pl["Total"] - pl["<1 mois"]).astype(int)
        res["anomalies_perf"]["OT planification >3 mois"] = pl[">3 mois"].astype(int)
        res["anomalies_perf"]["OT planification 1mois< <3mois"] = pl["1 mois < <3 mois"].astype(int)
        res["anomalies_perf"]["OT exécution <1 mois"] = (ex["Total"] - ex["<1 mois"]).astype(int)
        res["anomalies_perf"]["OT exécution >3 mois"] = ex[">3 mois"].astype(int)
        res["anomalies_perf"]["OT exécution 1mois< <3mois"] = ex["1 mois < <3 mois"].astype(int)
        res["anomalies_perf"]["Performance Graissage"] = (g_df["_d"] - g_df["_n"]).astype(int)
        res["anomalies_perf"]["Performance Inspection"] = (ins_df["_d"] - ins_df["_n"]).astype(int)
        res["anomalies_perf"]["Performance Appels Systématiques"] = (sys_df["_d"] - sys_df["_n"]).astype(int)

        res["anomalies_qual"] = pd.DataFrame(index=posts)

        res["anomalies_qual"]["Taux d'approbation des Avis"] = (tca["Total"] - tca["APRV"]).astype(int)
        res["anomalies_qual"]["OT LANC ESTIME"] = la["NON"].astype(int)
        res["anomalies_qual"]["Backlog préparation caractérisé"] = pc["NON CARACTERISE"].astype(int)
        res["anomalies_qual"]["Backlog planification caractérisé"] = plc["NON CARACTERISE"].astype(int)
        res["anomalies_qual"]["OT CONFIME"] = res['ot_confime']["NON"].astype(int)
        res["anomalies_qual"]["OT_COR_EGAL"] = res['ot_cor_egal']["NON"].astype(int)
        res["anomalies_qual"]["OT Fiabilité"] = 0
        res["anomalies_qual"]["Total Avis de Panne"] = 0

        return res

    # ==================== SIDEBAR ====================
    with st.sidebar:
        logo_b64 = get_logo_base64()
        if logo_b64:
            st.markdown('<img src="data:image/png;base64,%s" style="width:100%%;border-radius:8px;margin-bottom:10px">' % logo_b64, unsafe_allow_html=True)

        ot_file = st.file_uploader("Fichier OT (.xlsx/.xls)", type=["xlsx","xls"], key="ot_up")
        av_file = st.file_uploader("Fichier Avis (.xlsx/.xls)", type=["xlsx","xls"], key="av_up")

        st.markdown("---")
        poste_filter = st.multiselect("Filtrer par poste", options=[], key="poste_filt")
        st.markdown("---")
        show_histo = st.checkbox("Analyse historique", value=False)
        st.markdown("---")
        if st.button("Effacer cache", use_container_width=True):
            st.cache_data.clear()
            st.success("Cache effacé!")
            st.rerun()

    # ==================== HEADER ====================
    hdr = '<div class="mh"><h1>Dashboard KPI Maintenance</h1>'
    if logo_b64:
        hdr += '<img class="logo" src="data:image/png;base64,%s">' % logo_b64
    hdr += '<div class="db">%s</div></div>' % fichier_date
    st.markdown(hdr, unsafe_allow_html=True)

    # ==================== DATA LOAD ====================
    if not ot_file or not av_file:
        st.markdown('<div class="es" style="margin-top:80px;font-size:20px">Veuillez charger les fichiers OT et Avis dans la barre latérale pour commencer.</div>', unsafe_allow_html=True)
        st.stop()

    ot_bytes = ot_file.read()
    av_bytes = av_file.read()

    df, avf, posts, now_ts = prepare_data(ot_bytes, av_bytes, fichier_date)

    if not posts:
        st.markdown('<div class="es" style="margin-top:80px;font-size:20px">Aucun poste de travail trouvé (SF1/SF2) dans les données fournies.</div>', unsafe_allow_html=True)
        st.stop()

    if poste_filter:
        posts = [p for p in posts if p in poste_filter]

    kpi_res = calc_kpis(df, avf, now_ts, posts)
    ckdf = kpi_res['ckdf']
    ano_perf = kpi_res['anomalies_perf']
    ano_qual = kpi_res['anomalies_qual']
    dfp = kpi_res['dfp']

    # ==================== CARDS ====================
    total_ot = len(dfp)
    clotures = len(dfp[dfp["Statut OT"].isin(["CLOT","TCLO"])])
    lances = len(dfp[dfp["Statut OT"]=="LANC"])
    cres = len(dfp[dfp["Statut OT"]=="CRÉÉ"])
    total_ano_p = int(ano_perf.sum().sum())
    total_ano_q = int(ano_qual.sum().sum())

    st.markdown('<div class="cr">'
        '<div class="cc c1"><div class="cv">%d</div><div class="cl">Total OT</div></div>'
        '<div class="cc c2"><div class="cv">%d</div><div class="cl">Cloturés</div></div>'
        '<div class="cc c3"><div class="cv">%d</div><div class="cl">Lancés</div></div>'
        '<div class="cc c4"><div class="cv">%d</div><div class="cl">Créés</div></div>'
        '<div class="cc c5"><div class="cv">%d</div><div class="cl">Anomalies Perf.</div></div>'
        '<div class="cc c6"><div class="cv">%d</div><div class="cl">Anomalies Qual.</div></div>'
        '<div class="cc c7"><div class="cv">%d</div><div class="cl">Postes</div></div>'
        '<div class="cc c8"><div class="cv">%d</div><div class="cl">Avis sans OT</div></div>'
        '</div>' % (total_ot, clotures, lances, cres, total_ano_p, total_ano_q, len(posts), len(avf)),
        unsafe_allow_html=True)

    # ==================== HELPERS POUR TABLEAUX HTML ====================
    def html_kpi_table(ckdf_sub, kpi_list, table_class, score_label):
        cols = ["Poste de travail"] + kpi_list + ["Score"]
        h = '<table class="tw %s"><thead><tr>' % table_class
        h += ''.join('<th>%s</th>' % c for c in cols)
        h += '</tr></thead><tbody>'
        # Ligne Cible
        h += '<tr><td style="font-weight:800;background:#f0f4f8">Cible</td>'
        for kpi in kpi_list:
            h += '<td style="text-align:center;font-weight:800;background:#f0f4f8">%.0f</td>' % CIBLE[kpi]
        h += '<td style="background:#f0f4f8"></td></tr>'
        # Lignes postes
        scores = []
        for poste, row in ckdf_sub.iterrows():
            h += '<tr><td style="font-weight:700">%s</td>' % poste
            met = 0
            for kpi in kpi_list:
                val = round(float(row[kpi]), 1)
                target = CIBLE[kpi]
                if kpi in LOWER_BETTER:
                    good = val <= target
                else:
                    good = val >= target
                if good:
                    style = "background:#d1fae5;color:#065f46;font-weight:700;"
                    met += 1
                else:
                    style = "background:#fee2e2;color:#991b1b;font-weight:700;"
                h += '<td style="text-align:center;%s">%.1f</td>' % (style, val)
            score = round(met / len(kpi_list) * 100, 1)
            scores.append(score)
            if score >= 80:
                s_style = "background:#d1fae5;color:#065f46;font-weight:800;"
            elif score >= 50:
                s_style = "background:#fef3c7;color:#92400e;font-weight:800;"
            else:
                s_style = "background:#fee2e2;color:#991b1b;font-weight:800;"
            h += '<td style="text-align:center;%s">%.1f%%</td>' % (s_style, score)
            h += '</tr>'
        # Ligne Total
        h += '<tr class="cb"><td style="font-weight:800">Moyenne</td>'
        for kpi in kpi_list:
            h += '<td style="text-align:center;font-weight:800">%.1f</td>' % round(ckdf_sub[kpi].mean(), 1)
        avg_score = round(np.mean(scores), 1)
        h += '<td style="text-align:center;font-weight:800">%.1f%%</td>' % avg_score
        h += '</tr></tbody></table>'
        return h

    def html_anomalies_table(ano_df, kpi_list, table_class):
        cols = ["Poste de travail"] + kpi_list
        h = '<table class="tw %s"><thead><tr>' % table_class
        h += ''.join('<th>%s</th>' % c for c in cols)
        h += '</tr></thead><tbody>'
        for poste, row in ano_df.iterrows():
            h += '<tr><td style="font-weight:700;color:#000">%s</td>' % poste
            for kpi in kpi_list:
                val = int(row.get(kpi, 0))
                h += '<td style="text-align:center;color:#000;font-weight:700">%d</td>' % val
            h += '</tr>'
        # Ligne Total
        h += '<tr class="cb"><td style="font-weight:800">Total</td>'
        for kpi in kpi_list:
            total_val = int(ano_df[kpi].sum())
            h += '<td style="text-align:center;font-weight:800">%d</td>' % total_val
        h += '</tr></tbody></table>'
        return h

    # ==================== BUILD EXCEL DATA ====================
    pcols = ["Poste de travail"] + QK + ["Score Performance"]
    prows = []
    cible_row = {"Poste de travail": "Cible"}
    for kpi in QK: cible_row[kpi] = CIBLE[kpi]
    cible_row["Score Performance"] = ""
    prows.append(cible_row)
    for poste in posts:
        row = {"Poste de travail": poste}
        met = 0
        for kpi in QK:
            val = round(float(ckdf.loc[poste, kpi]), 1)
            row[kpi] = val
            target = CIBLE[kpi]
            if kpi in LOWER_BETTER:
                if val <= target: met += 1
            else:
                if val >= target: met += 1
        row["Score Performance"] = round(met / len(QK) * 100, 1)
        prows.append(row)

    qcols = ["Poste de travail"] + PK + ["Score Qualite"]
    qrows = []
    cible_row_q = {"Poste de travail": "Cible"}
    for kpi in PK: cible_row_q[kpi] = CIBLE[kpi]
    cible_row_q["Score Qualite"] = ""
    qrows.append(cible_row_q)
    for poste in posts:
        row = {"Poste de travail": poste}
        met = 0
        for kpi in PK:
            val = round(float(ckdf.loc[poste, kpi]), 1)
            row[kpi] = val
            target = CIBLE[kpi]
            if kpi in LOWER_BETTER:
                if val <= target: met += 1
            else:
                if val >= target: met += 1
        row["Score Qualite"] = round(met / len(PK) * 100, 1)
        qrows.append(row)

    ano_p_c = ["Poste de travail"] + QK
    ano_p_r = []
    for poste in posts:
        row = {"Poste de travail": poste}
        for kpi in QK:
            row[kpi] = int(ano_perf.loc[poste, kpi])
        ano_p_r.append(row)
    total_ano_p_row = {"Poste de travail": "Total"}
    for kpi in QK:
        total_ano_p_row[kpi] = int(ano_perf[kpi].sum())
    ano_p_r.append(total_ano_p_row)

    ano_q_c = ["Poste de travail"] + PK
    ano_q_r = []
    for poste in posts:
        row = {"Poste de travail": poste}
        for kpi in PK:
            row[kpi] = int(ano_qual.loc[poste, kpi])
        ano_q_r.append(row)
    total_ano_q_row = {"Poste de travail": "Total"}
    for kpi in PK:
        total_ano_q_row[kpi] = int(ano_qual[kpi].sum())
    ano_q_r.append(total_ano_q_row)

    save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, fichier_date)

    # ==================== TABS ====================
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Performance", "📋 Qualité", "📈 Synthèse", "📉 Historique", "🎯 Plan d'Action"])

    # ==================== TAB PERFORMANCE ====================
    with tab1:
        st.markdown('<div class="stl">Indicateurs de Performance</div>', unsafe_allow_html=True)
        st.markdown(html_kpi_table(ckdf[QK], QK, "pt", "Score Performance"), unsafe_allow_html=True)

        st.markdown('<div class="stl">Anomalies Performance</div>', unsafe_allow_html=True)
        st.markdown(html_anomalies_table(ano_perf, QK, "at"), unsafe_allow_html=True)

        # Barres de progression Performance
        st.markdown('<div class="stl">Détail par KPI - Performance</div>', unsafe_allow_html=True)
        for kpi in QK:
            target = CIBLE[kpi]
            avg_val = round(ckdf[kpi].mean(), 1)
            pct = min(avg_val, 120) / 120 * 100
            if kpi in LOWER_BETTER:
                bar_color = "#10b981" if avg_val <= target else "#ef4444"
            else:
                bar_color = "#10b981" if avg_val >= target else "#ef4444"
            target_pct = target / 120 * 100
            st.markdown('<div class="ca"><div class="ct">%s</div>'
                '<div class="car">'
                '<div class="cal">%s</div>'
                '<div class="cab">'
                '<div class="caf" style="width:%.1f%%;background:%s"></div>'
                '<div class="target-mark" style="left:%.1f%%"></div>'
                '</div>'
                '<div class="cav-out">%.1f%%</div>'
                '<div class="cav-tgt">(Cible: %.0f)</div>'
                '</div></div>' % (kpi, kpi, pct, bar_color, target_pct, avg_val, target),
                unsafe_allow_html=True)

        # Graphiques Performance
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            piv_corr = build_statut_pivot(dfp[(dfp["Nº appel pl.entret."].fillna(0)==0)&(dfp["Contient SOPL"]==1)], posts)
            show_pie_pair(piv_corr, "Correctifs")
        with col_g2:
            st.markdown('<div class="stl">Age Preparation (CRÉÉ)</div>', unsafe_allow_html=True)
            pr_piv = pd.pivot_table(dfp[(dfp["Statut OT"]=="CRÉÉ")&(dfp["Statut utilisateur"].str.contains(r"\bCRPR\b",case=False,na=False))],
                                    index="Poste travail princ.", columns="ap", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            show_simple_pie(pr_piv, "Répartition par âge - Préparation")

    # ==================== TAB QUALITE ====================
    with tab2:
        st.markdown('<div class="stl">Indicateurs de Qualité</div>', unsafe_allow_html=True)
        st.markdown(html_kpi_table(ckdf[PK], PK, "qt", "Score Qualite"), unsafe_allow_html=True)

        st.markdown('<div class="stl">Anomalies Qualité</div>', unsafe_allow_html=True)
        st.markdown(html_anomalies_table(ano_qual, PK, "at"), unsafe_allow_html=True)

        # Barres de progression Qualité
        st.markdown('<div class="stl">Détail par KPI - Qualité</div>', unsafe_allow_html=True)
        for kpi in PK:
            target = CIBLE[kpi]
            avg_val = round(ckdf[kpi].mean(), 1)
            pct = min(avg_val, 120) / 120 * 100
            if kpi in LOWER_BETTER:
                bar_color = "#10b981" if avg_val <= target else "#ef4444"
            else:
                bar_color = "#10b981" if avg_val >= target else "#ef4444"
            target_pct = target / 120 * 100
            st.markdown('<div class="ca"><div class="ct">%s</div>'
                '<div class="car">'
                '<div class="cal">%s</div>'
                '<div class="cab">'
                '<div class="caf" style="width:%.1f%%;background:%s"></div>'
                '<div class="target-mark" style="left:%.1f%%"></div>'
                '</div>'
                '<div class="cav-out">%.1f%%</div>'
                '<div class="cav-tgt">(Cible: %.0f)</div>'
                '</div></div>' % (kpi, kpi, pct, bar_color, target_pct, avg_val, target),
                unsafe_allow_html=True)

        # Graphiques Qualité
        col_q1, col_q2 = st.columns(2)
        with col_q1:
            pc_piv = pd.pivot_table(dfp[dfp["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            show_simple_pie(pc_piv, "Backlog Préparation - Caractérisation")
        with col_q2:
            plc_piv = pd.pivot_table(dfp[(dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==0)], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            show_simple_pie(plc_piv, "Backlog Planification - Caractérisation")

    # ==================== TAB SYNTHESE ====================
    with tab3:
        st.markdown('<div class="stl">Tableau de Synthèse Général</div>', unsafe_allow_html=True)
        all_kpi = QK + PK
        synth_cols = ["Poste de travail"] + all_kpi + ["Score Perf.", "Score Qual."]
        h = '<table class="tw synth-tbl"><thead><tr>'
        h += '<th class="poste-cell">Poste</th>'
        for kpi in all_kpi:
            short = kpi.replace("TAUX_REALISATION_CORRECTIF/PT","TAUX REAL.").replace("Performance ","Perf. ").replace("préparation","Prép.").replace("planification","Plan.").replace("exécution","Exec.").replace("1mois< <3mois","1-3m").replace("caractérisé","Carac.").replace("Appels Systématiques","App.Sys.").replace("d'approbation des Avis","Approb.Avis")
            h += '<th>%s</th>' % short
        h += '<th>Score P.</th><th>Score Q.</th></tr></thead><tbody>'

        for poste in posts:
            h += '<tr><td class="poste-cell">%s</td>' % poste
            met_p = 0; met_q = 0
            for kpi in all_kpi:
                val = round(float(ckdf.loc[poste, kpi]), 1)
                target = CIBLE[kpi]
                if kpi in LOWER_BETTER:
                    good = val <= target
                else:
                    good = val >= target
                if good:
                    style = "background:#d1fae5;color:#065f46;font-weight:600;"
                    if kpi in QK: met_p += 1
                    else: met_q += 1
                else:
                    style = "background:#fee2e2;color:#991b1b;font-weight:600;"
                h += '<td style="%s">%.1f</td>' % (style, val)
            sp = round(met_p / len(QK) * 100, 1)
            sq = round(met_q / len(PK) * 100, 1)
            sp_style = "background:#d1fae5;color:#065f46;font-weight:700;" if sp >= 80 else "background:#fee2e2;color:#991b1b;font-weight:700;"
            sq_style = "background:#d1fae5;color:#065f46;font-weight:700;" if sq >= 80 else "background:#fee2e2;color:#991b1b;font-weight:700;"
            h += '<td style="%s">%.1f%%</td>' % (sp_style, sp)
            h += '<td style="%s">%.1f%%</td>' % (sq_style, sq)
            h += '</tr>'
        h += '</tbody></table>'
        st.markdown(h, unsafe_allow_html=True)

        # Radar chart
        st.markdown('<div class="stl">Radar - Moyenne Générale</div>', unsafe_allow_html=True)
        radar_kpi = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois","Performance Graissage","Performance Inspection","Taux d'approbation des Avis","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé"]
        radar_labels = [k.replace("TAUX_REALISATION_CORRECTIF/PT","Taux Réal.").replace("OT préparation <1 mois","Prép.<1m").replace("OT planification <1 mois","Plan.<1m").replace("OT exécution <1 mois","Exec.<1m").replace("Performance Graissage","Graissage").replace("Performance Inspection","Inspection").replace("Taux d'approbation des Avis","Approb.Avis").replace("OT LANC ESTIME","OT Estimés").replace("Backlog préparation caractérisé","Back.Prep").replace("Backlog planification caractérisé","Back.Plan") for k in radar_kpi]
        radar_vals = [round(ckdf[k].mean(), 1) for k in radar_kpi]
        radar_targets = [CIBLE[k] for k in radar_kpi]
        fig_r = go.Figure()
        fig_r.add_trace(go.Scatterpolar(r=radar_vals, theta=radar_labels, fill='toself', name='Valeur', line_color='#3b82f6'))
        fig_r.add_trace(go.Scatterpolar(r=radar_targets, theta=radar_labels, fill='toself', name='Cible', line_color='#ef4444', line_dash='dash'))
        fig_r.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 120])), height=500, showlegend=True)
        st.plotly_chart(fig_r, use_container_width=True)

    # ==================== TAB HISTORIQUE ====================
    with tab4:
        if show_histo:
            kpis_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
            hist_df = load_historical_kpis(kpis_path)
            if hist_df.empty:
                st.markdown('<div class="es">Aucune donnée historique disponible. Les données sont sauvegardées à chaque exécution.</div>', unsafe_allow_html=True)
            else:
                var_df = calculate_variations(hist_df)
                if var_df.empty:
                    st.markdown('<div class="es">Pas assez de périodes pour calculer les variations.</div>', unsafe_allow_html=True)
                else:
                    journal = generate_journal(var_df)
                    if journal.empty:
                        st.markdown('<div class="es">Aucune variation significative (>=5%%) détectée.</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="stl">Journal des Variations Significatives (>=5%%)</div>', unsafe_allow_html=True)
                        st.dataframe(journal[["Date precedente","Date actuelle","Poste","Type","KPI","Valeur precedente","Valeur actuelle","Ecart %","Sens"]].reset_index(drop=True), use_container_width=True, height=500)

                    best5, worst5 = calculate_rankings(var_df)
                    if not best5.empty and not worst5.empty:
                        col_r1, col_r2 = st.columns(2)
                        with col_r1:
                            st.markdown('<div class="stl">Top 5 Amélioration</div>', unsafe_allow_html=True)
                            st.dataframe(best5, use_container_width=True, hide_index=True)
                        with col_r2:
                            st.markdown('<div class="stl">Top 5 Dégradation</div>', unsafe_allow_html=True)
                            st.dataframe(worst5, use_container_width=True, hide_index=True)

                    # Graphique évolution
                    st.markdown('<div class="stl">Evolution des Scores</div>', unsafe_allow_html=True)
                    perf_hist = hist_df[hist_df["_section"]=="perf"].copy()
                    qual_hist = hist_df[hist_df["_section"]=="qual"].copy()
                    if "Score Performance" in perf_hist.columns and "Date_parsed" in perf_hist.columns:
                        score_perf_by_date = perf_hist.groupby("Date_parsed")["Score Performance"].mean()
                        fig_ev = go.Figure()
                        fig_ev.add_trace(go.Scatter(x=score_perf_by_date.index, y=score_perf_by_date.values, mode='lines+markers', name='Score Performance', line_color='#10b981'))
                        if "Score Qualite" in qual_hist.columns:
                            score_qual_by_date = qual_hist.groupby("Date_parsed")["Score Qualite"].mean()
                            fig_ev.add_trace(go.Scatter(x=score_qual_by_date.index, y=score_qual_by_date.values, mode='lines+markers', name='Score Qualité', line_color='#3b82f6'))
                        fig_ev.update_layout(height=400, yaxis_title="Score (%)", xaxis_title="Date")
                        st.plotly_chart(fig_ev, use_container_width=True)
        else:
            st.markdown('<div class="es">Activez "Analyse historique" dans la barre latérale pour voir cette section.</div>', unsafe_allow_html=True)

    # ==================== TAB PLAN D'ACTION ====================
    with tab5:
        st.markdown('<div class="stl">Plan d\'Action - KPIs Non Atteints</div>', unsafe_allow_html=True)

        plan_rows = []
        for kpi in ALL_KPI:
            target = CIBLE[kpi]
            avg_val = round(ckdf[kpi].mean(), 1)
            if kpi in LOWER_BETTER:
                atteint = avg_val <= target
            else:
                atteint = avg_val >= target
            if not atteint:
                resp = KPI_RESP_MAP.get(kpi, "")
                action = ACT_MAP.get(kpi, "")
                ecart = round(avg_val - target, 1)
                plan_rows.append({
                    "KPI": kpi,
                    "Moyenne": avg_val,
                    "Cible": target,
                    "Ecart": ecart,
                    "Responsable": resp,
                    "Action": action
                })

        if not plan_rows:
            st.markdown('<div class="es" style="color:#10b981;font-size:18px;font-weight:700">Tous les KPIs sont atteints ! Félicitations.</div>', unsafe_allow_html=True)
        else:
            h = '<table class="plan-action-table"><thead><tr>'
            h += '<th>KPI</th><th>Moyenne</th><th>Cible</th><th>Ecart</th><th>Responsable</th><th>Action Corrective</th>'
            h += '</tr></thead><tbody>'
            for r in plan_rows:
                ecart_color = "#991b1b" if r["Ecart"] > 0 and r["KPI"] not in LOWER_BETTER else "#991b1b"
                h += '<tr>'
                h += '<td style="font-weight:800;text-align:left">%s</td>' % r["KPI"]
                h += '<td>%.1f%%</td>' % r["Moyenne"]
                h += '<td>%.0f%%</td>' % r["Cible"]
                h += '<td style="color:%s;font-weight:700">%+.1f</td>' % (ecart_color, r["Ecart"])
                h += '<td>%s</td>' % r["Responsable"]
                h += '<td style="text-align:left">%s</td>' % r["Action"]
                h += '</tr>'
            h += '</tbody></table>'
            st.markdown(h, unsafe_allow_html=True)

        # Détail par poste des KPIs non atteints
        st.markdown('<div class="stl">Détail par Poste - KPIs Non Atteints</div>', unsafe_allow_html=True)
        detail_rows = []
        for poste in posts:
            for kpi in ALL_KPI:
                val = round(float(ckdf.loc[poste, kpi]), 1)
                target = CIBLE[kpi]
                if kpi in LOWER_BETTER:
                    ok = val <= target
                else:
                    ok = val >= target
                if not ok:
                    detail_rows.append({
                        "Poste": poste,
                        "KPI": kpi,
                        "Valeur": val,
                        "Cible": target,
                        "Ecart": round(val - target, 1),
                        "Responsable": KPI_RESP_MAP.get(kpi, ""),
                        "Action": ACT_MAP.get(kpi, "")
                    })

        if detail_rows:
            h2 = '<table class="plan-action-table"><thead><tr>'
            h2 += '<th>Poste</th><th>KPI</th><th>Valeur</th><th>Cible</th><th>Ecart</th><th>Responsable</th><th>Action</th>'
            h2 += '</tr></thead><tbody>'
            for r in detail_rows:
                h2 += '<tr>'
                h2 += '<td style="text-align:left;font-weight:700">%s</td>' % r["Poste"]
                h2 += '<td style="text-align:left">%s</td>' % r["KPI"]
                h2 += '<td>%.1f%%</td>' % r["Valeur"]
                h2 += '<td>%.0f%%</td>' % r["Cible"]
                h2 += '<td style="color:#991b1b;font-weight:700">%+.1f</td>' % r["Ecart"]
                h2 += '<td>%s</td>' % r["Responsable"]
                h2 += '<td style="text-align:left">%s</td>' % r["Action"]
                h2 += '</tr>'
            h2 += '</tbody></table>'
            st.markdown(h2, unsafe_allow_html=True)
        else:
            st.markdown('<div class="es" style="color:#10b981">Aucun KPI non atteint par poste.</div>', unsafe_allow_html=True)

    # ==================== FOOTER ====================
    st.markdown('<div class="footer">Dashboard KPI Maintenance — %s — Généré automatiquement</div>' % fichier_date, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
