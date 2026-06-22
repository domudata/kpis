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
    raise ValueError("Format de fichier non reconnu.")

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
    avf = raw_av[(raw_av["Ordre"].isna()|(raw_av["Ordre"].astype(str).str.strip()==""))&(raw_av["Type d'avis"].isin(["ZU","Z4","ZR","ZP"]))].copy()
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
# NOUVEAU : Page Suivi & Évaluation
# ============================================================

def page_suivi_evaluation(posts, hist_df, var_df):
    st.markdown(
        '<div class="stl">📋 Résumé — Journal des Variations Significatives</div>',
        unsafe_allow_html=True
    )

    if hist_df.empty:
        st.markdown(
            """<div class="es" style="padding:50px 20px;">
                <div style="font-size:56px;margin-bottom:16px;">📊</div>
                <div style="font-size:16px;font-weight:700;color:#1e293b;margin-bottom:8px;">
                    Aucune donnée historique disponible
                </div>
                <div style="font-size:13px;color:#64748b;">
                    Générez au moins 2 rapports depuis l'onglet <strong>Synthèse</strong> pour visualiser les évolutions.
                </div>
            </div>""",
            unsafe_allow_html=True
        )
        return

    dates_labels = sorted(hist_df["Date"].unique())
    dates_parsed = sorted(hist_df["Date_parsed"].dropna().unique())

    if len(dates_parsed) < 2:
        st.markdown(
            f"""<div class="es" style="padding:50px 20px;">
                <div style="font-size:56px;margin-bottom:16px;">📈</div>
                <div style="font-size:16px;font-weight:700;color:#1e293b;margin-bottom:8px;">
                    Données insuffisantes
                </div>
                <div style="font-size:13px;color:#64748b;">
                    Au moins <strong>2 périodes</strong> sont nécessaires. Période(s) disponible(s) : <strong>{len(dates_parsed)}</strong>
                </div>
            </div>""",
            unsafe_allow_html=True
        )
        return

    perf_hist = hist_df[hist_df["_section"] == "perf"].copy()
    qual_hist = hist_df[hist_df["_section"] == "qual"].copy()

    st.markdown(
        f"""<div style="
            background:linear-gradient(135deg,#eff6ff 0%,#e0e7ff 100%);
            padding:14px 22px;
            border-radius:10px;
            border-left:5px solid #3b82f6;
            margin-bottom:18px;
            font-size:14px;
            display:flex;
            align-items:center;
            gap:20px;
            flex-wrap:wrap;
            box-shadow:0 2px 8px rgba(59,130,246,0.08);
        ">
            <span style="display:flex;align-items:center;gap:6px;">
                <span style="font-size:18px;">📅</span>
                <strong>Période :</strong> {dates_labels[0]} → {dates_labels[-1]}
            </span>
            <span style="display:flex;align-items:center;gap:6px;">
                <span style="font-size:18px;">📊</span>
                <strong>Périodes :</strong> {len(dates_labels)}
            </span>
            <span style="display:flex;align-items:center;gap:6px;">
                <span style="font-size:18px;">🏭</span>
                <strong>Postes :</strong> {len(posts)}
            </span>
        </div>""",
        unsafe_allow_html=True
    )

    st.markdown("""<style>
    .poste-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #1e40af 100%);
        color: #fff;
        padding: 10px 18px;
        border-radius: 10px 10px 0 0;
        font-weight: 800;
        font-size: 14px;
        display: flex;
        align-items: center;
        gap: 10px;
        letter-spacing: 0.3px;
        box-shadow: 0 2px 8px rgba(30,58,95,0.15);
    }
    .poste-header .badge {
        background: rgba(255,255,255,0.18);
        padding: 3px 12px;
        border-radius: 5px;
        font-size: 10px;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        font-weight: 700;
    }
    .chart-label {
        text-align: center;
        font-weight: 800;
        font-size: 12px;
        padding: 7px 10px;
        border-radius: 8px 8px 0 0;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    .chart-label.perf { color: #059669; background: #ecfdf5; border: 1px solid #d1fae5; border-bottom: none; }
    .chart-label.qual { color: #2563eb; background: #eff6ff; border: 1px solid #dbeafe; border-bottom: none; }
    .chart-footer {
        text-align: center;
        font-size: 11px;
        padding: 5px 8px;
        border-radius: 0 0 8px 8px;
    }
    .chart-footer.perf { background: #ecfdf5; border: 1px solid #d1fae5; border-top: none; }
    .chart-footer.qual { background: #eff6ff; border: 1px solid #dbeafe; border-top: none; }
    .poste-separator { height: 20px; }
    .synth-card {
        background: #fff;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        padding: 16px;
        margin: 10px 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        overflow-x: auto;
    }
    .synth-card h3 {
        font-size: 15px;
        font-weight: 800;
        margin: 0 0 12px 0;
        padding-bottom: 8px;
        border-bottom: 2px solid #e2e8f0;
    }
    .synth-tbl-detailed {
        width: 100%;
        border-collapse: collapse;
        font-family: 'Inter', sans-serif;
        font-size: 11px;
    }
    .synth-tbl-detailed thead th {
        background: #1e3a5f;
        color: #fff;
        font-weight: 700;
        font-size: 10px;
        padding: 8px 6px;
        border: 1px solid #1e3a5f;
        white-space: nowrap;
        text-transform: uppercase;
        letter-spacing: 0.3px;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .synth-tbl-detailed thead th:first-child {
        position: sticky;
        left: 0;
        z-index: 11;
        min-width: 140px;
    }
    .synth-tbl-detailed tbody td {
        padding: 6px 5px;
        border: 1px solid #e2e8f0;
        text-align: center;
        vertical-align: middle;
        color: #1e293b !important;
    }
    .synth-tbl-detailed tbody td:first-child {
        position: sticky;
        left: 0;
        background: #fff !important;
        z-index: 5;
        font-weight: 800;
        text-align: left;
        white-space: nowrap;
        border-right: 2px solid #e2e8f0;
    }
    .synth-tbl-detailed tbody tr:nth-child(even) td:first-child { background: #f8fafc !important; }
    .synth-tbl-detailed tbody tr:hover td:first-child { background: #eff6ff !important; }
    .synth-tbl-detailed tbody tr:nth-child(even) td { background: #f8fafc; }
    .synth-tbl-detailed tbody tr:hover td { background: #eff6ff !important; }
    .cell-improv { background: #d1fae5 !important; color: #065f46 !important; font-weight: 700; }
    .cell-degrad { background: #fee2e2 !important; color: #991b1b !important; font-weight: 700; }
    .cell-stable { background: #f8fafc !important; color: #475569 !important; }
    .cell-arrow { font-size: 10px; display: block; margin-top: 2px; }
    .detail-btn-row { display: flex; gap: 12px; margin: 16px 0 8px 0; }
    .detail-btn-row > div { flex: 1; }
    .kpi-summary-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
        gap: 8px;
        margin: 12px 0;
    }
    .kpi-summary-card {
        border-radius: 8px;
        padding: 10px 12px;
        text-align: center;
        border: 1px solid;
    }
    .kpi-summary-card .kpi-val { font-size: 22px; font-weight: 900; }
    .kpi-summary-card .kpi-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }
    .kpi-summary-card .kpi-diff { font-size: 11px; font-weight: 700; margin-top: 2px; }
    </style>""", unsafe_allow_html=True)

    for idx, poste in enumerate(posts):
        p_data = perf_hist[perf_hist["Poste de travail"] == poste].sort_values("Date_parsed")
        q_data = qual_hist[qual_hist["Poste de travail"] == poste].sort_values("Date_parsed")

        accent = "#10b981" if idx % 2 == 0 else "#3b82f6"
        st.markdown(
            f"""<div class="poste-header" style="border-left:5px solid {accent};">
                <span class="badge">Poste</span>
                {poste}
            </div>""",
            unsafe_allow_html=True
        )

        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div class="chart-label perf">📈 Performance Globale</div>', unsafe_allow_html=True)
            if not p_data.empty and "Score Performance" in p_data.columns:
                vals = p_data["Score Performance"].values
                last_v = vals[-1]
                prev_v = vals[-2] if len(vals) > 1 else last_v
                diff = last_v - prev_v
                is_up = diff >= 0
                trend_icon = "▲" if is_up else "▼"
                trend_clr = "#059669" if is_up else "#ef4444"

                fig_p = go.Figure()
                fig_p.add_trace(go.Scatter(
                    x=p_data["Date_parsed"], y=p_data["Score Performance"],
                    fill='tozeroy', fillcolor='rgba(5,150,105,0.07)',
                    line=dict(width=0), showlegend=False, hoverinfo='skip'
                ))
                fig_p.add_trace(go.Scatter(
                    x=p_data["Date_parsed"], y=p_data["Score Performance"],
                    mode='lines+markers', name='Performance',
                    line=dict(color='#059669', width=3, shape='spline'),
                    marker=dict(size=9, color='#059669', line=dict(color='white', width=2.5)),
                    hovertemplate='<b>%{x|%d/%m/%Y}</b><br>Performance : <b>%{y:.1f}%</b><extra></extra>'
                ))
                fig_p.add_hline(y=85, line_dash="dash", line_color="#ef4444", line_width=1.5,
                    annotation_text="Cible 85%", annotation_position="top right",
                    annotation_font=dict(size=9, color="#ef4444", family="Inter"))
                fig_p.update_layout(
                    margin=dict(t=8, b=32, l=48, r=16), height=210,
                    yaxis=dict(range=[0, 110], dtick=20, title=dict(text="%", font=dict(size=10, color="#64748b"))),
                    xaxis=dict(showgrid=True, gridcolor='#f1f5f9', tickfont=dict(size=9, color="#64748b")),
                    showlegend=False, plot_bgcolor='#fafbfc', paper_bgcolor='white'
                )
                st.plotly_chart(fig_p, use_container_width=True, config={"displayModeBar": False})
                st.markdown(
                    f"""<div class="chart-footer perf">
                        <span style="color:{trend_clr};font-weight:900;font-size:14px;">{trend_icon} {last_v:.1f}%</span>
                        <span style="color:#64748b;margin-left:8px;">vs {prev_v:.1f}% préc.</span>
                        <span style="color:{trend_clr};margin-left:4px;font-weight:700;">({diff:+.1f})</span>
                    </div>""", unsafe_allow_html=True)
            else:
                st.markdown('<div class="es" style="padding:30px 10px;background:#fafbfc;border:1px solid #e2e8f0;border-radius:0 0 8px 8px;font-size:12px;">Données indisponibles</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="chart-label qual">📈 Qualité Globale</div>', unsafe_allow_html=True)
            if not q_data.empty and "Score Qualite" in q_data.columns:
                vals = q_data["Score Qualite"].values
                last_v = vals[-1]
                prev_v = vals[-2] if len(vals) > 1 else last_v
                diff = last_v - prev_v
                is_up = diff >= 0
                trend_icon = "▲" if is_up else "▼"
                trend_clr = "#2563eb" if is_up else "#ef4444"

                fig_q = go.Figure()
                fig_q.add_trace(go.Scatter(
                    x=q_data["Date_parsed"], y=q_data["Score Qualite"],
                    fill='tozeroy', fillcolor='rgba(37,99,235,0.07)',
                    line=dict(width=0), showlegend=False, hoverinfo='skip'
                ))
                fig_q.add_trace(go.Scatter(
                    x=q_data["Date_parsed"], y=q_data["Score Qualite"],
                    mode='lines+markers', name='Qualité',
                    line=dict(color='#2563eb', width=3, shape='spline'),
                    marker=dict(size=9, color='#2563eb', line=dict(color='white', width=2.5)),
                    hovertemplate='<b>%{x|%d/%m/%Y}</b><br>Qualité : <b>%{y:.1f}%</b><extra></extra>'
                ))
                fig_q.add_hline(y=85, line_dash="dash", line_color="#ef4444", line_width=1.5,
                    annotation_text="Cible 85%", annotation_position="top right",
                    annotation_font=dict(size=9, color="#ef4444", family="Inter"))
                fig_q.update_layout(
                    margin=dict(t=8, b=32, l=48, r=16), height=210,
                    yaxis=dict(range=[0, 110], dtick=20, title=dict(text="%", font=dict(size=10, color="#64748b"))),
                    xaxis=dict(showgrid=True, gridcolor='#f1f5f9', tickfont=dict(size=9, color="#64748b")),
                    showlegend=False, plot_bgcolor='#fafbfc', paper_bgcolor='white'
                )
                st.plotly_chart(fig_q, use_container_width=True, config={"displayModeBar": False})
                st.markdown(
                    f"""<div class="chart-footer qual">
                        <span style="color:{trend_clr};font-weight:900;font-size:14px;">{trend_icon} {last_v:.1f}%</span>
                        <span style="color:#64748b;margin-left:8px;">vs {prev_v:.1f}% préc.</span>
                        <span style="color:{trend_clr};margin-left:4px;font-weight:700;">({diff:+.1f})</span>
                    </div>""", unsafe_allow_html=True)
            else:
                st.markdown('<div class="es" style="padding:30px 10px;background:#fafbfc;border:1px solid #e2e8f0;border-radius:0 0 8px 8px;font-size:12px;">Données indisponibles</div>', unsafe_allow_html=True)

        st.markdown('<div class="poste-separator"></div>', unsafe_allow_html=True)

    # ── Synthèse d'évolution détaillée ──
    st.markdown('<div class="stl">🔍 Synthèse d\'Évolution Détaillée</div>', unsafe_allow_html=True)

    sel_c1, sel_c2 = st.columns(2)
    with sel_c1:
        idx_from = max(0, len(dates_labels) - 2)
        date_from = st.selectbox("📅 Date de début", dates_labels, index=idx_from,
            key="synth_date_from", format_func=lambda x: f"Période : {x}")
    with sel_c2:
        date_to = st.selectbox("📅 Date de fin", dates_labels, index=len(dates_labels) - 1,
            key="synth_date_to", format_func=lambda x: f"Période : {x}")

    idx_from_val = dates_labels.index(date_from)
    idx_to_val = dates_labels.index(date_to)
    if idx_from_val >= idx_to_val:
        st.markdown("""<div style="background:#fef3c7;padding:12px 20px;border-radius:8px;border-left:4px solid #f59e0b;
            font-size:13px;color:#92400e;font-weight:600;">⚠️ La date de début doit être antérieure à la date de fin.</div>""", unsafe_allow_html=True)
        return

    st.markdown('<div class="detail-btn-row">', unsafe_allow_html=True)
    bc1, bc2 = st.columns(2)
    with bc1:
        btn_perf = st.button("📋 Détails — Synthèse d'évolution Performance", key="btn_detail_perf", use_container_width=True)
    with bc2:
        btn_qual = st.button("📋 Détails — Synthèse d'évolution Qualité", key="btn_detail_qual", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_perf:
        _render_synthesis_detail(perf_hist, date_from, date_to, posts, "Performance", QK, "Score Performance", "#059669", "#ecfdf5", "#d1fae5")
    if btn_qual:
        _render_synthesis_detail(qual_hist, date_from, date_to, posts, "Qualité", PK, "Score Qualite", "#2563eb", "#eff6ff", "#dbeafe")


def _render_synthesis_detail(hist_section, date_from, date_to, posts,
                              section_name, kpi_list, score_col,
                              accent_color, bg_light, bg_border):
    prev_data = hist_section[hist_section["Date"] == date_from].copy()
    curr_data = hist_section[hist_section["Date"] == date_to].copy()

    if prev_data.empty or curr_data.empty:
        st.markdown(f"""<div class="es" style="padding:30px;background:#fef3c7;border:1px solid #fde68a;border-radius:8px;">
            ⚠️ Données indisponibles pour l'une des périodes sélectionnées ({date_from} ou {date_to}).</div>""", unsafe_allow_html=True)
        return

    st.markdown(
        f"""<div style="background:linear-gradient(135deg,{bg_light},#fff);padding:14px 22px;border-radius:10px;
            border-left:5px solid {accent_color};margin:14px 0 10px 0;box-shadow:0 2px 10px rgba(0,0,0,0.05);">
            <div style="font-size:16px;font-weight:900;color:{accent_color};margin-bottom:4px;">
                📊 Synthèse d'évolution {section_name}</div>
            <div style="font-size:13px;color:#475569;">
                Entre <strong>{date_from}</strong> et <strong>{date_to}</strong> — {len(posts)} postes — {len(kpi_list)} indicateurs</div>
        </div>""", unsafe_allow_html=True)

    all_kpi_cols = kpi_list + [score_col]
    nb_improv = nb_degrad = nb_stable = 0
    total_diff_score = 0; nb_postes_score = 0

    for poste in posts:
        pr = prev_data[prev_data["Poste de travail"] == poste]
        cr = curr_data[curr_data["Poste de travail"] == poste]
        for kpi in all_kpi_cols:
            if kpi not in pr.columns or kpi not in cr.columns: continue
            try: pv = float(pr[kpi].values[0]) if len(pr) > 0 else None
            except Exception: pv = None
            try: cv = float(cr[kpi].values[0]) if len(cr) > 0 else None
            except Exception: cv = None
            if pv is None or cv is None: continue
            diff = cv - pv; is_lb = kpi in LOWER_BETTER
            if abs(diff) <= 0.5: nb_stable += 1
            elif (diff > 0.5 and not is_lb) or (diff < -0.5 and is_lb): nb_improv += 1
            else: nb_degrad += 1
            if kpi == score_col: total_diff_score += diff; nb_postes_score += 1

    avg_score_diff = (total_diff_score / nb_postes_score) if nb_postes_score > 0 else 0
    total_cells = nb_improv + nb_degrad + nb_stable

    st.markdown(
        f"""<div class="kpi-summary-grid">
            <div class="kpi-summary-card" style="border-color:#d1fae5;background:#f0fdf4;">
                <div class="kpi-val" style="color:#059669;">{nb_improv}</div>
                <div class="kpi-label" style="color:#065f46;">Améliorations</div>
                <div class="kpi-diff" style="color:#059669;">{((nb_improv/total_cells)*100) if total_cells else 0:.0f}% des KPIs</div>
            </div>
            <div class="kpi-summary-card" style="border-color:#fee2e2;background:#fef2f2;">
                <div class="kpi-val" style="color:#ef4444;">{nb_degrad}</div>
                <div class="kpi-label" style="color:#991b1b;">Dégradations</div>
                <div class="kpi-diff" style="color:#ef4444;">{((nb_degrad/total_cells)*100) if total_cells else 0:.0f}% des KPIs</div>
            </div>
            <div class="kpi-summary-card" style="border-color:#e2e8f0;background:#f8fafc;">
                <div class="kpi-val" style="color:#64748b;">{nb_stable}</div>
                <div class="kpi-label" style="color:#475569;">Stables</div>
                <div class="kpi-diff" style="color:#64748b;">{((nb_stable/total_cells)*100) if total_cells else 0:.0f}% des KPIs</div>
            </div>
            <div class="kpi-summary-card" style="border-color:{bg_border};background:{bg_light};">
                <div class="kpi-val" style="color:{accent_color};">{avg_score_diff:+.1f}</div>
                <div class="kpi-label" style="color:{accent_color};">Moy. Score {section_name}</div>
                <div class="kpi-diff" style="color:{accent_color};">pts sur la période</div>
            </div>
        </div>""", unsafe_allow_html=True)

    cols_header = ["Poste de travail"] + [k for k in kpi_list] + [score_col]
    html = f'<div class="synth-card"><h3 style="color:{accent_color};">Tableau détaillé — {section_name} ({date_from} → {date_to})</h3>'
    html += '<div style="overflow-x:auto;max-height:600px;overflow-y:auto;"><table class="synth-tbl-detailed"><thead><tr>'
    for c in cols_header:
        display_name = c
        if len(display_name) > 18 and c != score_col:
            display_name = display_name.replace("préparation", "prép.").replace("planification", "plan.").replace("exécution", "exec.")
        if c == score_col: display_name = "★ Score"
        html += f'<th title="{c}">{display_name}</th>'
    html += '</tr></thead><tbody>'

    for poste in posts:
        pr = prev_data[prev_data["Poste de travail"] == poste]
        cr = curr_data[curr_data["Poste de travail"] == poste]
        html += f'<tr><td>{poste}</td>'
        for kpi in all_kpi_cols:
            if kpi not in pr.columns or kpi not in cr.columns:
                html += '<td style="color:#94a3b8;">—</td>'; continue
            try: pv = float(pr[kpi].values[0]) if len(pr) > 0 else None
            except Exception: pv = None
            try: cv = float(cr[kpi].values[0]) if len(cr) > 0 else None
            except Exception: cv = None
            if pv is None or cv is None:
                html += '<td style="color:#94a3b8;">—</td>'; continue
            diff = cv - pv; pct = (diff / pv * 100) if pv != 0 else 0; is_lb = kpi in LOWER_BETTER
            if abs(diff) <= 0.5: css_class = "cell-stable"; arrow = "—"; arrow_color = "#94a3b8"
            elif (diff > 0.5 and not is_lb) or (diff < -0.5 and is_lb): css_class = "cell-improv"; arrow = "▲"; arrow_color = "#059669"
            else: css_class = "cell-degrad"; arrow = "▼"; arrow_color = "#ef4444"
            html += f'<td class="{css_class}">{cv:.1f}%<span class="cell-arrow" style="color:{arrow_color};">{arrow} {pct:+.1f}%</span></td>'
        html += '</tr>'

    html += '<tr style="background:#1e3a5f !important;"><td style="color:#fff !important;font-weight:900;">MOYENNE</td>'
    for kpi in all_kpi_cols:
        vals_curr = []
        for poste in posts:
            cr = curr_data[curr_data["Poste de travail"] == poste]
            if kpi in cr.columns and len(cr) > 0:
                try: vals_curr.append(float(cr[kpi].values[0]))
                except Exception: pass
        if vals_curr:
            avg = sum(vals_curr) / len(vals_curr)
            html += f'<td style="background:#1e3a5f !important;color:#fff !important;font-weight:800;font-size:12px;">{avg:.1f}%</td>'
        else:
            html += '<td style="background:#1e3a5f !important;color:rgba(255,255,255,0.4) !important;">—</td>'
    html += '</tr></tbody></table></div></div>'
    st.markdown(html, unsafe_allow_html=True)

    st.markdown(
        f"""<div style="display:flex;gap:20px;flex-wrap:wrap;padding:8px 0;font-size:12px;font-weight:600;">
            <span style="display:flex;align-items:center;gap:5px;">
                <span style="display:inline-block;width:14px;height:14px;background:#d1fae5;border-radius:3px;border:1px solid #a7f3d0;"></span>
                Amélioration</span>
            <span style="display:flex;align-items:center;gap:5px;">
                <span style="display:inline-block;width:14px;height:14px;background:#fee2e2;border-radius:3px;border:1px solid #fca5a5;"></span>
                Dégradation</span>
            <span style="display:flex;align-items:center;gap:5px;">
                <span style="display:inline-block;width:14px;height:14px;background:#f8fafc;border-radius:3px;border:1px solid #e2e8f0;"></span>
                Stable (écart ≤ 0.5 pt)</span>
        </div>""", unsafe_allow_html=True)


# ============================================================
# FIN NOUVEAU — Suite du code existant
# ============================================================

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
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5d0;border-radius:3px}
    
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
            go.Pie(labels=counts.index, values=counts.values, hole=0.4, sort=False,
                textinfo="percent", textposition="outside", pull=pull_list,
                marker=dict(colors=colors, line=dict(color="white", width=2))))
        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>Nombre : %{value}<br>Pourcentage : %{percent}<extra></extra>",
            textfont=dict(size=13, family='Inter, sans-serif'))
        fig.update_layout(
            title=dict(text=title, x=0.5, xanchor='center', font=dict(size=16)), 
            height=500, showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"),
            margin=dict(t=80, b=80, l=40, r=40))
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
        
        conf_pv=pd.pivot_table(df[df["Statut OT"].isin(["CLOT","TCLO"])],index="Poste travail princ.",columns="OT CONFIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: conf_pv[c]=conf_pv.get(c,0)
        conf_pv["Total"]=conf_pv["OUI"]+conf_pv["NON"]; conf_pv["OT CONFIME"]=ckpi(conf_pv["OUI"],conf_pv["Total"])
        res['ot_confime']=conf_pv
        
        cor_pv=pd.pivot_table(df[df["Statut OT"].isin(["CLOT","TCLO"])],index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: cor_pv[c]=cor_pv.get(c,0)
        cor_pv["Total"]=cor_pv["OUI"]+cor_pv["NON"]; cor_pv["OT_COR_EGAL"]=ckpi(cor_pv["OUI"],cor_pv["Total"])
        res['ot_cor_egal']=cor_pv
        
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
            "Performance Graissage":g_df["Performance Graissage"],"Performance Inspection":ins_df["Performance Inspection"],"Performance Appels Systématiques":sys_df["Performance Appels Systématiques"]
        })
        
        res['cqdf']=pd.DataFrame({
            "Taux d'approbation des Avis":tca["Taux d'approbation des Avis"],
            "OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],
            "Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":conf_pv["OT CONFIME"],
            "OT_COR_EGAL":cor_pv["OT_COR_EGAL"],
            "OT Fiabilité":fiab_s,
            "Total Avis de Panne":avpan_s
        })
        
        # Scores globaux
        res['score_perf']=res['ckdf'].mean(axis=1)
        res['score_qual']=res['cqdf'].mean(axis=1)
        
        return res

    # ── Sidebar ──
    with st.sidebar:
        logo_b64=get_logo_base64()
        if logo_b64:
            st.markdown(f'<img src="data:image/png;base64,{logo_b64}" class="logo" style="width:100%;max-width:180px;margin:0 auto 10px;display:block;">',unsafe_allow_html=True)
        st.markdown("### 📁 Fichiers sources")
        ot_file=st.file_uploader("Fichier OT (.xlsx)",type=["xlsx","xls"],key="ot_up")
        av_file=st.file_uploader("Fichier Avis (.xlsx)",type=["xlsx","xls"],key="av_up")
        st.markdown("---")
        poste_filter=st.multiselect("🏭 Postes",options=[],key="poste_filt")
        st.markdown("---")
        st.markdown('<div style="color:rgba(255,255,255,.5);font-size:11px;text-align:center;padding:10px 0;">Dashboard KPI v2.0<br>Suivi & Évaluation</div>',unsafe_allow_html=True)

    # ── Chargement des données ──
    if not ot_file or not av_file:
        st.markdown("""<div style="min-height:60vh;display:flex;flex-direction:column;align-items:center;justify-content:center;">
            <div style="font-size:80px;margin-bottom:20px;">📂</div>
            <h2 style="color:#1e3a5f;font-weight:800;">Chargement des fichiers</h2>
            <p style="color:#64748b;font-size:16px;margin-top:8px;">Veuillez charger les fichiers OT et Avis depuis la barre latérale.</p>
        </div>""",unsafe_allow_html=True)
        return

    try:
        df, avf, apm, now_ts = prepare_data(ot_file.read(), av_file.read(), fichier_date)
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return

    if not apm:
        st.markdown('<div class="es" style="padding:50px;">Aucun poste de travail détecté dans les données.</div>',unsafe_allow_html=True)
        return

    # Filtre postes
    filtered_posts = poste_filter if poste_filter else apm

    # Calcul des KPIs
    kpi_data = calc_kpis(df, avf, now_ts, filtered_posts)
    ckdf = kpi_data['ckdf']
    cqdf = kpi_data['cqdf']
    score_perf = kpi_data['score_perf']
    score_qual = kpi_data['score_qual']

    # ── Header ──
    logo_html = ''
    logo_b64=get_logo_base64()
    if logo_b64:
        logo_html=f'<img src="data:image/png;base64,{logo_b64}" class="logo">'
    st.markdown(f"""<div class="mh">{logo_html}<h1>Dashboard KPI Maintenance</h1><div class="db">{fichier_date}</div></div>""",unsafe_allow_html=True)

    # ── Cartes résumé ──
    total_ot = len(df)
    clotures = len(df[df["Statut OT"].isin(["CLOT","TCLO"])])
    taux_gen = (clotures/total_ot*100) if total_ot>0 else 100
    avg_p = score_perf.mean()
    avg_q = score_qual.mean()

    st.markdown(f"""<div class="cr">
        <div class="cc c1"><div class="cv">{total_ot}</div><div class="cl">Total OT</div></div>
        <div class="cc c2"><div class="cv">{taux_gen:.1f}%</div><div class="cl">Taux Réalisation</div></div>
        <div class="cc c3"><div class="cv">{avg_p:.1f}%</div><div class="cl">Score Performance</div></div>
        <div class="cc c4"><div class="cv">{avg_q:.1f}%</div><div class="cl">Score Qualité</div></div>
    </div>""",unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════
    # ONGLETS — AJOUT DE "Suivi & Évaluation"
    # ═══════════════════════════════════════════════════════════
    tab_synthese, tab_suivi_eval, tab_suivi_kpi, tab_plan = st.tabs([
        "📊 Synthèse", "📈 Suivi & Évaluation", "🔍 Suivi KPI", "📝 Plan d'Action"
    ])

    # ── Onglet Synthèse (votre code existant) ──
    with tab_synthese:
        st.markdown('<div class="stl">📊 Synthèse des Indicateurs</div>',unsafe_allow_html=True)

        # Performance
        st.markdown('<div class="stl" style="color:#059669;border-left-color:#059669;">Indicateurs de Performance</div>',unsafe_allow_html=True)
        pcols=["Poste de travail"]+QK+["Score Performance"]
        prows=[]
        for poste in filtered_posts:
            row={"Poste de travail":poste}
            for kpi in QK: row[kpi]=round(ckdf.loc[poste,kpi],1) if poste in ckdf.index else 0
            row["Score Performance"]=round(score_perf.loc[poste],1) if poste in score_perf.index else 0
            prows.append(row)
        # Ligne cible
        cible_row={"Poste de travail":"Cible"}
        for kpi in QK: cible_row[kpi]=CIBLE.get(kpi,"")
        cible_row["Score Performance"]=""
        prows.append(cible_row)

        # Détection anomalies performance
        ano_p_r=[]; ano_p_c=["Poste","KPI","Valeur","Cible","Ecart","Action"]
        for poste in filtered_posts:
            for kpi in QK:
                val=ckdf.loc[poste,kpi] if poste in ckdf.index else 0
                cible=CIBLE.get(kpi,100)
                is_lb=kpi in LOWER_BETTER
                if (not is_lb and val<cible) or (is_lb and val>cible):
                    ano_p_r.append({"Poste":poste,"KPI":kpi,"Valeur":round(val,1),"Cible":cible,
                        "Ecart":round(val-cible,1),"Action":ACT_MAP.get(kpi,"")})

        # Qualité
        st.markdown('<div class="stl" style="color:#2563eb;border-left-color:#2563eb;">Indicateurs de Qualité</div>',unsafe_allow_html=True)
        qcols=["Poste de travail"]+PK+["Score Qualite"]
        qrows=[]
        for poste in filtered_posts:
            row={"Poste de travail":poste}
            for kpi in PK: row[kpi]=round(cqdf.loc[poste,kpi],1) if poste in cqdf.index else 0
            row["Score Qualite"]=round(score_qual.loc[poste],1) if poste in score_qual.index else 0
            qrows.append(row)
        cible_row_q={"Poste de travail":"Cible"}
        for kpi in PK: cible_row_q[kpi]=CIBLE.get(kpi,"")
        cible_row_q["Score Qualite"]=""
        qrows.append(cible_row_q)

        ano_q_r=[]; ano_q_c=["Poste","KPI","Valeur","Cible","Ecart","Action"]
        for poste in filtered_posts:
            for kpi in PK:
                val=cqdf.loc[poste,kpi] if poste in cqdf.index else 0
                cible=CIBLE.get(kpi,100)
                is_lb=kpi in LOWER_BETTER
                if (not is_lb and val<cible) or (is_lb and val>cible):
                    ano_q_r.append({"Poste":poste,"KPI":kpi,"Valeur":round(val,1),"Cible":cible,
                        "Ecart":round(val-cible,1),"Action":ACT_MAP.get(kpi,"")})

        # Bouton sauvegarde
        if st.button("💾 Sauvegarder les KPIs",key="save_kpis",use_container_width=True):
            save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,fichier_date)
            st.success("KPIs sauvegardés avec succès !")

        # Tableaux HTML
        def render_kpi_table(rows, cols, table_class):
            h='<table class="tw %s"><thead><tr>'%table_class+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
            for i,r in enumerate(rows):
                is_cible = r.get("Poste de travail")=="Cible"
                cls=' class="cb"' if is_cible else ''
                h+=f'<tr{cls}>'
                for c in cols:
                    v=r.get(c,"")
                    if isinstance(v,float): v=f"{v:.1f}%"
                    h+=f'<td>{v}</td>'
                h+='</tr>'
            h+='</tbody></table>'
            return h

        st.markdown(render_kpi_table(prows,pcols,"pt"),unsafe_allow_html=True)
        st.markdown(render_kpi_table(qrows,qcols,"qt"),unsafe_allow_html=True)

        # Anomalies
        if ano_p_r:
            st.markdown('<div class="stl" style="color:#ef4444;border-left-color:#ef4444;">⚠️ Anomalies Performance</div>',unsafe_allow_html=True)
            ah='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in ano_p_c)+'</tr></thead><tbody>'
            for r in ano_p_r:
                ah+=f'<tr><td>{r["Poste"]}</td><td>{r["KPI"]}</td><td>{r["Valeur"]}%</td><td>{r["Cible"]}%</td><td>{r["Ecart"]:+.1f}</td><td style="text-align:left;max-width:300px;">{r["Action"]}</td></tr>'
            ah+='</tbody></table>'
            st.markdown(ah,unsafe_allow_html=True)

        if ano_q_r:
            st.markdown('<div class="stl" style="color:#ef4444;border-left-color:#ef4444;">⚠️ Anomalies Qualité</div>',unsafe_allow_html=True)
            ah='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in ano_q_c)+'</tr></thead><tbody>'
            for r in ano_q_r:
                ah+=f'<tr><td>{r["Poste"]}</td><td>{r["KPI"]}</td><td>{r["Valeur"]}%</td><td>{r["Cible"]}%</td><td>{r["Ecart"]:+.1f}</td><td style="text-align:left;max-width:300px;">{r["Action"]}</td></tr>'
            ah+='</tbody></table>'
            st.markdown(ah,unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════
    # NOUVEAU : Onglet Suivi & Évaluation
    # ═══════════════════════════════════════════════════════════
    with tab_suivi_eval:
        kpis_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
        hist_df_suivi = load_historical_kpis(kpis_path)
        var_df_suivi = calculate_variations(hist_df_suivi)
        page_suivi_evaluation(filtered_posts, hist_df_suivi, var_df_suivi)

    # ── Onglet Suivi KPI (votre code existant) ──
    with tab_suivi_kpi:
        st.markdown('<div class="stl">🔍 Suivi Détaillé des KPIs par Poste</div>',unsafe_allow_html=True)
        sel_poste = st.selectbox("Sélectionner un poste", filtered_posts, key="kpi_poste_sel")
        if sel_poste:
            st.markdown(f'<div class="stl" style="color:#059669;border-left-color:#059669;">Performance — {sel_poste}</div>',unsafe_allow_html=True)
            p_vals = ckdf.loc[sel_poste] if sel_poste in ckdf.index else pd.Series()
            for kpi in QK:
                val = p_vals.get(kpi,0)
                cible = CIBLE.get(kpi,100)
                pct = min(val,100)
                clr = "#059669" if (kpi not in LOWER_BETTER and val>=cible) or (kpi in LOWER_BETTER and val<=cible) else "#ef4444"
                st.markdown(f"""<div class="ca"><div class="ct">{kpi}</div>
                    <div class="car">
                        <div class="cal">{kpi}</div>
                        <div class="cab"><div class="caf" style="width:{pct}%;background:{clr};"></div>
                            <div class="target-mark" style="left:{cible}%;"></div></div>
                        <div class="cav-out">{val:.1f}%</div>
                        <div class="cav-tgt">Cible: {cible}%</div>
                    </div></div>""",unsafe_allow_html=True)

            st.markdown(f'<div class="stl" style="color:#2563eb;border-left-color:#2563eb;">Qualité — {sel_poste}</div>',unsafe_allow_html=True)
            q_vals = cqdf.loc[sel_poste] if sel_poste in cqdf.index else pd.Series()
            for kpi in PK:
                val = q_vals.get(kpi,0)
                cible = CIBLE.get(kpi,100)
                pct = min(val,100)
                clr = "#2563eb" if val>=cible else "#ef4444"
                st.markdown(f"""<div class="ca"><div class="ct">{kpi}</div>
                    <div class="car">
                        <div class="cal">{kpi}</div>
                        <div class="cab"><div class="caf" style="width:{pct}%;background:{clr};"></div>
                            <div class="target-mark" style="left:{cible}%;"></div></div>
                        <div class="cav-out">{val:.1f}%</div>
                        <div class="cav-tgt">Cible: {cible}%</div>
                    </div></div>""",unsafe_allow_html=True)

    # ── Onglet Plan d'Action (votre code existant) ──
    with tab_plan:
        st.markdown('<div class="stl">📝 Plan d\'Action</div>',unsafe_allow_html=True)
        all_anomalies = ano_p_r + ano_q_r if 'ano_p_r' in dir() else []
        if all_anomalies:
            pa_html='<table class="plan-action-table"><thead><tr><th>Poste</th><th>KPI</th><th>Valeur</th><th>Cible</th><th>Écart</th><th>Responsable</th><th>Action Corrective</th><th>Statut</th></tr></thead><tbody>'
            for r in all_anomalies:
                resp = KPI_RESP_MAP.get(r["KPI"],"—")
                pa_html+=f'<tr><td>{r["Poste"]}</td><td>{r["KPI"]}</td><td>{r["Valeur"]}%</td><td>{r["Cible"]}%</td><td>{r["Ecart"]:+.1f}</td><td>{resp}</td><td style="text-align:left;max-width:250px;">{r["Action"]}</td><td>⏳ En cours</td></tr>'
            pa_html+='</tbody></table>'
            st.markdown(pa_html,unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Aucune anomalie détectée. Tous les indicateurs sont conformes aux cibles.</div>',unsafe_allow_html=True)

    # Footer
    st.markdown('<div class="footer">Dashboard KPI Maintenance — Suivi & Évaluation — ' + fichier_date + '</div>',unsafe_allow_html=True)

if __name__ == "__main__":
    main()
