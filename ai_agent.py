# -*- coding: utf-8 -*-
"""
AI Agent pour l'analyse des KPI Maintenance
Génère une structure JSON pour présentation PowerPoint
"""

import json
from openai import OpenAI
from typing import Dict, List, Any, Optional
import pandas as pd
import numpy as np

# ============================================================
# Configuration du client OpenAI
# ============================================================
client = OpenAI(
    api_key="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",
    base_url="https://openrouter.ai/api/v1"
)

# ============================================================
# Prompt Système
# ============================================================
SYSTEM_PROMPT = """Tu es un ingénieur maintenance senior avec plus de 20 ans d'expérience en maintenance industrielle, fiabilité, gestion des actifs et amélioration continue.

Tu travailles dans une usine chimique.

Tu analyses un tableau de bord SAP PM contenant des KPI de performance et de qualité de la maintenance.

Tu dois préparer une présentation PowerPoint destinée au Directeur Maintenance.

Les données proviennent des filtres sélectionnés par l'utilisateur.

Division : {division}

Atelier : {atelier}

Métier : {metier}

Période :

Du {date_debut}

Au {date_fin}

Tu recevras les éléments suivants :

* Les KPI Performance
* Les KPI Qualité
* Les tableaux d'anomalies
* Les tendances d'évolution
* Les scores globaux
* Les variations par rapport aux périodes précédentes
* Les plans d'action existants

Ta mission est de réaliser une analyse professionnelle.

Règles :

* Ne jamais inventer de données.
* Utiliser uniquement les informations fournies.
* Donner des conclusions courtes et professionnelles.
* Ne jamais dépasser une phrase de conclusion par slide.
* Identifier les KPI les plus critiques.
* Expliquer les causes probables.
* Prioriser les actions selon leur impact.
* Utiliser un vocabulaire de maintenance industrielle.

Retourner uniquement un JSON valide.

Structure attendue :

{{
"title":"",
"subtitle":"",
"executive_summary":"",

"slide1":{{
"conclusion":""
}},

"slide2":{{
"conclusion":""
}},

"slide3":{{
"conclusion":""
}},

"slide4":{{
"root_causes":[]
}},

"slide5":{{
"recommendations":[]
}},

"slide6":{{
"final_conclusion":""
}}

}}

Contenu attendu :

Slide 0
Titre :
Analyse des KPI Performance et Qualité

Sous-titre :
Division
Atelier
Métier
Période

Slide 1
Analyse des KPI Performance et Qualité.

La conclusion doit résumer l'état global des performances.

Slide 2
Analyse du tableau des anomalies.

Identifier les postes de travail présentant le plus grand nombre d'anomalies.

Slide 3
Analyse des tendances.

Comparer les performances avec les périodes précédentes.

Identifier les KPI en amélioration ou en dégradation.

Slide 4
Identifier les causes racines les plus probables.

Classer les causes par ordre de criticité.

Slide 5
Élaborer un plan d'action.

Pour chaque action fournir :

* Action
* Responsable
* Priorité
* Échéance
* Impact attendu

Slide 6
Produire un résumé exécutif destiné au Directeur Maintenance.

Présenter :

* Les principaux constats.
* Les risques.
* Les opportunités d'amélioration.
* La priorité stratégique.
* Une conclusion finale.

Les réponses doivent être professionnelles, précises et directement exploitables dans une réunion de performance maintenance."""


# ============================================================
# Fonctions de formatage des données
# ============================================================

def format_kpi_table(df, section_name: str) -> str:
    """Formate un DataFrame KPI en texte lisible pour le prompt."""
    if df is None or df.empty:
        return f"Aucune donnée {section_name} disponible."
    
    lines = [f"=== {section_name.upper()} ==="]
    df_str = df.copy().astype(str)
    
    # En-tête
    header = " | ".join(df_str.columns.tolist())
    lines.append(header)
    lines.append("-" * min(len(header), 120))
    
    # Données
    for idx, row in df_str.iterrows():
        line = " | ".join(row.tolist())
        lines.append(line)
    
    return "\n".join(lines)


def format_anomalies_table(df, section_name: str) -> str:
    """Formate un tableau d'anomalies en texte lisible."""
    if df is None or df.empty:
        return f"Aucune anomalie {section_name} identifiée."
    
    lines = [f"=== ANOMALIES {section_name.upper()} ==="]
    df_str = df.copy().astype(str)
    
    header = " | ".join(df_str.columns.tolist())
    lines.append(header)
    lines.append("-" * min(len(header), 120))
    
    for idx, row in df_str.iterrows():
        line = " | ".join(row.tolist())
        lines.append(line)
    
    total = len(df_str)
    lines.append(f"\nTotal anomalies {section_name}: {total}")
    
    return "\n".join(lines)


def format_trends(var_df) -> str:
    """Formate les données de tendances pour le prompt."""
    if var_df is None or var_df.empty:
        return "Aucune donnée de tendance disponible (pas d'historique suffisant)."
    
    lines = ["=== TENDANCES D'ÉVOLUTION ==="]
    
    # Filtrer les variations significatives si la colonne existe
    if "Significatif" in var_df.columns:
        significant = var_df[var_df["Significatif"] == True]
    else:
        # Sinon filtrer sur l'écart %
        if "Ecart %" in var_df.columns:
            significant = var_df[var_df["Ecart %"].abs() >= 5]
        else:
            significant = var_df
    
    if significant.empty:
        lines.append("Aucune variation significative (>=5%) détectée sur la période.")
        return "\n".join(lines)
    
    # Grouper par sens d'évolution
    for sens in ["Degradation", "Amelioration", "Stable"]:
        if "Sens" in significant.columns:
            subset = significant[significant["Sens"] == sens]
        else:
            continue
            
        if subset.empty:
            continue
        
        lines.append(f"\n--- {sens.upper()} ---")
        count = 0
        for _, row in subset.iterrows():
            if count >= 20:  # Limiter à 20 lignes par catégorie
                lines.append(f"... et {len(subset) - 20} autres variations")
                break
            poste = row.get("Poste", "N/A")
            kpi = row.get("KPI", "N/A")
            prev = row.get("Valeur precedente", "N/A")
            curr = row.get("Valeur actuelle", "N/A")
            ecart_pct = row.get("Ecart %", 0)
            try:
                ecart_str = f"{float(ecart_pct):+.1f}%"
            except (ValueError, TypeError):
                ecart_str = "N/A"
            lines.append(f"• {poste} | {kpi}: {prev} → {curr} ({ecart_str})")
            count += 1
    
    return "\n".join(lines)


def format_rankings(best_df, worst_df) -> str:
    """Formate les classements des postes."""
    lines = ["=== CLASSEMENT DES POSTES ==="]
    
    if best_df is not None and not best_df.empty:
        lines.append("\n--- TOP 5 POSTES EN AMÉLIORATION ---")
        for _, row in best_df.head(5).iterrows():
            score = row.get("Score variation", 0)
            try:
                score_str = f"{float(score):+.1f} pts"
            except (ValueError, TypeError):
                score_str = "N/A"
            lines.append(f"• {row.get('Poste', 'N/A')}: {score_str}")
    
    if worst_df is not None and not worst_df.empty:
        lines.append("\n--- TOP 5 POSTES EN DÉGRADATION ---")
        for _, row in worst_df.head(5).iterrows():
            score = row.get("Score variation", 0)
            try:
                score_str = f"{float(score):+.1f} pts"
            except (ValueError, TypeError):
                score_str = "N/A"
            lines.append(f"• {row.get('Poste', 'N/A')}: {score_str}")
    
    if len(lines) == 1:
        lines.append("Aucun classement disponible.")
    
    return "\n".join(lines)


def format_existing_actions(actions_list) -> str:
    """Formate la liste des actions existantes."""
    if not actions_list:
        return "Aucun plan d'action existant."
    
    lines = ["=== PLANS D'ACTION EXISTANTS ==="]
    for i, action in enumerate(actions_list[:15], 1):  # Limiter à 15 actions
        if isinstance(action, dict):
            action_text = action.get("action", action.get("Action", "N/A"))
            responsable = action.get("responsable", action.get("Responsable", "N/A"))
            priorite = action.get("priorite", action.get("Priorité", "N/A"))
            lines.append(f"{i}. {action_text} | Responsable: {responsable} | Priorité: {priorite}")
        elif isinstance(action, str):
            lines.append(f"{i}. {action}")
    
    return "\n".join(lines)


# ============================================================
# Construction du prompt utilisateur
# ============================================================

def build_user_prompt(
    division: str,
    atelier: str,
    metier: str,
    date_debut: str,
    date_fin: str,
    kpi_performance_df,
    kpi_qualite_df,
    anomalies_perf_df=None,
    anomalies_qual_df=None,
    variations_df=None,
    best_ranking_df=None,
    worst_ranking_df=None,
    score_performance_global: float = 0.0,
    score_qualite_global: float = 0.0,
    total_ot_analyses: int = 0,
    total_anomalies: int = 0,
    existing_actions: Optional[List[Dict]] = None
) -> str:
    """Construit le prompt utilisateur avec toutes les données formatées."""
    
    prompt_parts = [
        "DONNÉES À ANALYSER :",
        "",
        f"DIVISION: {division}",
        f"ATELIER: {atelier}",
        f"MÉTIER: {metier}",
        f"PÉRIODE: Du {date_debut} au {date_fin}",
        "",
        "=" * 60,
        "SCORES GLOBAUX",
        "=" * 60,
        f"Score Performance Global: {score_performance_global:.1f}%",
        f"Score Qualité Global: {score_qualite_global:.1f}%",
        f"Total OT Analysés: {total_ot_analyses}",
        f"Total Anomalies Détectées: {total_anomalies}",
        "",
        format_kpi_table(kpi_performance_df, "KPI Performance"),
        "",
        format_kpi_table(kpi_qualite_df, "KPI Qualité"),
        "",
        format_anomalies_table(anomalies_perf_df, "Performance"),
        "",
        format_anomalies_table(anomalies_qual_df, "Qualité"),
        "",
        format_trends(variations_df),
        "",
        format_rankings(best_ranking_df, worst_ranking_df),
        "",
        format_existing_actions(existing_actions),
        "",
        "=" * 60,
        "INSTRUCTIONS FINALES",
        "=" * 60,
        "Analyser ces données et retourner UNIQUEMENT le JSON valide avec la structure demandée.",
        "Les conclusions doivent être courtes (une phrase max par slide).",
        "Utiliser le vocabulaire maintenance industrielle SAP PM.",
        "Ne pas inventer de données non fournies."
    ]
    
    return "\n".join(prompt_parts)


# ============================================================
# Fonction principale de génération
# ============================================================

def generate_analysis(
    division: str,
    atelier: str,
    metier: str,
    date_debut: str,
    date_fin: str,
    kpi_performance_df,
    kpi_qualite_df,
    anomalies_perf_df=None,
    anomalies_qual_df=None,
    variations_df=None,
    best_ranking_df=None,
    worst_ranking_df=None,
    score_performance_global: float = 0.0,
    score_qualite_global: float = 0.0,
    total_ot_analyses: int = 0,
    total_anomalies: int = 0,
    existing_actions: Optional[List[Dict]] = None,
    model: str = "deepseek/deepseek-chat"
) -> Dict[str, Any]:
    """
    Génère l'analyse AI des KPI maintenance.
    
    Args:
        division: Nom de la division
        atelier: Nom de l'atelier
        metier: Métier concerné
        date_debut: Date de début de la période (format dd/mm/yyyy)
        date_fin: Date de fin de la période (format dd/mm/yyyy)
        kpi_performance_df: DataFrame des KPI Performance
        kpi_qualite_df: DataFrame des KPI Qualité
        anomalies_perf_df: DataFrame des anomalies Performance (optionnel)
        anomalies_qual_df: DataFrame des anomalies Qualité (optionnel)
        variations_df: DataFrame des variations/tendances (optionnel)
        best_ranking_df: DataFrame du classement meilleur (optionnel)
        worst_ranking_df: DataFrame du classement pire (optionnel)
        score_performance_global: Score performance global en %
        score_qualite_global: Score qualité global en %
        total_ot_analyses: Nombre total d'OT analysés
        total_anomalies: Nombre total d'anomalies
        existing_actions: Liste des actions existantes (optionnel)
        model: Modèle AI à utiliser (défaut: deepseek/deepseek-chat)
    
    Returns:
        Dictionnaire JSON avec la structure attendue pour la présentation
    """
    
    # Préparer le système prompt avec les variables
    system_prompt = SYSTEM_PROMPT.format(
        division=division,
        atelier=atelier,
        metier=metier,
        date_debut=date_debut,
        date_fin=date_fin
    )
    
    # Construire le prompt utilisateur
    user_prompt = build_user_prompt(
        division=division,
        atelier=atelier,
        metier=metier,
        date_debut=date_debut,
        date_fin=date_fin,
        kpi_performance_df=kpi_performance_df,
        kpi_qualite_df=kpi_qualite_df,
        anomalies_perf_df=anomalies_perf_df,
        anomalies_qual_df=anomalies_qual_df,
        variations_df=variations_df,
        best_ranking_df=best_ranking_df,
        worst_ranking_df=worst_ranking_df,
        score_performance_global=score_performance_global,
        score_qualite_global=score_qualite_global,
        total_ot_analyses=total_ot_analyses,
        total_anomalies=total_anomalies,
        existing_actions=existing_actions
    )
    
    try:
        # Appel à l'API OpenAI
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,  # Température basse pour des réponses cohérentes et précises
            max_tokens=4000,
            response_format={"type": "json_object"}
        )
        
        # Extraire le contenu de la réponse
        content = response.choices[0].message.content
        
        # Nettoyer le contenu si nécessaire (retirer les blocs de code markdown)
        if content.startswith("```json"):
            content = content[7:]
        elif content.startswith("```"):
            content = content[3:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()
        
        # Parser le JSON
        result = json.loads(content)
        
        # Valider et compléter la structure
        result = validate_and_complete_structure(
            result, division, atelier, metier, date_debut, date_fin
        )
        
        return result
        
    except json.JSONDecodeError as e:
        return get_error_structure(f"Erreur de parsing JSON: {str(e)}")
    except Exception as e:
        return get_error_structure(f"Erreur lors de l'appel API: {str(e)}")


# ============================================================
# Validation et complétion de la structure
# ============================================================

def validate_and_complete_structure(
    result: Dict, 
    division: str, 
    atelier: str, 
    metier: str, 
    date_debut: str, 
    date_fin: str
) -> Dict:
    """Valide et complète la structure JSON retournée par l'IA."""
    
    # Structure minimale attendue
    default_structure = {
        "title": "Analyse des KPI Performance et Qualité",
        "subtitle": f"{division} | {atelier} | {metier} | Du {date_debut} au {date_fin}",
        "executive_summary": "",
        "slide1": {"conclusion": ""},
        "slide2": {"conclusion": ""},
        "slide3": {"conclusion": ""},
        "slide4": {"root_causes": []},
        "slide5": {"recommendations": []},
        "slide6": {"final_conclusion": ""}
    }
    
    # Fusionner avec le résultat (le résultat écrase les défauts)
    for key in default_structure:
        if key not in result:
            result[key] = default_structure[key]
        elif isinstance(default_structure[key], dict) and isinstance(result[key], dict):
            for subkey in default_structure[key]:
                if subkey not in result[key]:
                    result[key][subkey] = default_structure[key][subkey]
    
    # S'assurer que le titre et sous-titre sont corrects
    result["title"] = default_structure["title"]
    result["subtitle"] = default_structure["subtitle"]
    
    # S'assurer que les listes sont bien des listes
    if not isinstance(result.get("slide4", {}).get("root_causes"), list):
        result["slide4"]["root_causes"] = []
    if not isinstance(result.get("slide5", {}).get("recommendations"), list):
        result["slide5"]["recommendations"] = []
    
    # Valider et normaliser les causes racines
    valid_criticites = ["Critique", "Élevée", "Moyenne", "Faible"]
    valid_causes = []
    for cause in result["slide4"]["root_causes"]:
        if isinstance(cause, dict):
            validated_cause = {
                "cause": str(cause.get("cause", cause.get("Cause", "")))[:300],
                "criticite": cause.get("criticite", cause.get("Criticité", "Moyenne"))
            }
            # Normaliser la criticité
            crit = validated_cause["criticite"]
            if crit not in valid_criticites:
                validated_cause["criticite"] = "Moyenne"
            valid_causes.append(validated_cause)
        elif isinstance(cause, str):
            valid_causes.append({"cause": cause[:300], "criticite": "Moyenne"})
    result["slide4"]["root_causes"] = valid_causes
    
    # Valider et normaliser les recommandations
    valid_priorities = ["Haute", "Moyenne", "Basse"]
    valid_rec = []
    for rec in result["slide5"]["recommendations"]:
        if isinstance(rec, dict):
            # Gérer les différentes clés possibles (avec ou sans accents)
            action = rec.get("action", rec.get("Action", ""))
            responsable = rec.get("responsable", rec.get("Responsable", ""))
            priorite = rec.get("priorite", rec.get("Priorité", rec.get("Priorite", "Moyenne")))
            echeance = rec.get("echeance", rec.get("Échéance", rec.get("Echeance", "")))
            impact = rec.get("impact_attendu", rec.get("Impact attendu", rec.get("impact", "")))
            
            validated_rec = {
                "action": str(action)[:250],
                "responsable": str(responsable)[:60],
                "priorite": priorite if priorite in valid_priorities else "Moyenne",
                "echeance": str(echeance)[:60],
                "impact_attendu": str(impact)[:200]
            }
            valid_rec.append(validated_rec)
        elif isinstance(rec, str):
            valid_rec.append({
                "action": rec[:250],
                "responsable": "À définir",
                "priorite": "Moyenne",
                "echeance": "À définir",
                "impact_attendu": ""
            })
    result["slide5"]["recommendations"] = valid_rec
    
    return result


def get_error_structure(error_message: str) -> Dict:
    """Retourne une structure d'erreur standardisée."""
    return {
        "title": "Analyse des KPI Performance et Qualité",
        "subtitle": "Erreur lors de l'analyse",
        "executive_summary": "",
        "slide1": {"conclusion": f"Erreur lors de l'analyse: {error_message}"},
        "slide2": {"conclusion": "Données non disponibles en raison d'une erreur technique."},
        "slide3": {"conclusion": "Données non disponibles en raison d'une erreur technique."},
        "slide4": {"root_causes": [{"cause": f"Erreur technique: {error_message}", "criticite": "Critique"}]},
        "slide5": {"recommendations": []},
        "slide6": {"final_conclusion": "Veuillez réessayer ultérieurement ou vérifier les données d'entrée."}
    }


# ============================================================
# Fonction d'intégration pour Streamlit
# ============================================================

def generate_analysis_from_dashboard(
    st,
    division: str,
    atelier: str,
    metier: str,
    date_debut,
    date_fin,
    kpi_perf_rows: List[Dict],
    kpi_perf_cols: List[str],
    kpi_qual_rows: List[Dict],
    kpi_qual_cols: List[str],
    ano_perf_rows: List[Dict],
    ano_perf_cols: List[str],
    ano_qual_rows: List[Dict],
    ano_qual_cols: List[str],
    var_df,
    best_ranking_df,
    worst_ranking_df,
    score_perf: float,
    score_qual: float,
    total_ot: int,
    total_anomalies: int,
    plan_action_df=None,
    model: str = "deepseek/deepseek-chat",
    show_progress: bool = True
) -> Dict[str, Any]:
    """
    Fonction wrapper pour générer l'analyse depuis le dashboard Streamlit.
    
    Args:
        st: Instance Streamlit
        division, atelier, metier: Filtres sélectionnés
        date_debut, date_fin: Dates de la période
        kpi_perf_rows, kpi_perf_cols: Données KPI Performance
        kpi_qual_rows, kpi_qual_cols: Données KPI Qualité
        ano_perf_rows, ano_perf_cols: Anomalies Performance
        ano_qual_rows, ano_qual_cols: Anomalies Qualité
        var_df: DataFrame des variations
        best_ranking_df, worst_ranking_df: Classements
        score_perf, score_qual: Scores globaux
        total_ot, total_anomalies: Totaux
        plan_action_df: DataFrame du plan d'action (optionnel)
        model: Modèle AI à utiliser
        show_progress: Afficher la barre de progression
    
    Returns:
        Dictionnaire JSON avec la structure attendue
    """
    
    # Convertir les dates en string
    if hasattr(date_debut, 'strftime'):
        date_debut_str = date_debut.strftime("%d/%m/%Y")
    else:
        date_debut_str = str(date_debut)
    
    if hasattr(date_fin, 'strftime'):
        date_fin_str = date_fin.strftime("%d/%m/%Y")
    else:
        date_fin_str = str(date_fin)
    
    # Convertir les listes de dictionnaires en DataFrame
    kpi_perf_df = pd.DataFrame(kpi_perf_rows, columns=kpi_perf_cols) if kpi_perf_rows else pd.DataFrame()
    kpi_qual_df = pd.DataFrame(kpi_qual_rows, columns=kpi_qual_cols) if kpi_qual_rows else pd.DataFrame()
    ano_perf_df = pd.DataFrame(ano_perf_rows, columns=ano_perf_cols) if ano_perf_rows else pd.DataFrame()
    ano_qual_df = pd.DataFrame(ano_qual_rows, columns=ano_qual_cols) if ano_qual_rows else pd.DataFrame()
    
    # Extraire les actions existantes du plan d'action
    existing_actions = None
    if plan_action_df is not None and not plan_action_df.empty:
        existing_actions = plan_action_df.to_dict('records')
    
    # Afficher la progression
    if show_progress:
        progress_container = st.empty()
        progress_container.markdown("""
        <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#1e3a5f,#2563eb);
                    border-radius:12px;color:white;margin:10px 0">
            <div style="font-size:24px;margin-bottom:10px">🤖 Analyse IA en cours...</div>
            <div style="font-size:14px;opacity:0.8">Génération de l'analyse des KPI maintenance</div>
        </div>
        """, unsafe_allow_html=True)
    
    try:
        # Générer l'analyse
        result = generate_analysis(
            division=division,
            atelier=atelier,
            metier=metier,
            date_debut=date_debut_str,
            date_fin=date_fin_str,
            kpi_performance_df=kpi_perf_df,
            kpi_qualite_df=kpi_qual_df,
            anomalies_perf_df=ano_perf_df if not ano_perf_df.empty else None,
            anomalies_qual_df=ano_qual_df if not ano_qual_df.empty else None,
            variations_df=var_df,
            best_ranking_df=best_ranking_df,
            worst_ranking_df=worst_ranking_df,
            score_performance_global=score_perf,
            score_qualite_global=score_qual,
            total_ot_analyses=total_ot,
            total_anomalies=total_anomalies,
            existing_actions=existing_actions,
            model=model
        )
        
        if show_progress:
            progress_container.markdown("""
            <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#059669,#10b981);
                        border-radius:12px;color:white;margin:10px 0">
                <div style="font-size:24px;margin-bottom:10px">✅ Analyse terminée avec succès</div>
                <div style="font-size:14px;opacity:0.8">Les résultats sont prêts pour la présentation</div>
            </div>
            """, unsafe_allow_html=True)
        
        return result
        
    except Exception as e:
        if show_progress:
            progress_container.markdown(f"""
            <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#dc2626,#ef4444);
                        border-radius:12px;color:white;margin:10px 0">
                <div style="font-size:24px;margin-bottom:10px">❌ Erreur lors de l'analyse</div>
                <div style="font-size:14px;opacity:0.8">{str(e)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        return get_error_structure(str(e))


# ============================================================
# Test unitaire
# ============================================================

if __name__ == "__main__":
    print("=" * 60)
    print("TEST UNITAIRE - AI Agent KPI Maintenance")
    print("=" * 60)
    
    # Données de test
    test_perf_df = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-002", "SF2-001", "Total general"],
        "TAUX_REALISATION_CORRECTIF/PT": [72.5, 68.0, 92.3, 77.6],
        "OT préparation <1 mois": [55.0, 48.0, 88.0, 63.7],
        "OT préparation >3 mois": [18.0, 22.0, 3.0, 14.3],
        "OT planification <1 mois": [62.0, 58.0, 90.0, 70.0],
        "Score Performance": [65.0, 58.0, 91.0, 71.3]
    })
    
    test_qual_df = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-002", "SF2-001", "Total general"],
        "Taux d'approbation des Avis": [88.0, 82.0, 98.0, 89.3],
        "OT CONFIME": [75.0, 68.0, 100.0, 81.0],
        "Performance Graissage": [92.0, 85.0, 98.0, 91.7],
        "Score Qualite": [85.0, 78.3, 98.7, 87.3]
    })
    
    test_anomalies = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-001", "SF1-002", "SF1-002", "SF1-002"],
        "KPI": [
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "OT préparation >3 mois"
        ],
        "Valeur": [72.5, 55.0, 68.0, 48.0, 22.0],
        "Cible": [85.0, 80.0, 85.0, 80.0, 5.0],
        "Écart": [-12.5, -25.0, -17.0, -32.0, 17.0]
    })
    
    test_variations = pd.DataFrame({
        "Poste": ["SF1-001", "SF1-002", "SF2-001", "SF1-001", "SF1-002"],
        "KPI": [
            "TAUX_REALISATION_CORRECTIF/PT",
            "TAUX_REALISATION_CORRECTIF/PT",
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "OT préparation <1 mois"
        ],
        "Valeur precedente": [78.0, 75.0, 90.0, 65.0, 60.0],
        "Valeur actuelle": [72.5, 68.0, 92.3, 55.0, 48.0],
        "Ecart %": [-7.1, -9.3, 2.6, -15.4, -20.0],
        "Sens": ["Degradation", "Degradation", "Amelioration", "Degradation", "Degradation"],
        "Significatif": [True, True, False, True, True]
    })
    
    test_best = pd.DataFrame({
        "Poste": ["SF2-001", "SF2-002"],
        "Score variation": [8.5, 3.2]
    })
    
    test_worst = pd.DataFrame({
        "Poste": ["SF1-002", "SF1-001"],
        "Score variation": [-35.0, -22.0]
    })
    
    # Exécuter le test
    print("\nAppel à l'API IA...")
    result = generate_analysis(
        division="Division Chimique Nord",
        atelier="Atelier Production A",
        metier="Mécanique",
        date_debut="01/01/2024",
        date_fin="31/01/2024",
        kpi_performance_df=test_perf_df,
        kpi_qualite_df=test_qual_df,
        anomalies_perf_df=test_anomalies,
        anomalies_qual_df=pd.DataFrame(),  # Pas d'anomalies qualité dans le test
        variations_df=test_variations,
        best_ranking_df=test_best,
        worst_ranking_df=test_worst,
        score_performance_global=71.3,
        score_qualite_global=87.3,
        total_ot_analyses=156,
        total_anomalies=5
    )
    
    print("\n" + "=" * 60)
    print("RÉSULTAT JSON:")
    print("=" * 60)
    print(json.dumps(result, indent=2, ensure_ascii=False))# -*- coding: utf-8 -*-
"""
AI Agent pour l'analyse des KPI Maintenance
Génère une structure JSON pour présentation PowerPoint
"""

import json
from openai import OpenAI
from typing import Dict, List, Any, Optional
import pandas as pd
import numpy as np

# ============================================================
# Configuration du client OpenAI
# ============================================================
client = OpenAI(
    api_key="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",
    base_url="https://openrouter.ai/api/v1"
)

# ============================================================
# Prompt Système
# ============================================================
SYSTEM_PROMPT = """Tu es un ingénieur maintenance senior avec plus de 20 ans d'expérience en maintenance industrielle, fiabilité, gestion des actifs et amélioration continue.

Tu travailles dans une usine chimique.

Tu analyses un tableau de bord SAP PM contenant des KPI de performance et de qualité de la maintenance.

Tu dois préparer une présentation PowerPoint destinée au Directeur Maintenance.

Les données proviennent des filtres sélectionnés par l'utilisateur.

Division : {division}

Atelier : {atelier}

Métier : {metier}

Période :

Du {date_debut}

Au {date_fin}

Tu recevras les éléments suivants :

* Les KPI Performance
* Les KPI Qualité
* Les tableaux d'anomalies
* Les tendances d'évolution
* Les scores globaux
* Les variations par rapport aux périodes précédentes
* Les plans d'action existants

Ta mission est de réaliser une analyse professionnelle.

Règles :

* Ne jamais inventer de données.
* Utiliser uniquement les informations fournies.
* Donner des conclusions courtes et professionnelles.
* Ne jamais dépasser une phrase de conclusion par slide.
* Identifier les KPI les plus critiques.
* Expliquer les causes probables.
* Prioriser les actions selon leur impact.
* Utiliser un vocabulaire de maintenance industrielle.

Retourner uniquement un JSON valide.

Structure attendue :

{{
"title":"",
"subtitle":"",
"executive_summary":"",

"slide1":{{
"conclusion":""
}},

"slide2":{{
"conclusion":""
}},

"slide3":{{
"conclusion":""
}},

"slide4":{{
"root_causes":[]
}},

"slide5":{{
"recommendations":[]
}},

"slide6":{{
"final_conclusion":""
}}

}}

Contenu attendu :

Slide 0
Titre :
Analyse des KPI Performance et Qualité

Sous-titre :
Division
Atelier
Métier
Période

Slide 1
Analyse des KPI Performance et Qualité.

La conclusion doit résumer l'état global des performances.

Slide 2
Analyse du tableau des anomalies.

Identifier les postes de travail présentant le plus grand nombre d'anomalies.

Slide 3
Analyse des tendances.

Comparer les performances avec les périodes précédentes.

Identifier les KPI en amélioration ou en dégradation.

Slide 4
Identifier les causes racines les plus probables.

Classer les causes par ordre de criticité.

Slide 5
Élaborer un plan d'action.

Pour chaque action fournir :

* Action
* Responsable
* Priorité
* Échéance
* Impact attendu

Slide 6
Produire un résumé exécutif destiné au Directeur Maintenance.

Présenter :

* Les principaux constats.
* Les risques.
* Les opportunités d'amélioration.
* La priorité stratégique.
* Une conclusion finale.

Les réponses doivent être professionnelles, précises et directement exploitables dans une réunion de performance maintenance."""


# ============================================================
# Fonctions de formatage des données
# ============================================================

def format_kpi_table(df, section_name: str) -> str:
    """Formate un DataFrame KPI en texte lisible pour le prompt."""
    if df is None or df.empty:
        return f"Aucune donnée {section_name} disponible."
    
    lines = [f"=== {section_name.upper()} ==="]
    df_str = df.copy().astype(str)
    
    # En-tête
    header = " | ".join(df_str.columns.tolist())
    lines.append(header)
    lines.append("-" * min(len(header), 120))
    
    # Données
    for idx, row in df_str.iterrows():
        line = " | ".join(row.tolist())
        lines.append(line)
    
    return "\n".join(lines)


def format_anomalies_table(df, section_name: str) -> str:
    """Formate un tableau d'anomalies en texte lisible."""
    if df is None or df.empty:
        return f"Aucune anomalie {section_name} identifiée."
    
    lines = [f"=== ANOMALIES {section_name.upper()} ==="]
    df_str = df.copy().astype(str)
    
    header = " | ".join(df_str.columns.tolist())
    lines.append(header)
    lines.append("-" * min(len(header), 120))
    
    for idx, row in df_str.iterrows():
        line = " | ".join(row.tolist())
        lines.append(line)
    
    total = len(df_str)
    lines.append(f"\nTotal anomalies {section_name}: {total}")
    
    return "\n".join(lines)


def format_trends(var_df) -> str:
    """Formate les données de tendances pour le prompt."""
    if var_df is None or var_df.empty:
        return "Aucune donnée de tendance disponible (pas d'historique suffisant)."
    
    lines = ["=== TENDANCES D'ÉVOLUTION ==="]
    
    # Filtrer les variations significatives si la colonne existe
    if "Significatif" in var_df.columns:
        significant = var_df[var_df["Significatif"] == True]
    else:
        # Sinon filtrer sur l'écart %
        if "Ecart %" in var_df.columns:
            significant = var_df[var_df["Ecart %"].abs() >= 5]
        else:
            significant = var_df
    
    if significant.empty:
        lines.append("Aucune variation significative (>=5%) détectée sur la période.")
        return "\n".join(lines)
    
    # Grouper par sens d'évolution
    for sens in ["Degradation", "Amelioration", "Stable"]:
        if "Sens" in significant.columns:
            subset = significant[significant["Sens"] == sens]
        else:
            continue
            
        if subset.empty:
            continue
        
        lines.append(f"\n--- {sens.upper()} ---")
        count = 0
        for _, row in subset.iterrows():
            if count >= 20:  # Limiter à 20 lignes par catégorie
                lines.append(f"... et {len(subset) - 20} autres variations")
                break
            poste = row.get("Poste", "N/A")
            kpi = row.get("KPI", "N/A")
            prev = row.get("Valeur precedente", "N/A")
            curr = row.get("Valeur actuelle", "N/A")
            ecart_pct = row.get("Ecart %", 0)
            try:
                ecart_str = f"{float(ecart_pct):+.1f}%"
            except (ValueError, TypeError):
                ecart_str = "N/A"
            lines.append(f"• {poste} | {kpi}: {prev} → {curr} ({ecart_str})")
            count += 1
    
    return "\n".join(lines)


def format_rankings(best_df, worst_df) -> str:
    """Formate les classements des postes."""
    lines = ["=== CLASSEMENT DES POSTES ==="]
    
    if best_df is not None and not best_df.empty:
        lines.append("\n--- TOP 5 POSTES EN AMÉLIORATION ---")
        for _, row in best_df.head(5).iterrows():
            score = row.get("Score variation", 0)
            try:
                score_str = f"{float(score):+.1f} pts"
            except (ValueError, TypeError):
                score_str = "N/A"
            lines.append(f"• {row.get('Poste', 'N/A')}: {score_str}")
    
    if worst_df is not None and not worst_df.empty:
        lines.append("\n--- TOP 5 POSTES EN DÉGRADATION ---")
        for _, row in worst_df.head(5).iterrows():
            score = row.get("Score variation", 0)
            try:
                score_str = f"{float(score):+.1f} pts"
            except (ValueError, TypeError):
                score_str = "N/A"
            lines.append(f"• {row.get('Poste', 'N/A')}: {score_str}")
    
    if len(lines) == 1:
        lines.append("Aucun classement disponible.")
    
    return "\n".join(lines)


def format_existing_actions(actions_list) -> str:
    """Formate la liste des actions existantes."""
    if not actions_list:
        return "Aucun plan d'action existant."
    
    lines = ["=== PLANS D'ACTION EXISTANTS ==="]
    for i, action in enumerate(actions_list[:15], 1):  # Limiter à 15 actions
        if isinstance(action, dict):
            action_text = action.get("action", action.get("Action", "N/A"))
            responsable = action.get("responsable", action.get("Responsable", "N/A"))
            priorite = action.get("priorite", action.get("Priorité", "N/A"))
            lines.append(f"{i}. {action_text} | Responsable: {responsable} | Priorité: {priorite}")
        elif isinstance(action, str):
            lines.append(f"{i}. {action}")
    
    return "\n".join(lines)


# ============================================================
# Construction du prompt utilisateur
# ============================================================

def build_user_prompt(
    division: str,
    atelier: str,
    metier: str,
    date_debut: str,
    date_fin: str,
    kpi_performance_df,
    kpi_qualite_df,
    anomalies_perf_df=None,
    anomalies_qual_df=None,
    variations_df=None,
    best_ranking_df=None,
    worst_ranking_df=None,
    score_performance_global: float = 0.0,
    score_qualite_global: float = 0.0,
    total_ot_analyses: int = 0,
    total_anomalies: int = 0,
    existing_actions: Optional[List[Dict]] = None
) -> str:
    """Construit le prompt utilisateur avec toutes les données formatées."""
    
    prompt_parts = [
        "DONNÉES À ANALYSER :",
        "",
        f"DIVISION: {division}",
        f"ATELIER: {atelier}",
        f"MÉTIER: {metier}",
        f"PÉRIODE: Du {date_debut} au {date_fin}",
        "",
        "=" * 60,
        "SCORES GLOBAUX",
        "=" * 60,
        f"Score Performance Global: {score_performance_global:.1f}%",
        f"Score Qualité Global: {score_qualite_global:.1f}%",
        f"Total OT Analysés: {total_ot_analyses}",
        f"Total Anomalies Détectées: {total_anomalies}",
        "",
        format_kpi_table(kpi_performance_df, "KPI Performance"),
        "",
        format_kpi_table(kpi_qualite_df, "KPI Qualité"),
        "",
        format_anomalies_table(anomalies_perf_df, "Performance"),
        "",
        format_anomalies_table(anomalies_qual_df, "Qualité"),
        "",
        format_trends(variations_df),
        "",
        format_rankings(best_ranking_df, worst_ranking_df),
        "",
        format_existing_actions(existing_actions),
        "",
        "=" * 60,
        "INSTRUCTIONS FINALES",
        "=" * 60,
        "Analyser ces données et retourner UNIQUEMENT le JSON valide avec la structure demandée.",
        "Les conclusions doivent être courtes (une phrase max par slide).",
        "Utiliser le vocabulaire maintenance industrielle SAP PM.",
        "Ne pas inventer de données non fournies."
    ]
    
    return "\n".join(prompt_parts)


# ============================================================
# Fonction principale de génération
# ============================================================

def generate_analysis(
    division: str,
    atelier: str,
    metier: str,
    date_debut: str,
    date_fin: str,
    kpi_performance_df,
    kpi_qualite_df,
    anomalies_perf_df=None,
    anomalies_qual_df=None,
    variations_df=None,
    best_ranking_df=None,
    worst_ranking_df=None,
    score_performance_global: float = 0.0,
    score_qualite_global: float = 0.0,
    total_ot_analyses: int = 0,
    total_anomalies: int = 0,
    existing_actions: Optional[List[Dict]] = None,
    model: str = "deepseek/deepseek-chat"
) -> Dict[str, Any]:
    """
    Génère l'analyse AI des KPI maintenance.
    
    Args:
        division: Nom de la division
        atelier: Nom de l'atelier
        metier: Métier concerné
        date_debut: Date de début de la période (format dd/mm/yyyy)
        date_fin: Date de fin de la période (format dd/mm/yyyy)
        kpi_performance_df: DataFrame des KPI Performance
        kpi_qualite_df: DataFrame des KPI Qualité
        anomalies_perf_df: DataFrame des anomalies Performance (optionnel)
        anomalies_qual_df: DataFrame des anomalies Qualité (optionnel)
        variations_df: DataFrame des variations/tendances (optionnel)
        best_ranking_df: DataFrame du classement meilleur (optionnel)
        worst_ranking_df: DataFrame du classement pire (optionnel)
        score_performance_global: Score performance global en %
        score_qualite_global: Score qualité global en %
        total_ot_analyses: Nombre total d'OT analysés
        total_anomalies: Nombre total d'anomalies
        existing_actions: Liste des actions existantes (optionnel)
        model: Modèle AI à utiliser (défaut: deepseek/deepseek-chat)
    
    Returns:
        Dictionnaire JSON avec la structure attendue pour la présentation
    """
    
    # Préparer le système prompt avec les variables
    system_prompt = SYSTEM_PROMPT.format(
        division=division,
        atelier=atelier,
        metier=metier,
        date_debut=date_debut,
        date_fin=date_fin
    )
    
    # Construire le prompt utilisateur
    user_prompt = build_user_prompt(
        division=division,
        atelier=atelier,
        metier=metier,
        date_debut=date_debut,
        date_fin=date_fin,
        kpi_performance_df=kpi_performance_df,
        kpi_qualite_df=kpi_qualite_df,
        anomalies_perf_df=anomalies_perf_df,
        anomalies_qual_df=anomalies_qual_df,
        variations_df=variations_df,
        best_ranking_df=best_ranking_df,
        worst_ranking_df=worst_ranking_df,
        score_performance_global=score_performance_global,
        score_qualite_global=score_qualite_global,
        total_ot_analyses=total_ot_analyses,
        total_anomalies=total_anomalies,
        existing_actions=existing_actions
    )
    
    try:
        # Appel à l'API OpenAI
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,  # Température basse pour des réponses cohérentes et précises
            max_tokens=4000,
            response_format={"type": "json_object"}
        )
        
        # Extraire le contenu de la réponse
        content = response.choices[0].message.content
        
        # Nettoyer le contenu si nécessaire (retirer les blocs de code markdown)
        if content.startswith("```json"):
            content = content[7:]
        elif content.startswith("```"):
            content = content[3:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()
        
        # Parser le JSON
        result = json.loads(content)
        
        # Valider et compléter la structure
        result = validate_and_complete_structure(
            result, division, atelier, metier, date_debut, date_fin
        )
        
        return result
        
    except json.JSONDecodeError as e:
        return get_error_structure(f"Erreur de parsing JSON: {str(e)}")
    except Exception as e:
        return get_error_structure(f"Erreur lors de l'appel API: {str(e)}")


# ============================================================
# Validation et complétion de la structure
# ============================================================

def validate_and_complete_structure(
    result: Dict, 
    division: str, 
    atelier: str, 
    metier: str, 
    date_debut: str, 
    date_fin: str
) -> Dict:
    """Valide et complète la structure JSON retournée par l'IA."""
    
    # Structure minimale attendue
    default_structure = {
        "title": "Analyse des KPI Performance et Qualité",
        "subtitle": f"{division} | {atelier} | {metier} | Du {date_debut} au {date_fin}",
        "executive_summary": "",
        "slide1": {"conclusion": ""},
        "slide2": {"conclusion": ""},
        "slide3": {"conclusion": ""},
        "slide4": {"root_causes": []},
        "slide5": {"recommendations": []},
        "slide6": {"final_conclusion": ""}
    }
    
    # Fusionner avec le résultat (le résultat écrase les défauts)
    for key in default_structure:
        if key not in result:
            result[key] = default_structure[key]
        elif isinstance(default_structure[key], dict) and isinstance(result[key], dict):
            for subkey in default_structure[key]:
                if subkey not in result[key]:
                    result[key][subkey] = default_structure[key][subkey]
    
    # S'assurer que le titre et sous-titre sont corrects
    result["title"] = default_structure["title"]
    result["subtitle"] = default_structure["subtitle"]
    
    # S'assurer que les listes sont bien des listes
    if not isinstance(result.get("slide4", {}).get("root_causes"), list):
        result["slide4"]["root_causes"] = []
    if not isinstance(result.get("slide5", {}).get("recommendations"), list):
        result["slide5"]["recommendations"] = []
    
    # Valider et normaliser les causes racines
    valid_criticites = ["Critique", "Élevée", "Moyenne", "Faible"]
    valid_causes = []
    for cause in result["slide4"]["root_causes"]:
        if isinstance(cause, dict):
            validated_cause = {
                "cause": str(cause.get("cause", cause.get("Cause", "")))[:300],
                "criticite": cause.get("criticite", cause.get("Criticité", "Moyenne"))
            }
            # Normaliser la criticité
            crit = validated_cause["criticite"]
            if crit not in valid_criticites:
                validated_cause["criticite"] = "Moyenne"
            valid_causes.append(validated_cause)
        elif isinstance(cause, str):
            valid_causes.append({"cause": cause[:300], "criticite": "Moyenne"})
    result["slide4"]["root_causes"] = valid_causes
    
    # Valider et normaliser les recommandations
    valid_priorities = ["Haute", "Moyenne", "Basse"]
    valid_rec = []
    for rec in result["slide5"]["recommendations"]:
        if isinstance(rec, dict):
            # Gérer les différentes clés possibles (avec ou sans accents)
            action = rec.get("action", rec.get("Action", ""))
            responsable = rec.get("responsable", rec.get("Responsable", ""))
            priorite = rec.get("priorite", rec.get("Priorité", rec.get("Priorite", "Moyenne")))
            echeance = rec.get("echeance", rec.get("Échéance", rec.get("Echeance", "")))
            impact = rec.get("impact_attendu", rec.get("Impact attendu", rec.get("impact", "")))
            
            validated_rec = {
                "action": str(action)[:250],
                "responsable": str(responsable)[:60],
                "priorite": priorite if priorite in valid_priorities else "Moyenne",
                "echeance": str(echeance)[:60],
                "impact_attendu": str(impact)[:200]
            }
            valid_rec.append(validated_rec)
        elif isinstance(rec, str):
            valid_rec.append({
                "action": rec[:250],
                "responsable": "À définir",
                "priorite": "Moyenne",
                "echeance": "À définir",
                "impact_attendu": ""
            })
    result["slide5"]["recommendations"] = valid_rec
    
    return result


def get_error_structure(error_message: str) -> Dict:
    """Retourne une structure d'erreur standardisée."""
    return {
        "title": "Analyse des KPI Performance et Qualité",
        "subtitle": "Erreur lors de l'analyse",
        "executive_summary": "",
        "slide1": {"conclusion": f"Erreur lors de l'analyse: {error_message}"},
        "slide2": {"conclusion": "Données non disponibles en raison d'une erreur technique."},
        "slide3": {"conclusion": "Données non disponibles en raison d'une erreur technique."},
        "slide4": {"root_causes": [{"cause": f"Erreur technique: {error_message}", "criticite": "Critique"}]},
        "slide5": {"recommendations": []},
        "slide6": {"final_conclusion": "Veuillez réessayer ultérieurement ou vérifier les données d'entrée."}
    }


# ============================================================
# Fonction d'intégration pour Streamlit
# ============================================================

def generate_analysis_from_dashboard(
    st,
    division: str,
    atelier: str,
    metier: str,
    date_debut,
    date_fin,
    kpi_perf_rows: List[Dict],
    kpi_perf_cols: List[str],
    kpi_qual_rows: List[Dict],
    kpi_qual_cols: List[str],
    ano_perf_rows: List[Dict],
    ano_perf_cols: List[str],
    ano_qual_rows: List[Dict],
    ano_qual_cols: List[str],
    var_df,
    best_ranking_df,
    worst_ranking_df,
    score_perf: float,
    score_qual: float,
    total_ot: int,
    total_anomalies: int,
    plan_action_df=None,
    model: str = "deepseek/deepseek-chat",
    show_progress: bool = True
) -> Dict[str, Any]:
    """
    Fonction wrapper pour générer l'analyse depuis le dashboard Streamlit.
    
    Args:
        st: Instance Streamlit
        division, atelier, metier: Filtres sélectionnés
        date_debut, date_fin: Dates de la période
        kpi_perf_rows, kpi_perf_cols: Données KPI Performance
        kpi_qual_rows, kpi_qual_cols: Données KPI Qualité
        ano_perf_rows, ano_perf_cols: Anomalies Performance
        ano_qual_rows, ano_qual_cols: Anomalies Qualité
        var_df: DataFrame des variations
        best_ranking_df, worst_ranking_df: Classements
        score_perf, score_qual: Scores globaux
        total_ot, total_anomalies: Totaux
        plan_action_df: DataFrame du plan d'action (optionnel)
        model: Modèle AI à utiliser
        show_progress: Afficher la barre de progression
    
    Returns:
        Dictionnaire JSON avec la structure attendue
    """
    
    # Convertir les dates en string
    if hasattr(date_debut, 'strftime'):
        date_debut_str = date_debut.strftime("%d/%m/%Y")
    else:
        date_debut_str = str(date_debut)
    
    if hasattr(date_fin, 'strftime'):
        date_fin_str = date_fin.strftime("%d/%m/%Y")
    else:
        date_fin_str = str(date_fin)
    
    # Convertir les listes de dictionnaires en DataFrame
    kpi_perf_df = pd.DataFrame(kpi_perf_rows, columns=kpi_perf_cols) if kpi_perf_rows else pd.DataFrame()
    kpi_qual_df = pd.DataFrame(kpi_qual_rows, columns=kpi_qual_cols) if kpi_qual_rows else pd.DataFrame()
    ano_perf_df = pd.DataFrame(ano_perf_rows, columns=ano_perf_cols) if ano_perf_rows else pd.DataFrame()
    ano_qual_df = pd.DataFrame(ano_qual_rows, columns=ano_qual_cols) if ano_qual_rows else pd.DataFrame()
    
    # Extraire les actions existantes du plan d'action
    existing_actions = None
    if plan_action_df is not None and not plan_action_df.empty:
        existing_actions = plan_action_df.to_dict('records')
    
    # Afficher la progression
    if show_progress:
        progress_container = st.empty()
        progress_container.markdown("""
        <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#1e3a5f,#2563eb);
                    border-radius:12px;color:white;margin:10px 0">
            <div style="font-size:24px;margin-bottom:10px">🤖 Analyse IA en cours...</div>
            <div style="font-size:14px;opacity:0.8">Génération de l'analyse des KPI maintenance</div>
        </div>
        """, unsafe_allow_html=True)
    
    try:
        # Générer l'analyse
        result = generate_analysis(
            division=division,
            atelier=atelier,
            metier=metier,
            date_debut=date_debut_str,
            date_fin=date_fin_str,
            kpi_performance_df=kpi_perf_df,
            kpi_qualite_df=kpi_qual_df,
            anomalies_perf_df=ano_perf_df if not ano_perf_df.empty else None,
            anomalies_qual_df=ano_qual_df if not ano_qual_df.empty else None,
            variations_df=var_df,
            best_ranking_df=best_ranking_df,
            worst_ranking_df=worst_ranking_df,
            score_performance_global=score_perf,
            score_qualite_global=score_qual,
            total_ot_analyses=total_ot,
            total_anomalies=total_anomalies,
            existing_actions=existing_actions,
            model=model
        )
        
        if show_progress:
            progress_container.markdown("""
            <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#059669,#10b981);
                        border-radius:12px;color:white;margin:10px 0">
                <div style="font-size:24px;margin-bottom:10px">✅ Analyse terminée avec succès</div>
                <div style="font-size:14px;opacity:0.8">Les résultats sont prêts pour la présentation</div>
            </div>
            """, unsafe_allow_html=True)
        
        return result
        
    except Exception as e:
        if show_progress:
            progress_container.markdown(f"""
            <div style="padding:20px;text-align:center;background:linear-gradient(135deg,#dc2626,#ef4444);
                        border-radius:12px;color:white;margin:10px 0">
                <div style="font-size:24px;margin-bottom:10px">❌ Erreur lors de l'analyse</div>
                <div style="font-size:14px;opacity:0.8">{str(e)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        return get_error_structure(str(e))


# ============================================================
# Test unitaire
# ============================================================

if __name__ == "__main__":
    print("=" * 60)
    print("TEST UNITAIRE - AI Agent KPI Maintenance")
    print("=" * 60)
    
    # Données de test
    test_perf_df = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-002", "SF2-001", "Total general"],
        "TAUX_REALISATION_CORRECTIF/PT": [72.5, 68.0, 92.3, 77.6],
        "OT préparation <1 mois": [55.0, 48.0, 88.0, 63.7],
        "OT préparation >3 mois": [18.0, 22.0, 3.0, 14.3],
        "OT planification <1 mois": [62.0, 58.0, 90.0, 70.0],
        "Score Performance": [65.0, 58.0, 91.0, 71.3]
    })
    
    test_qual_df = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-002", "SF2-001", "Total general"],
        "Taux d'approbation des Avis": [88.0, 82.0, 98.0, 89.3],
        "OT CONFIME": [75.0, 68.0, 100.0, 81.0],
        "Performance Graissage": [92.0, 85.0, 98.0, 91.7],
        "Score Qualite": [85.0, 78.3, 98.7, 87.3]
    })
    
    test_anomalies = pd.DataFrame({
        "Poste de travail": ["SF1-001", "SF1-001", "SF1-002", "SF1-002", "SF1-002"],
        "KPI": [
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "OT préparation >3 mois"
        ],
        "Valeur": [72.5, 55.0, 68.0, 48.0, 22.0],
        "Cible": [85.0, 80.0, 85.0, 80.0, 5.0],
        "Écart": [-12.5, -25.0, -17.0, -32.0, 17.0]
    })
    
    test_variations = pd.DataFrame({
        "Poste": ["SF1-001", "SF1-002", "SF2-001", "SF1-001", "SF1-002"],
        "KPI": [
            "TAUX_REALISATION_CORRECTIF/PT",
            "TAUX_REALISATION_CORRECTIF/PT",
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois",
            "OT préparation <1 mois"
        ],
        "Valeur precedente": [78.0, 75.0, 90.0, 65.0, 60.0],
        "Valeur actuelle": [72.5, 68.0, 92.3, 55.0, 48.0],
        "Ecart %": [-7.1, -9.3, 2.6, -15.4, -20.0],
        "Sens": ["Degradation", "Degradation", "Amelioration", "Degradation", "Degradation"],
        "Significatif": [True, True, False, True, True]
    })
    
    test_best = pd.DataFrame({
        "Poste": ["SF2-001", "SF2-002"],
        "Score variation": [8.5, 3.2]
    })
    
    test_worst = pd.DataFrame({
        "Poste": ["SF1-002", "SF1-001"],
        "Score variation": [-35.0, -22.0]
    })
    
    # Exécuter le test
    print("\nAppel à l'API IA...")
    result = generate_analysis(
        division="Division Chimique Nord",
        atelier="Atelier Production A",
        metier="Mécanique",
        date_debut="01/01/2024",
        date_fin="31/01/2024",
        kpi_performance_df=test_perf_df,
        kpi_qualite_df=test_qual_df,
        anomalies_perf_df=test_anomalies,
        anomalies_qual_df=pd.DataFrame(),  # Pas d'anomalies qualité dans le test
        variations_df=test_variations,
        best_ranking_df=test_best,
        worst_ranking_df=test_worst,
        score_performance_global=71.3,
        score_qualite_global=87.3,
        total_ot_analyses=156,
        total_anomalies=5
    )
    
    print("\n" + "=" * 60)
    print("RÉSULTAT JSON:")
    print("=" * 60)
    print(json.dumps(result, indent=2, ensure_ascii=False))
