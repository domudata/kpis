from openai import OpenAI
import json

client = OpenAI(
    api_key="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",
    base_url="https://openrouter.ai/api/v1"
)
def ask_qwen(prompt):

    response = client.chat.completions.create(
        model="qwen/qwen3.6-plus",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    return response.choices[0].message.content
def analyse_kpi(
    ckdf,
    pscores,
    qscores,
    division,
    atelier,
    metier,
    date_debut,
    date_fin
):

    prompt = f"""
Tu es un ingénieur maintenance senior.

Contexte

Division : {division}
Atelier : {atelier}
Métier : {metier}

Période

Du {date_debut}
Au {date_fin}

Voici les KPI :

{ckdf.to_markdown()}

Score Performance

{pscores}

Score Qualité

{qscores}

Analyse uniquement les KPI.

Réponds en JSON.

Format :

{{
"resume":"",
"kpi_critiques":"",
"points_forts":"",
"points_faibles":"",
"conclusion":""
}}
"""

    txt = ask_qwen(prompt)

    return txt
