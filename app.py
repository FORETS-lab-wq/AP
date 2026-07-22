# -*- coding: utf-8 -*-
"""
Générateur de plannings de formation — IMSR / ECF-CERCA
Application Streamlit réutilisable : saisie des informations, des modules
(CCP/BC/UC) et des périodes de modalité, puis génération d'un fichier Excel
téléchargeable au format du modèle "Planning_contrat_pro".
"""

import io
from datetime import date, timedelta

import streamlit as st
import pandas as pd

from planning_engine import generate_workbook

st.set_page_config(page_title="Générateur de plannings IMSR", page_icon="📅", layout="wide")

TRAININGS_LONG = ["TP ECSR", "TFP FMESR", "CQP RUESRC", "CQP AGEC"]
TRAININGS_SHORT = [
    "Stage de réactualisation des connaissances",
    "Stage d'habilitation post-permis",
    "Formation continue",
]
CONTRACT_TYPES = [
    "Contrat de professionnalisation",
    "Contrat d'apprentissage",
    "Stagiaire de la formation continue",
    "Autre / à préciser",
]
MODALITES = ["Centre", "FOAD", "Entreprise", "Examens"]
MODULES_MODALITES = ["Présentiel", "Distanciel"]

# Modules prédéfinis selon le type de formation (modifiables par l'utilisateur)
MODULES_BY_TRAINING = {
    "TP ECSR": ["CCP1", "CCP2"],
    "TFP FMESR": ["BC1", "BC2", "BC3"],
    "CQP RUESRC": ["BC1", "BC2", "BC3"],
    "CQP AGEC": ["BC1", "BC2", "BC3"],
}

st.title("📅 Générateur de plannings de formation — IMSR")
st.caption(
    "Outil réutilisable : renseignez les informations ci-dessous puis téléchargez "
    "le fichier Excel généré. Relancez l'outil pour chaque nouveau planning."
)

# ---------------------------------------------------------------------------
# 1. Informations générales
# ---------------------------------------------------------------------------
st.header("1. Informations générales")

col1, col2 = st.columns(2)
with col1:
    family = st.radio(
        "Famille de formation",
        ["Formation longue (TP ECSR, TFP FMESR, CQP RUESRC, CQP AGEC)",
         "Formation courte (réactualisation, habilitation post-permis, formation continue)"],
        key="family",
    )
    is_long = family.startswith("Formation longue")
    training_type = st.selectbox(
        "Type de formation / action",
        TRAININGS_LONG if is_long else TRAININGS_SHORT,
    )

with col2:
    learner_name = st.text_input("Nom de l'apprenant(e) / stagiaire", placeholder="Prénom NOM")
    contract_type = st.selectbox("Type de contrat / statut", CONTRACT_TYPES)
    if contract_type == "Autre / à préciser":
        contract_type = st.text_input("Précisez le type de contrat", value="")

col3, col4 = st.columns(2)
with col3:
    default_start = date.today()
    start_date = st.date_input("Date de début", value=default_start, format="DD/MM/YYYY")
with col4:
    if is_long:
        default_end = date(start_date.year + 1, 8, 31) if start_date.month >= 9 else date(start_date.year, 8, 31)
    else:
        default_end = start_date + timedelta(days=4)
    end_date = st.date_input("Date de fin", value=default_end, format="DD/MM/YYYY")

if end_date < start_date:
    st.error("La date de fin doit être postérieure à la date de début.")
    st.stop()

mode = "long" if is_long else "short"
if is_long:
    nb_years = end_date.year - start_date.year + (1 if end_date.month >= 9 else 0)
    st.info(
        f"Le planning couvrira {max(1, nb_years)} feuille(s) « année scolaire » "
        f"(calendrier Septembre → Août), du {start_date.strftime('%d/%m/%Y')} "
        f"au {end_date.strftime('%d/%m/%Y')}."
    )

st.divider()

# ---------------------------------------------------------------------------
# 2. Modules (CCP / BC / UC) et volumes horaires
# ---------------------------------------------------------------------------
st.header("2. Modules (CCP / BC / UC) et volumes horaires")
st.caption(
    "Cet onglet récapitule les modules de l'action : la modalité (présentiel / distanciel) "
    "et le volume horaire, distingué entre heures de formation et heures de stage. "
    "Il reste indépendant des cases du calendrier (qui restent en Centre / FOAD / Entreprise / Examens)."
)

available_codes = MODULES_BY_TRAINING.get(training_type, [])
code_column = (
    st.column_config.SelectboxColumn("Module", options=available_codes)
    if available_codes
    else st.column_config.TextColumn("Module")
)

default_modules = pd.DataFrame(
    [{
        "Code": available_codes[0] if available_codes else "",
        "Intitulé du module": "",
        "Modalité": "Présentiel",
        "Heures de formation": 0,
        "Heures de stage": 0,
    }]
)
modules_df = st.data_editor(
    default_modules,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Code": code_column,
        "Modalité": st.column_config.SelectboxColumn(options=MODULES_MODALITES),
        "Heures de formation": st.column_config.NumberColumn(min_value=0, step=1),
        "Heures de stage": st.column_config.NumberColumn(min_value=0, step=1),
    },
    key="modules_editor",
)

total_formation = pd.to_numeric(modules_df["Heures de formation"], errors="coerce").fillna(0).sum()
total_stage = pd.to_numeric(modules_df["Heures de stage"], errors="coerce").fillna(0).sum()
c_m1, c_m2, c_m3 = st.columns(3)
c_m1.metric("Total heures de formation", f"{int(total_formation)} h")
c_m2.metric("Total heures de stage", f"{int(total_stage)} h")
c_m3.metric("Total général", f"{int(total_formation + total_stage)} h")

st.divider()

# ---------------------------------------------------------------------------
# 3. Directives récurrentes (ex: "CCP2 en FOAD tous les jeudis")
# ---------------------------------------------------------------------------
st.header("3. Directives récurrentes")
st.caption(
    "Pour les règles qui se répètent chaque semaine (ex. « CCP2 en distanciel tous les "
    "jeudis »), plutôt que de redéfinir des périodes. Une directive s'applique à un jour "
    "de semaine donné, sur la plage de dates choisie, et **prend le pas** sur la "
    "répartition par périodes définie à l'étape 4. Le module lié (optionnel) permet de "
    "cumuler automatiquement les heures planifiées dans l'onglet Modules."
)

module_codes = [c for c in modules_df["Code"].tolist() if isinstance(c, str) and c.strip()]

default_directives = pd.DataFrame(
    [{
        "Jour de la semaine": "Jeudi",
        "Modalité": "FOAD",
        "Module lié (optionnel)": "",
        "Début": start_date,
        "Fin": end_date,
    }]
).iloc[0:0]  # tableau vide au départ

directives_df = st.data_editor(
    default_directives,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Jour de la semaine": st.column_config.SelectboxColumn(
            options=["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        ),
        "Modalité": st.column_config.SelectboxColumn(options=MODALITES),
        "Module lié (optionnel)": st.column_config.SelectboxColumn(
            options=[""] + module_codes
        ),
        "Début": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Fin": st.column_config.DateColumn(format="DD/MM/YYYY"),
    },
    key="directives_editor",
)

st.divider()

# ---------------------------------------------------------------------------
# 4. Remarques et structuration libre
# ---------------------------------------------------------------------------
st.header("4. Remarques et structuration")
st.caption(
    "Espace libre pour noter la structuration générale du planning (grandes périodes, "
    "alternance, points de vigilance...). Ce texte est repris tel quel dans la zone "
    "« COMMENTAIRES » du fichier Excel généré. Le remplissage précis des jours reste "
    "piloté par les directives récurrentes de l'étape 3."
)

comments = st.text_area(
    "Notes / structuration libre",
    height=140,
    placeholder=(
        "Ex : Septembre-décembre en Centre, alternance FOAD/Entreprise à partir de janvier, "
        "période d'examens prévue en mai, semaine banalisée pour le CCP2 fin juin..."
    ),
)

with st.expander("💡 S'inspirer d'un planning existant (optionnel)"):
    st.caption(
        "Importez un fichier Excel de planning déjà réalisé pour le consulter pendant votre "
        "saisie — il ne sera pas utilisé dans le fichier généré, c'est uniquement pour vous "
        "en inspirer."
    )
    reference_file = st.file_uploader(
        "Planning existant (.xlsx)", type=["xlsx"], key="reference_upload"
    )
    if reference_file is not None:
        try:
            xls = pd.ExcelFile(reference_file)
            sheet_choice = st.selectbox("Feuille à consulter", xls.sheet_names, key="reference_sheet")
            preview_df = pd.read_excel(reference_file, sheet_name=sheet_choice, header=None, nrows=45)
            st.dataframe(preview_df, use_container_width=True, height=420)
        except Exception as e:
            st.warning(f"Impossible de lire ce fichier : {e}")

st.divider()

# ---------------------------------------------------------------------------
# 5. Génération
# ---------------------------------------------------------------------------
st.header("5. Générer le planning")

if st.button("📥 Générer le fichier Excel", type="primary"):
    modules = []
    for _, row in modules_df.iterrows():
        code = row.get("Code", "")
        intitule = row.get("Intitulé du module", "")
        modalite = row.get("Modalité", "")
        h_formation = row.get("Heures de formation", 0) or 0
        h_stage = row.get("Heures de stage", 0) or 0
        if (isinstance(code, str) and code.strip()) or (isinstance(intitule, str) and intitule.strip()):
            modules.append({
                "code": code,
                "intitule": intitule,
                "modalite": modalite,
                "heures_formation": h_formation,
                "heures_stage": h_stage,
            })

    directives = []
    for _, row in directives_df.iterrows():
        jour = row.get("Jour de la semaine")
        mod = row.get("Modalité")
        module = row.get("Module lié (optionnel)") or None
        d0 = row.get("Début")
        d1 = row.get("Fin")
        if not jour or not mod:
            continue
        if pd.isna(d0):
            d0 = None
        elif isinstance(d0, pd.Timestamp):
            d0 = d0.date()
        if pd.isna(d1):
            d1 = None
        elif isinstance(d1, pd.Timestamp):
            d1 = d1.date()
        directives.append({
            "jour": jour, "modalite": mod, "module": module or None,
            "debut": d0 if d0 is not None else start_date,
            "fin": d1 if d1 is not None else end_date,
        })

    wb = generate_workbook(
        learner_name=learner_name,
        contract_type=contract_type,
        training_type=training_type,
        mode=mode,
        start=start_date,
        end=end_date,
        periods=[],
        modules=modules,
        directives=directives,
        comments=comments,
    )

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    safe_name = (learner_name or "planning").strip().replace(" ", "_") or "planning"
    filename = f"Planning_{training_type.replace(' ', '_')}_{safe_name}.xlsx"

    st.success("Planning généré avec succès.")
    st.download_button(
        "⬇️ Télécharger le fichier Excel",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
