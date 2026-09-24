# -*- coding: utf-8 -*-
"""
============================================================================
 APPLICATION DE RAPPROCHEMENT BANCAIRE AUTOMATISÉ (Bank Reconciliation)
----------------------------------------------------------------------------
 Algorithme : Many-to-One avec agrégation ERP, condition temporelle
              "Weekend-Aware" et tolérance financière (<= 700 Dhs).
 Stack      : Streamlit + Pandas + openpyxl
============================================================================
 Lancement :  streamlit run app_rapprochement_bancaire.py
============================================================================
"""

import io
from datetime import timedelta

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIGURATION GLOBALE
# ---------------------------------------------------------------------------
TOLERANCE = 700.0  # Tolérance absolue en Dhs

st.set_page_config(
    page_title="Rapprochement Bancaire Automatisé",
    page_icon="🏦",
    layout="wide",
)

# ---------------------------------------------------------------------------
# CONSTANTES DE COLONNES
# ---------------------------------------------------------------------------
ERP_REQUIRED = ["Reference", "Description", "Nom", "N°Document",
                "Date_ERP", "Debit", "Credit"]
RLV_REQUIRED = ["Date_RLV", "Description", "Debit", "Credit"]


# ===========================================================================
# 1. CHARGEMENT & PARSING DES DONNÉES
# ===========================================================================
def normaliser_colonnes(df: pd.DataFrame) -> pd.DataFrame:
    """Supprime les espaces superflus dans les noms de colonnes."""
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_numeric_safe(serie: pd.Series) -> pd.Series:
    """
    Convertit une colonne en numérique de manière robuste :
    gère les séparateurs de milliers, virgules décimales et espaces.
    """
    if serie.dtype == object:
        serie = (
            serie.astype(str)
            .str.replace("\u00a0", "", regex=False)   # espace insécable
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
    return pd.to_numeric(serie, errors="coerce").fillna(0.0)


@st.cache_data(show_spinner=False)
def charger_donnees(file_bytes: bytes):
    """
    Charge le fichier Excel unique et retourne (df_erp, df_rlv).
    - Vérifie la présence des onglets ERP et RLV.
    - Vérifie la présence des colonnes obligatoires.
    - Convertit IMPÉRATIVEMENT Date_ERP / Date_RLV en datetime réels
      (pd.to_datetime avec errors='coerce') pour éviter l'erreur
      'AttributeError: Can only use .dt accessor with datetimelike values'.
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    onglets = {s.strip().upper(): s for s in xls.sheet_names}

    if "ERP" not in onglets or "RLV" not in onglets:
        raise ValueError(
            f"Le fichier doit contenir les onglets 'ERP' et 'RLV'. "
            f"Onglets trouvés : {xls.sheet_names}"
        )

    df_erp = normaliser_colonnes(pd.read_excel(xls, sheet_name=onglets["ERP"]))
    df_rlv = normaliser_colonnes(pd.read_excel(xls, sheet_name=onglets["RLV"]))

    # --- Vérification des colonnes requises -------------------------------
    manquantes_erp = [c for c in ERP_REQUIRED if c not in df_erp.columns]
    manquantes_rlv = [c for c in RLV_REQUIRED if c not in df_rlv.columns]
    if manquantes_erp:
        raise ValueError(f"Onglet ERP : colonnes manquantes -> {manquantes_erp}")
    if manquantes_rlv:
        raise ValueError(f"Onglet RLV : colonnes manquantes -> {manquantes_rlv}")

    # --- PARSING TEMPOREL OBLIGATOIRE -------------------------------------
    df_erp["Date_ERP"] = pd.to_datetime(df_erp["Date_ERP"], errors="coerce")
    df_rlv["Date_RLV"] = pd.to_datetime(df_rlv["Date_RLV"], errors="coerce")

    # --- Conversion numérique robuste des montants ------------------------
    for col in ["Debit", "Credit"]:
        df_erp[col] = to_numeric_safe(df_erp[col])
        df_rlv[col] = to_numeric_safe(df_rlv[col])

    # --- Montant net (Debit - Credit) -------------------------------------
    df_erp["Montant_ERP"] = df_erp["Debit"] - df_erp["Credit"]
    df_rlv["Montant_RLV"] = df_rlv["Debit"] - df_rlv["Credit"]

    return df_erp, df_rlv


# ===========================================================================
# 2. LOGIQUE TEMPORELLE (WEEKEND-AWARE)
# ===========================================================================
def dates_rlv_valides(date_erp: pd.Timestamp):
    """
    Retourne la liste des dates RLV acceptables pour une Date_ERP donnée.

    Règles :
      - Lundi(0) à Jeudi(3) : Date_RLV == Date_ERP + 1 jour.
      - Vendredi(4), Samedi(5), Dimanche(6) : Date_RLV == Lundi OU Mardi
        qui suit la Date_ERP.
    """
    if pd.isna(date_erp):
        return []

    jour = date_erp.weekday()  # Lundi=0 ... Dimanche=6

    if jour <= 3:  # Lundi -> Jeudi
        return [date_erp.normalize() + timedelta(days=1)]

    # Vendredi / Samedi / Dimanche -> chercher le lundi suivant
    delta_vers_lundi = (7 - jour)  # Ven=3, Sam=2, Dim=1 jours jusqu'au lundi
    lundi = date_erp.normalize() + timedelta(days=delta_vers_lundi)
    mardi = lundi + timedelta(days=1)
    return [lundi, mardi]


# ===========================================================================
# 3. MOTEUR DE RAPPROCHEMENT (MANY-TO-ONE + TOLÉRANCE)
# ===========================================================================
def lancer_rapprochement(df_erp: pd.DataFrame, df_rlv: pd.DataFrame):
    """
    Exécute l'algorithme complet et renvoie (matched_df, unmatched_rlv_df).
    """
    # --- Agrégation ERP par Reference (Many) ------------------------------
    agg = (
        df_erp.groupby("Reference", dropna=False)
        .agg(
            Somme_ERP=("Montant_ERP", "sum"),
            Debit_ERP=("Debit", "sum"),
            Credit_ERP=("Credit", "sum"),
            Date_ERP=("Date_ERP", "min"),
            Nom=("Nom", "first"),
            Description_ERP=("Description", "first"),
            NbLignes_ERP=("Montant_ERP", "size"),
            Documents=("N°Document",
                       lambda s: ", ".join(sorted({str(x) for x in s if pd.notna(x)}))),
        )
        .reset_index()
    )

    # Copie de travail du RLV avec indicateur d'utilisation
    rlv = df_rlv.reset_index(drop=True).copy()
    rlv["_used"] = False

    matched_rows = []
    match_id = 0  # identifiant unique de correspondance (1 par ligne RLV rapprochée)

    # --- Boucle de correspondance (One) -----------------------------------
    for _, grp in agg.iterrows():
        dates_ok = dates_rlv_valides(grp["Date_ERP"])
        if not dates_ok:
            continue

        # Candidats RLV : bonne date ET non encore utilisés
        masque = (
            (~rlv["_used"])
            & (rlv["Date_RLV"].dt.normalize().isin(dates_ok))
        )
        candidats = rlv[masque]
        if candidats.empty:
            continue

        # Écart absolu + condition de tolérance
        ecarts_abs = (candidats["Montant_RLV"] - grp["Somme_ERP"]).abs()
        eligibles = ecarts_abs[ecarts_abs <= TOLERANCE]
        if eligibles.empty:
            continue

        # Meilleur candidat = écart minimal
        idx_best = eligibles.idxmin()
        ligne_rlv = rlv.loc[idx_best]

        # --- Écart = Montant_RLV - Somme_ERP ------------------------------
        # Calculé sur les valeurs absolues pour rester cohérent avec les
        # montants (positifs) affichés et exportés. Un écart POSITIF signifie
        # que le relevé (RLV) est supérieur à la somme ERP.
        ecart = abs(ligne_rlv["Montant_RLV"]) - abs(grp["Somme_ERP"])

        rlv.at[idx_best, "_used"] = True
        match_id += 1

        # --- Éclatement : une ligne de sortie par N°Document de l'ERP -----
        lignes_erp = df_erp[df_erp["Reference"] == grp["Reference"]]
        premier = True
        for _, lg in lignes_erp.iterrows():
            matched_rows.append({
                "Match_ID": match_id,
                "Reference": grp["Reference"],
                "Nom": grp["Nom"],
                "N°Document": lg["N°Document"],
                "Description_ERP": lg["Description"],
                "Date_ERP": lg["Date_ERP"],
                "Montant_Ligne_ERP": round(lg["Montant_ERP"], 2),
                # Champs de niveau "groupe / RLV" : uniquement sur 1re ligne
                "Somme_ERP": round(grp["Somme_ERP"], 2) if premier else pd.NA,
                "Date_RLV": ligne_rlv["Date_RLV"] if premier else pd.NaT,
                "Description_RLV": ligne_rlv["Description"] if premier else "",
                "Montant_RLV": round(ligne_rlv["Montant_RLV"], 2) if premier else pd.NA,
                "Écart": round(ecart, 2) if premier else pd.NA,
            })
            premier = False

    # --- Tableaux de sortie ----------------------------------------------
    # Ordre : bloc ERP détaillé (une ligne par N°Document), puis bloc
    #         groupe/RLV (rempli une seule fois par correspondance), puis Écart.
    colonnes_matched = [
        "Match_ID", "Reference", "Nom", "N°Document", "Description_ERP",
        "Date_ERP", "Montant_Ligne_ERP",
        "Somme_ERP", "Date_RLV", "Description_RLV", "Montant_RLV", "Écart",
    ]
    matched_df = pd.DataFrame(matched_rows, columns=colonnes_matched)

    # --- Remplissage de Date_RLV sur toutes les lignes du groupe ----------
    # ffill() (Forward Fill) par correspondance : la date du relevé est
    # propagée sur chaque N°Document du même groupe -> aucune case vide,
    # structure identique à la colonne Date_ERP.
    if not matched_df.empty:
        matched_df["Date_RLV"] = matched_df.groupby("Match_ID")["Date_RLV"].ffill()

        # --- Valeur absolue des montants (affichage & export) -------------
        # Les montants sont rendus positifs (.abs()) car un Credit produit
        # un net négatif (Debit - Credit). Seule la colonne 'Écart' garde
        # son signe (positif ou négatif) pour rester interprétable.
        for col_montant in ["Somme_ERP", "Montant_Ligne_ERP", "Montant_RLV"]:
            matched_df[col_montant] = matched_df[col_montant].abs()

    unmatched_rlv_df = (
        rlv[~rlv["_used"]]
        .drop(columns=["_used"])
        .reset_index(drop=True)
    )

    return matched_df, unmatched_rlv_df


# ===========================================================================
# 4. EXPORT EXCEL (2 ONGLETS)
# ===========================================================================
def generer_excel(matched_df: pd.DataFrame, unmatched_df: pd.DataFrame) -> bytes:
    """Construit un fichier Excel en mémoire avec 2 onglets distincts."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="dd/mm/yyyy") as writer:
        matched_df.to_excel(writer, sheet_name="Matched_Results", index=False)
        unmatched_df.to_excel(writer, sheet_name="Unmatched_RLV", index=False)

        # Ajustement automatique de la largeur des colonnes
        for sheet_name, df in [("Matched_Results", matched_df),
                               ("Unmatched_RLV", unmatched_df)]:
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns, start=1):
                largeur = max(
                    len(str(col)),
                    df[col].astype(str).map(len).max() if not df.empty else 10,
                )
                ws.column_dimensions[
                    ws.cell(row=1, column=i).column_letter
                ].width = min(largeur + 3, 45)

    buffer.seek(0)
    return buffer.getvalue()


# ===========================================================================
# 5. INTERFACE STREAMLIT
# ===========================================================================
def main():
    st.title("🏦 Rapprochement Bancaire Automatisé")
    st.caption(
        "Algorithme Many-to-One • Agrégation ERP par *Reference* • "
        "Condition temporelle *Weekend-Aware* • Tolérance ≤ 700 Dhs"
    )

    with st.expander("ℹ️ Règles de l'algorithme", expanded=False):
        st.markdown(
            """
            1. **Agrégation ERP** : regroupement des lignes par `Reference`, somme des montants nets (`Debit − Credit`).
            2. **Condition temporelle** :
               - `Date_ERP` du **Lundi au Jeudi** → `Date_RLV = Date_ERP + 1 jour`.
               - `Date_ERP` le **Vendredi / Samedi / Dimanche** → `Date_RLV` acceptée le **Lundi ou Mardi** suivant.
            3. **Tolérance** : correspondance validée si `|Montant_RLV − Somme_ERP| ≤ 700 Dhs`.
            4. **Écart** = `Montant_RLV − Somme_ERP`.
            """
        )

    # --- Chargement -------------------------------------------------------
    fichier = st.file_uploader(
        "📂 Importer le fichier Excel (onglets **ERP** et **RLV** obligatoires)",
        type=["xlsx", "xlsm"],
    )

    if fichier is None:
        st.info("Veuillez importer un fichier Excel pour commencer.")
        return

    try:
        df_erp, df_rlv = charger_donnees(fichier.getvalue())
    except Exception as e:  # noqa: BLE001
        st.error(f"❌ Erreur de chargement : {e}")
        return

    st.success(
        f"✅ Fichier chargé : **{len(df_erp)}** lignes ERP • "
        f"**{len(df_rlv)}** lignes RLV."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Aperçu ERP**")
        st.dataframe(df_erp.head(8), use_container_width=True, hide_index=True)
    with col2:
        st.markdown("**Aperçu RLV**")
        st.dataframe(df_rlv.head(8), use_container_width=True, hide_index=True)

    st.divider()

    # --- Bouton d'action --------------------------------------------------
    if st.button("🚀 Lancer le Rapprochement", type="primary", use_container_width=True):
        with st.spinner("Rapprochement en cours..."):
            matched_df, unmatched_df = lancer_rapprochement(df_erp, df_rlv)
        st.session_state["matched"] = matched_df
        st.session_state["unmatched"] = unmatched_df

    # --- Affichage des résultats -----------------------------------------
    if "matched" in st.session_state:
        matched_df = st.session_state["matched"]
        unmatched_df = st.session_state["unmatched"]

        # Indicateurs clés (basés sur les correspondances UNIQUES, pas les lignes éclatées)
        nb_corr = matched_df["Match_ID"].nunique() if not matched_df.empty else 0
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("✅ Correspondances", nb_corr)
        m2.metric("⚠️ RLV non rapprochés", len(unmatched_df))
        taux = (nb_corr / len(df_rlv) * 100) if len(df_rlv) else 0
        m3.metric("Taux de rapprochement", f"{taux:.1f} %")
        ecart_total = (
            matched_df["Écart"].dropna().sum() if not matched_df.empty else 0
        )
        m4.metric("Écart cumulé (Dhs)", f"{ecart_total:,.2f}")

        st.divider()

        # Tableau 1 : Matched Results
        st.subheader("📊 Tableau 1 — Correspondances (ERP agrégé ↔ RLV)")
        if matched_df.empty:
            st.warning("Aucune correspondance trouvée avec les règles définies.")
        else:
            st.dataframe(
                matched_df.style.format(
                    {
                        "Montant_Ligne_ERP": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "Somme_ERP": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "Montant_RLV": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "Écart": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                        "Date_ERP": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                        "Date_RLV": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        # Tableau 2 : Unmatched RLV
        st.subheader("🔎 Tableau 2 — Écritures RLV non rapprochées")
        if unmatched_df.empty:
            st.success("Toutes les écritures RLV ont été rapprochées. 🎉")
        else:
            st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

        st.divider()

        # --- Bouton d'exportation ----------------------------------------
        excel_bytes = generer_excel(matched_df, unmatched_df)
        st.download_button(
            label="💾 Télécharger le résultat (Excel — 2 onglets)",
            data=excel_bytes,
            file_name="resultat_rapprochement_bancaire.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
