import streamlit as st
import pandas as pd

# Charger les données
url = r"https://docs.google.com/spreadsheets/d/e/2PACX-1vRKnPCZpn7PKhzp0rQ69Ug-K5GeTvSylN9aBX-XpgbgCj-Jd4tL1f8jP-0phT-8UQ/pub?output=csv"
df = pd.read_csv(url)


# Nettoyer les données (remplir les NaN pour les colonnes fusionnées)
df["Problème"].fillna(method="ffill", inplace=True)

def get_defaillances(probleme):
    return df[df["Problème"] == probleme]["Code de Défaillance"].unique()

def get_causes():
    return df["Causes Codes"].dropna().unique()

def get_actions():
    return df["Action Codes"].dropna().unique()

# Interface Streamlit
st.title("🔧 Sélection des Closing Codes")

# Sélection du problème
probleme = st.selectbox("🟢 Sélectionnez un Probleme Code :", df["Problème"].unique())

defaillances = get_defaillances(probleme)
defaillance = st.selectbox("⚠️ Sélectionnez un Failure Code :", defaillances)

causes = get_causes()
cause = st.selectbox("❗ Sélectionnez un Cause Code :", causes)

actions = get_actions()
action = st.selectbox("🛠️ Sélectionnez un Action Code :", actions)

# Affichage du récapitulatif
st.subheader("📜 Récapitulatif de la sélection")
st.markdown(f"**Problème :** {probleme}")
st.markdown(f"**Défaillance :** {defaillance}")
st.markdown(f"**Cause :** {cause}")
st.markdown(f"**Action :** {action}")

# Sauvegarde en Excel
data_export = pd.DataFrame({
    "Problème": [probleme],
    "Défaillance": [defaillance],
    "Cause": [cause],
    "Action": [action]
})

if st.button("📥 Sauvegarder en Excel"):
    file_name = "closing_code_selection.xlsx"
    data_export.to_excel(file_name, index=False)
    st.success(f"Fichier enregistré sous {file_name}")
