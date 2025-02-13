import pandas as pd

# 🔹 Corrige le chemin du fichier en utilisant une des solutions ci-dessus
file_path = r"C:\Users\amouaoui\cernbox\WINDOWS\Desktop\Documents stage_CERN\Sheet_classification_donnees.xlsx"

# Charger le fichier Excel
df = pd.read_excel(file_path)

# Vérifier que la colonne "Position" existe
if "Position" not in df.columns:
    raise ValueError("La colonne 'Position' est introuvable dans le fichier.")

# Fonction pour extraire la troisième lettre
def extract_third_letter(position):
    if isinstance(position, str) and len(position) >= 3:
        return position[2]  # Troisième caractère (index 2 en Python)
    return None

# Ajouter une colonne temporaire pour le tri
df["Equipment_Type"] = df["Position"].apply(extract_third_letter)

# Filtrer les lignes où la troisième lettre est valide (non null)
df = df.dropna(subset=["Equipment_Type"])

# Séparer les équipements en plusieurs groupes selon la 3e lettre
grouped = {key: value for key, value in df.groupby("Equipment_Type")}

# Sauvegarder dans un fichier Excel avec plusieurs feuilles
output_path = "equipements_par_type.xlsx"
with pd.ExcelWriter(output_path) as writer:
    for key, data in grouped.items():
        if key:  # Vérifie que la clé n'est pas vide
            data.drop(columns=["Equipment_Type"], inplace=True)  # Supprime la colonne temporaire
            data.to_excel(writer, sheet_name=f"Type_{key}", index=False)

print(f"Fichier enregistré : {output_path}")
