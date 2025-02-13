import pandas as pd

# 🔹 Remplace par le chemin de ton fichier Excel contenant les Work Orders
file_path = r"C:\Users\amouaoui\cernbox\WINDOWS\Desktop\Documents stage_CERN\Récuperation des données sous format excel\RQF3040186_EL_Comments_V02_WO_With_comments.xlsx"


# Charger le fichier Excel
df = pd.read_excel(file_path)

# 🔹 Définition de la colonne qui contient l'identifiant des équipements
equipment_column = "Equipement"

# Vérifier que la colonne existe
if equipment_column not in df.columns:
    raise ValueError(f"La colonne '{equipment_column}' est introuvable dans le fichier.")

# Fonction pour extraire la troisième lettre
def extract_third_letter(equipment_id):
    if isinstance(equipment_id, str) and len(equipment_id) >= 3:
        return equipment_id[2]  # Troisième caractère (index 2 en Python)
    return None

# Ajouter une colonne temporaire pour le tri
df["Equipment_Type"] = df[equipment_column].apply(extract_third_letter)

# Filtrer les lignes où la troisième lettre est valide (non null)
df = df.dropna(subset=["Equipment_Type"])

# Séparer les Work Orders en plusieurs groupes selon la 3e lettre
grouped = {key: value for key, value in df.groupby("Equipment_Type")}

# 🔹 Nom du fichier de sortie
output_path = "Work_Orders_Classified.xlsx"

# Sauvegarder dans un fichier Excel avec plusieurs feuilles
with pd.ExcelWriter(output_path) as writer:
    for key, data in grouped.items():
        if key:  # Vérifie que la clé n'est pas vide
            data.drop(columns=["Equipment_Type"], inplace=True)  # Supprime la colonne temporaire
            data.to_excel(writer, sheet_name=f"Type_{key}", index=False)

print(f"Fichier enregistré : {output_path}")
