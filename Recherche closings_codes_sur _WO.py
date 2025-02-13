import pandas as pd
import re
import spacy
import nltk
from collections import Counter
from keybert import KeyBERT
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import CountVectorizer
import pandas as pd
import re
import spacy
import nltk
from keybert import KeyBERT
from nltk.corpus import stopwords

# Télécharger les ressources nécessaires
nltk.download('stopwords')
stop_words = set(nltk.corpus.stopwords.words('french') + nltk.corpus.stopwords.words('english'))

# Charger le modèle NLP français
nlp = spacy.load("fr_core_news_md")

# Charger le fichier Excel contenant les Work Orders
file_path = r"C:\Users\amouaoui\cernbox\WINDOWS\Desktop\Documents stage_CERN\Récuperation des données sous format excel\classifications des WO\Work_Orders_Grouped_By_EquipmentVf.xlsx"
df = pd.read_excel(file_path)

# Vérifier la colonne contenant les descriptions et commentaires
possible_columns = ["COMMENTS", "Description"]
for col in possible_columns:
    if col in df.columns:
        df[col] = df[col].astype(str)

# Fusionner les commentaires et descriptions
if "COMMENTS" in df.columns and "Description" in df.columns:
    df["combined_text"] = df["COMMENTS"] + " " + df["Description"]
elif "COMMENTS" in df.columns:
    df["combined_text"] = df["COMMENTS"]
elif "Description" in df.columns:
    df["combined_text"] = df["Description"]
else:
    raise KeyError("Aucune colonne pertinente ('COMMENTS' ou 'Description') trouvée dans le fichier.")

# Nettoyage avancé du texte
def clean_text(text):
    text = text.lower()
    text = re.sub(r"\d+", "", text)  # Supprimer les chiffres
    text = re.sub(r"[^\w\s]", "", text)  # Supprimer la ponctuation
    text = " ".join([word for word in text.split() if word not in stop_words])  # Supprimer les stopwords
    return text

df["cleaned_text"] = df["combined_text"].apply(clean_text)

# Appliquer KeyBERT pour extraire des mots-clés
kw_model = KeyBERT()
df["keywords"] = df["cleaned_text"].apply(lambda x: kw_model.extract_keywords(x, keyphrase_ngram_range=(1,3), stop_words=stop_words, top_n=5))

# Détection des entités nommées (NER) avec spaCy
def extract_entities(text):
    doc = nlp(text)
    failures, causes, actions = [], [], []

    for ent in doc.ents:
        if ent.label_ in ["MISC", "PRODUCT"]:  # Détection des équipements/pannes
            failures.append(ent.text)
        elif ent.label_ in ["CAUSE", "EVENT"]:  # Détection des causes
            causes.append(ent.text)
        elif ent.label_ in ["ACTION", "WORK_OF_ART"]:  # Détection des actions
            actions.append(ent.text)

    return ", ".join(set(failures)), ", ".join(set(causes)), ", ".join(set(actions))

df[["Failure Code", "Cause Code", "Action Code"]] = df["cleaned_text"].apply(lambda x: pd.Series(extract_entities(x)))

# Sauvegarde du fichier Excel avec les résultats
output_file = "structured_closing_codes.xlsx"
df.to_excel(output_file, index=False)

print(f"✅ Closing Codes extraits et enregistrés dans : {output_file}")
