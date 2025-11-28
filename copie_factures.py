import os
import shutil
import pandas as pd

# ------------------------------
# CONFIGURATION DES CHEMINS
# ------------------------------
fichier_excel = r"C:\Users\Mon ordi\Documents\FACT_MJ\VERIFICATION_FACT_ASTEN.xlsx"
dossier_source = r"C:\Users\Mon ordi\Documents\FACT_MJ\Fact-Backup"
dossier_destination = r"C:\Users\Mon ordi\Documents\FACT_MJ\Fact Non Int Trouvees\Fact_No_Int"

# Création du dossier destination s'il n'existe pas
os.makedirs(dossier_destination, exist_ok=True)

# ------------------------------
# LECTURE EXCEL AVEC BON EN-TETE
# ------------------------------
df = pd.read_excel(
    fichier_excel, 
    sheet_name="Fact_Non_Integrées_Sur_ASTEN",
    header=3
)

# ------------------------------
# FILTRER LES FACTURES DE NOVEMBRE
# ------------------------------
df["Date.Validation Backup"] = pd.to_datetime(df["Date.Validation Backup"], errors="coerce")
df_novembre = df[df["Date.Validation Backup"].dt.month == 10]

# ------------------------------
# EXTRAIRE LES EN-TETES UNIQUES
# ------------------------------
liste_entetes = df_novembre["En-tete"].dropna().unique()
print(f"Nombre de factures à traiter : {len(liste_entetes)}")

# ------------------------------
# LISTE DES FACTURES INTR
# ------------------------------
factures_introuvables = []

# ------------------------------
# RECHERCHE ET COPIE DES FICHIERS
# ------------------------------
for entete in liste_entetes:
    nom_fichier = entete
    fichier_trouve = False

    for racine, sous_dossiers, fichiers in os.walk(dossier_source):
        if nom_fichier in fichiers:
            chemin_source = os.path.join(racine, nom_fichier)
            chemin_destination = os.path.join(dossier_destination, nom_fichier)

            shutil.copy2(chemin_source, chemin_destination)
            print(f"Copié : {nom_fichier}")
            fichier_trouve = True
            break  # On arrête dès qu'on a trouvé

    if not fichier_trouve:
        factures_introuvables.append(nom_fichier)

# ------------------------------
# AFFICHAGE DES FACTURES INTR
# ------------------------------
nombre_introuvables = len(factures_introuvables)

if factures_introuvables:
    print(f"\nNombre de factures introuvables : {nombre_introuvables}")
    print("--- Liste des factures introuvables ---")
    for f in factures_introuvables:
        print(f)
else:
    print("\nToutes les factures ont été trouvées et copiées !")

print("\n--- Traitement terminé ---")