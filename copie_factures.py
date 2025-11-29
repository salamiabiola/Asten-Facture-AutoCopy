import os
import shutil
import pandas as pd

# ----------------------------------------
# CONFIGURATION CHEMINS
# ----------------------------------------
fichier_excel = r"C:\Users\Mon ordi\Documents\FACT_MJ\VERIFICATION_FACT_ASTEN.xlsx"
dossier_source = r"C:\Users\Mon ordi\Documents\FACT_MJ\Fact-Backup"
dossier_destination = r"C:\Users\Mon ordi\Documents\FACT_MJ\Fact Non Int Trouvees\Fact_No_Int"

os.makedirs(dossier_destination, exist_ok=True)

# ----------------------------------------
# LECTURE DE LA FEUILLE EXCEL
# ----------------------------------------
df = pd.read_excel(
    fichier_excel,
    sheet_name="Fact_Non_Integrées_Sur_ASTEN",
    header=3
)

df["Date.Validation Backup"] = pd.to_datetime(df["Date.Validation Backup"], errors="coerce")

# ----------------------------------------
# SELECTION DU MOIS
# ----------------------------------------
print("\n===== SELECTION DU MOIS DES FACTURES A TROUVER =====")
print("1 - Un seul mois")
print("2 - Plusieurs mois")
print("3 - Période (mois début -> mois fin)")

choix = input("\nChoisissez une option (1/2/3) : ")

if choix == "1":
    mois = int(input("\nEntrer un mois (1-12) : "))
    df_filtre = df[df["Date.Validation Backup"].dt.month == mois]
    print(f"\nMois sélectionné : {mois}")

elif choix == "2":
    mois_liste = input("\nEntrer les mois séparés par virgule (ex : 9,10,11) : ")
    mois_liste = [int(m.strip()) for m in mois_liste.split(",")]
    df_filtre = df[df["Date.Validation Backup"].dt.month.isin(mois_liste)]
    print(f"\nMois sélectionnés : {mois_liste}")

elif choix == "3":
    debut = int(input("\nMois début : "))
    fin   = int(input("Mois fin   : "))
    df_filtre = df[df["Date.Validation Backup"].dt.month.between(debut, fin)]
    print(f"\nPériode sélectionnée : {debut} -> {fin}")

else:
    print("Option invalide.")
    exit()

# ----------------------------------------
# EXTRACTION DES ENTÊTES
# ----------------------------------------
liste_entetes = df_filtre["En-tete"].dropna().unique()
print(f"\nNombre de factures à traiter : {len(liste_entetes)}")

factures_introuvables = []

# ----------------------------------------
# RECHERCHE ET COPIE
# ----------------------------------------
for entete in liste_entetes:
    trouve = False

    for racine, sous_dossiers, fichiers in os.walk(dossier_source):
        if entete in fichiers:
            shutil.copy2(os.path.join(racine, entete),
                         os.path.join(dossier_destination, entete))
            print("Copié :", entete)
            trouve = True
            break

    if not trouve:
        print("Introuvable :", entete)
        factures_introuvables.append(entete)

# ----------------------------------------
# RECAP FINAL
# ----------------------------------------
print("\n================== FIN DU TRAITEMENT ==================")
print("Factures copiées :", len(liste_entetes) - len(factures_introuvables))
print("Factures introuvables :", len(factures_introuvables))

if factures_introuvables:
    print("\nListe des introuvables :")
    for f in factures_introuvables:
        print(" -", f)

print("========================================================")
