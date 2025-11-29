# Facture-AutoCopy

Un script Python permettant d’identifier et de récupérer automatiquement les factures non intégrées sur un logiciel, à partir d’un fichier Excel consolidé.
Le script recherche ensuite ces factures dans un dossier source (avec sous-dossiers), et copie celles trouvées vers un dossier de sortie pour faciliter leur réintégration.

## Fonctionnalités

✔ Lecture automatique d'un fichier Excel consolidé</br>
✔ Filtrage des factures selon un mois spécifique</br>
✔ Recherche des fichiers dans tous les sous-dossiers du backup</br>
✔ Copie automatique des factures trouvées vers un dossier défini</br>
✔ Liste des factures introuvables affichée à la fin</br>
✔ Réduction du temps de recherche manuelle et d’intégration

## STRUCTURE D'UTILISATION

### Configuration Des Chemins

- VERIFICATION_FACT_ASTEN.xlsx                  # Fichier contenant les factures consolidées **à adapter**
- Fact-Backup/                                  # Dossier source contenant les factures (.txt) qui sont dans les sous-dossiers
- Fact Non Int Trouvees/                        # Destination des factures trouvées **à adapter**

### Lecture Excel Avec Bon En-Tete

- Sheet_name="Fact_Non_Integrées_Sur_ASTEN"     # Nom de la Feuille du fichier Excel consolidé **à adapter**

## Prérequis

Assurez-vous que **Python 3.10+** est installé.Si vous ne savez pas comment l’installer, vous pouvez consulter un tutoriel sur YouTube en recherchant : « comment installer Python sur Windows » dans la barre de recherche.

Installer aussi les dépendances nécessaires pour le script.
Dans votre Cmd Taper la commande suivante :
``````bash

pip install pandas openpyxl

``````
## Exécution du script

Dans le script, modifier ces chemins selon votre environnement :
``````python

fichier_excel = r"C:\Chemin\VERS\VERIFICATION_FACT_ASTEN.xlsx"
dossier_source = r"C:\Chemin\VERS\Fact-Backup"
dossier_destination = r"C:\Chemin\VERS\Fact Non Int Trouvees"
``````
Dans le script, modifier aussi le nom de la feuille Excel Consolidé selon votre environnement :
### EX :

ici sheet_name="Fact_Non_Integrées_Sur_ASTEN" ====> sheet_name="Nom_Feuille"
``````python
# ------------------------------
# LECTURE EXCEL AVEC BON EN-TETE
# ------------------------------
df = pd.read_excel(
    fichier_excel, 
    sheet_name="Fact_Non_Integrées_Sur_ASTEN",
    header=3
)
``````
Dans le script, modifier aussi la date des factures a récuperer selon votre environnement :
### EX :
ici ".dt.month == 10"(OCTOBRE) ====> ".dt.month == 11"(NOVEMBRE)
``````python
# ------------------------------
# FILTRER LES FACTURES DE NOVEMBRE
# ------------------------------

df["Date.Validation Backup"] = pd.to_datetime(df["Date.Validation Backup"], errors="coerce")
df_novembre = df[df["Date.Validation Backup"].dt.month == 10]
``````
Puis lancer dans le terminal qui est dans Visual studio :
``````bash
python copie_factures.py
``````
## Résultats obtenus

À la fin de l'exécution :

Les factures trouvées sont copiées automatiquement vers le dossier Fact Non Int Trouvees(Dossier Destination).
Une liste des factures introuvables est affichée dans le terminal de visual studio pour suivi manuel.