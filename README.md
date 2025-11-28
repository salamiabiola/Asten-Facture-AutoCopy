**# ------Asten-Facture-AutoCopy-----**

Un script Python permettant d’identifier et de récupérer automatiquement les factures non intégrées sur le logiciel ASTEN, à partir d’un fichier Excel consolidé.
Le script recherche ensuite ces factures dans un dossier source (avec sous-dossiers), et copie celles trouvées vers un dossier de sortie pour faciliter leur réintégration.

# ------Fonctionnalités-----

✔ Lecture automatique d'un fichier Excel consolidé
✔ Filtrage des factures selon un mois spécifique
✔ Recherche des fichiers dans tous les sous-dossiers du backup
✔ Copie automatique des factures trouvées vers un dossier défini
✔ Liste des factures introuvables affichée à la fin
✔ Réduction du temps de recherche manuelle et d’intégration

# ------Structure d’utilisation-----

``````
    ├── VERIFICATION_FACT_ASTEN.xlsx         # Fichier contenant les factures consolidées
├── Fact-Backup/                         # Dossier source contenant les factures (.txt)
│   ├── prdP2A_XXXX/                     # Plusieurs sous-dossiers
│   ├── prdFactureAvoirP2A_XXXX/
│   └── ...
└── Fact Non Int Trouvees/               # Destination des factures trouvées

``````

# ------Prérequis-----

Assurez-vous que Python 3.10+ est installé.

Installer les dépendances nécessaires :

``````bash

pip install pandas openpyxl

``````

# ------Exécution du script-----

Modifier les chemins dans le script selon votre environnement :

``````python

fichier_excel = r"C:\Chemin\VERS\VERIFICATION_FACT_ASTEN.xlsx"
dossier_source = r"C:\Chemin\VERS\Fact-Backup"
dossier_destination = r"C:\Chemin\VERS\Fact Non Int Trouvees"


``````

Puis lancer :

``````bash

python copie_factures.py

``````

# ------Résultats obtenus-----

À la fin de l'exécution :

    - Les factures trouvées sont copiées automatiquement vers le dossier Fact Non Int Trouvees
    - Une liste des factures introuvables est affichée pour suivi manuel