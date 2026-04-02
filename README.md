# LaboSoft Pro

Catalogue interactif des logiciels disponibles dans les laboratoires informatiques de l'UQAM.

## Fonctionnalites

- **Navigation par laboratoire** : Arts, Communication, Sciences, Education, Gestion, Sciences humaines & Langues
- **Recherche en temps reel** : filtrage instantane par nom de logiciel
- **Filtre par type de licence** : Open Source, Proprietaire, Gratuit, Freemium
- **Interface neon** : theme sombre avec accents de couleur par laboratoire
- **Statistiques** : compteur de logiciels affiches avec repartition par licence

## Types de licence

| Badge   | Description                                         |
|---------|-----------------------------------------------------|
| **OS**    | Open Source - code source libre                   |
| **PROP**  | Proprietaire - licence commerciale requise        |
| **FREE**  | Gratuit - pas open source, mais disponible sans frais |
| **FREEM** | Freemium - version gratuite limitee, payant pour les fonctionnalites completes |

## Pre-requis

- Python 3.8+
- tkinter (inclus avec Python sur la plupart des systemes)
- openpyxl (pour l'import/export Excel)

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
python labosoft.py
```

## Import / Export Excel

L'application permet de charger et exporter les donnees des laboratoires via des fichiers Excel (`.xlsx`).

### Format du fichier Excel

- **Chaque feuille (onglet)** du fichier = un laboratoire
- **Colonnes** (ligne 1 = en-tete) :
  - `Section` : categorie du logiciel (ex: "PC Windows 11", "MAC OS")
  - `Logiciel` : nom du logiciel (ex: "Photoshop", "Audacity")
  - `Licence` : type de licence (`OS`, `PROP`, `FREE`, ou `FREEM`)

### Utilisation

1. **Exporter** : Cliquez sur `EXPORTER EXCEL` pour generer un fichier avec les donnees actuelles
2. **Modifier** : Ouvrez le fichier `.xlsx` dans Excel, ajoutez/modifiez les logiciels
3. **Importer** : Cliquez sur `IMPORTER EXCEL` pour charger le fichier modifie

Un fichier template (`labosoft_template.xlsx`) est inclus comme exemple.

## Laboratoires couverts

- **ART** - Laboratoire Arts (Adobe CC, Maya, Rhino, etc.)
- **COM** - Laboratoire Communication (Ableton, Avid, DaVinci, etc.)
- **SCI** - Laboratoire Sciences (Chimie, Informatique, Unix, Mathematique)
- **EDU** - Sciences de l'education (SPSS, NetLogo, Twine, etc.)
- **GEST** - Sciences de la gestion (SAP, Tableau, MATLAB, etc.)
- **SHL** - Sciences humaines & Langues (SPSS, ArcGIS, Gephi, etc.)

## Licence

Ce projet est distribue sous licence MIT. Voir le fichier [LICENSE](LICENSE).
