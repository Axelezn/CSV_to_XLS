# CSV_to_XLS : Convertisseur de CSV en Excel 📊

Un outil autonome conçu pour **convertir et fusionner plusieurs fichiers CSV en un seul fichier Excel (.xlsx)** sous Windows.

---

## ✨ Fonctionnalités Clés

* **Interface Graphique (GUI)** : Utilise des boîtes de dialogue pour sélectionner facilement le dossier source et nommer le fichier de sortie, rendant l'outil accessible à tous.
* **Renommage Automatique** : Renomme les fichiers CSV à noms longs (ex: `ID_MOIS_ANNEE.csv` en `MOIS_ANNEE.csv`) pour éviter les erreurs de limite de 31 caractères pour les noms de feuilles Excel.
* **Double Mode de Fusion** : Offre un choix entre deux méthodes d'organisation des données.

---

## 🚀 Utilisation (Pour l'Utilisateur Final)

L'outil ne nécessite **aucune installation de Python ni de dépendances**.

1.  **Extraction** : Extrayez le fichier `csv_excel.exe` de son archive.
2.  **Lancement** : Double-cliquez sur l'exécutable **`csv_excel.exe`**.

### Processus Interactif 🗣️

Le programme vous guidera à travers les étapes suivantes :

1.  **Instructions Console** : Le terminal s'ouvrira, affichant les consignes initiales.
2.  **Sélection du Dossier** : Une fenêtre de dialogue apparaîtra (avec l'indication "Choisissez un dossier...") pour que vous sélectionniez le dossier contenant les fichiers CSV.
3.  **Nom du Fichier de Sortie** : Une seconde fenêtre vous demandera de nommer le fichier Excel final.
4.  **Choix du Mode de Fusion** : La console vous présentera le menu pour choisir le mode :

| Choix | Mode | Résultat dans Excel |
| :---: | :--- | :--- |
| **1** | Multi-Pages | Chaque fichier CSV sera placé sur une feuille séparée (nommée d'après le CSV, ex: `JANUARY_2024`). |
| **2** | Concaténation | Tous les fichiers CSV seront fusionnés et empilés dans une seule feuille nommée `Fusion_Totale`. |

---

## 💻 Développement et Compilation

Cette application a été compilée à partir d'un script Python, permettant à l'utilisateur d'exécuter le programme sans besoin d'installation supplémentaire.

### Outil de Compilation

L'outil utilisé pour cette transformation est **PyInstaller**. Toutes les dépendances sont incluses dans le fichier unique `csv_excel.exe`.

### Commande de Compilation

Le programme a été créé en utilisant la commande suivante, garantissant une distribution simple et portable :

```bash
python -m PyInstaller --onefile --name "csv_excel" conversion.py
