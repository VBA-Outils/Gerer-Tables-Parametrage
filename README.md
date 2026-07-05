# 📊 GererTables — Gestion avancée de tables Excel via VBA

![Langage](https://img.shields.io/badge/langage-VBA-blue)
![Excel](https://img.shields.io/badge/Excel-Tables%20structurées-green)
![Licence](https://img.shields.io/badge/Licence-MIT-yellow)

GererTables est un outil Excel/VBA permettant de gérer des **tables structurées** réparties par onglet, avec un **contrat d’interface centralisé**, et des fonctions d’**import/export CSV** ainsi que la génération de **requêtes SQL `INSERT INTO`**.

Ce projet vise à offrir une manière fiable, cohérente et automatisée de manipuler des données tabulaires dans Excel.

---

## 🧩 Fonctionnalités principales

### ✔️ Gestion des tables structurées
- Création automatique des tables selon le contrat d’interface  
- Vérification des colonnes et des formats  
- Ajustement automatique des largeurs  
- Validation des données

### ✔️ Import CSV → Table Excel
- Lecture d’un fichier CSV  
- Vérification de la conformité avec le contrat  
- Insertion des données dans la table cible  
- Rapport d’erreurs (colonnes manquantes, formats incorrects)

### ✔️ Export Table Excel → CSV
- Export propre et conforme  
- Gestion des séparateurs  
- Encodage UTF‑8 ou ANSI  
- Option d’inclure ou non les en‑têtes

### ✔️ Export Table Excel → SQL `INSERT INTO`
- Génération automatique des requêtes  
- Échappement des valeurs  
- Gestion des types (texte, numérique, date)  
- Export dans un fichier `.sql`

---

## 📁 Structure du classeur

Le classeur contient plusieurs onglets, chacun dédié à une table structurée.  
Un onglet particulier, **`Contrat d'interface`**, décrit :

- le nom de chaque table,
- les colonnes,
- leur ordre,
- leur type ou format,
- les règles de validation.

Ce contrat sert de **référence unique** pour toutes les opérations d’import/export.

---

## 🏗️ Architecture VBA

| Module / Classe     | Rôle |
|---------------------|------|
| `ContratInterface`  | Lecture et validation du contrat d’interface |
| `GestionTables`     | Création, vérification et manipulation des tables |
| `ImportCSV`         | Import de fichiers CSV |
| `ExportCSV`         | Export de tables vers CSV |
| `ExportSQL`         | Génération des requêtes SQL |
| `Utilitaires`       | Fonctions génériques (fichiers, chaînes, formats…) |

Chaque module est documenté via des commentaires structurés compatibles avec Doxygen/VBDoc.

---

## 🚀 Démarrage rapide

### 1. Télécharger le fichier Excel  
Récupérez `GererTables.xlsm` depuis le dépôt GitHub.

### 2. Activer les macros  
Excel → Options → Centre de gestion de la confidentialité → Paramètres des macros.

### 3. Définir vos tables  
Dans l’onglet **`Contrat d'interface`**, définissez les colonnes et formats.

### 4. Importer un CSV  
Menu → `Importer CSV` → Sélectionner le fichier → Choisir la table cible.

### 5. Exporter un CSV  
Menu → `Exporter CSV` → Choisir la table → Sélectionner l’emplacement.

### 6. Générer du SQL  
Menu → `Exporter SQL` → Un fichier `.sql` est généré.

---

## 🧪 Exemples

### Table structurée

| ID | Nom | Age |
| --- | --- | --- |
| 1 | Dupont | 32 |
| 2 | Martin | 45 |


### SQL généré

INSERT INTO Clients (ID, Nom, Age) VALUES 
(1, 'Dupont', 32),
(2, 'Martin', 45);


