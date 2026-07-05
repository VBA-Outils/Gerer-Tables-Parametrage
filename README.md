# 📊 GererTables — Gestion avancée de tables Excel via VBA

![Langage](https://img.shields.io/badge/langage-VBA-blue)
![Excel](https://img.shields.io/badge/Excel-Tables%20structurées-green)
![Licence](https://img.shields.io/badge/Licence-MIT-yellow)

GererTables est un outil Excel/VBA permettant de gérer des **tables structurées** réparties par onglet, avec un **contrat d’interface centralisé**, et des fonctions d’**import/export CSV** ainsi que la génération de **requêtes SQL `INSERT INTO`**.

Ce projet vise à offrir une manière fiable, cohérente et automatisée de manipuler des données tabulaires dans Excel.

---

## 📑 Table des matières

- [Présentation](#présentation)
- [Installation](#installation)
- [Fonctionnalités](#fonctionnalités)
- [Méthodologie](#méthodologie)
- [Mode d’emploi](#-mode-demploi--création-et-utilisation-des-tables)
  - [1️⃣ Préparer le contrat d’interface](#1️⃣-préparer-longlet-contrat-dinterface)
  - [2️⃣ Générer les tables](#2️⃣-générer-les-tables)
  - [3️⃣ Définir les relations](#3️⃣-définir-les-relations-entre-tables)
- [Licence](#licence)

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

---

## 📘 Mode d’emploi — Création et utilisation des tables

### 1️⃣ Préparer l’onglet **Contrat d’interface**

L’onglet *Contrat d’interface* constitue le schéma fonctionnel du classeur.  
Chaque ligne décrit une table ou une colonne.

Renseigner pour chaque entrée :

- **Nom de la table**
- **Nom de la colonne**
- **Format / Type de données** (Texte, Entier, Date, etc.)
- **Clé primaire** (si applicable)
- **Références** (liens vers d’autres tables)

Cet onglet sert de base à la génération automatique des tables.

---

### 2️⃣ Générer les tables

Une fois le contrat d’interface complété :

1. Cliquer sur le bouton **Créer tables**.
2. Le programme :
   - crée un **onglet par table** définie ;
   - insère un **tableau structuré** avec les colonnes du contrat ;
   - ajoute automatiquement sur la première ligne :
     - **Menu**
     - **Exporter CSV**
     - **Importer CSV**
     - **Exporter SQL**
   - inscrit le **nom de la table** en haut de l’onglet.

Les tables sont immédiatement opérationnelles pour la saisie, l’import/export et la génération SQL.

---

### 3️⃣ Définir les relations entre tables

Après la création des tables :

- La colonne **Table référencée** du contrat d’interface propose une **liste déroulante** contenant toutes les tables du classeur.
- Lorsque l’utilisateur sélectionne une table :
  - la colonne **Clé primaire référencée** se remplit automatiquement avec une **liste déroulante contenant les clés primaires de la table choisie**.

Ce mécanisme permet de définir facilement :

- des **relations 1‑N** (clé primaire → clé étrangère),
- des **références croisées**,
- des **contraintes de cohérence** entre tables.

Le contrat d’interface devient ainsi un véritable *catalogue de métadonnées* pilotant la structure du classeur.

---

## 🧩 Résultat : un système de tables piloté par métadonnées

Grâce à ce fonctionnement :

- la structure des tables est définie **une seule fois** dans le contrat d’interface ;
- les onglets sont générés automatiquement et restent synchronisés ;
- les relations entre tables sont gérées via des listes déroulantes dynamiques ;
- l’utilisateur dispose d’un environnement cohérent pour :
  - saisir des données,
  - importer/exporter en CSV,
  - générer des requêtes SQL,
  - maintenir des relations entre tables.

