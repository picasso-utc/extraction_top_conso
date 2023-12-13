# Exctracteur et exportateur des tops consomation

## Installation

Installer les packages nécessaires à l'utilisation du script:
- pandas
- re
- openpyxl (Dépendance optionel devenue nécessaire)

## Pré-requis

Sur Weez il faut demander et télécharger un export contenant au moins ces attributs sur une période donnée : 
- Prénom acheteur
- Nom acheteur
- Article
- Famille d'article
- Quantité
- Total TTC

## Utilisation

```
python main.py path_to_export save_path
```
Où **path_to_export** est le chemin pour accéder au fichier d'export (.xlsx) et **save_path** le chemin du dossier où enregistrement l'exportation.

## Auteurs

- Luca Rougeron (https://github.com/luc9a)
- Maxime Vaillant (https://github.com/maxime-vaillant)
