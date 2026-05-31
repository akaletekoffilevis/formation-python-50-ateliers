# Projet : Blog avec Flask

## Objectif
Créer une application web simple de blog utilisant le framework Flask et une base de données SQLite.

## Fonctionnalités requises
- Afficher la liste des articles de blog sur la page d'accueil
- Permettre de lire un article complet en cliquant dessus
- Permettre de créer un nouvel article via un formulaire
- Permettre de modifier un article existant
- Permettre de supprimer un article
- Stocker les articles dans une base de données SQLite
- Utiliser des templates HTML pour l'affichage

## Concepts à appliquer
- Framework Flask (routes, templates, request handling)
- Base de données SQLite avec le module sqlite3
- Modèle de données pour les articles (titre, contenu, date)
- Opérations CRUD (Create, Read, Update, Delete)
- Templates Jinja2 pour le HTML dynamique
- Formulaires HTML (méthodes GET et POST)

## Fichiers attendus
- `app.py` : Application Flask principale
- `requirements.txt` : Dépendances (doit contenir `flask`)
- Dossier `templates/` : contenant les fichiers HTML
  - `base.html` : template de base
  - `index.html` : liste des articles
  - `article.html` : affichage d'un article
  - `create.html` : formulaire de création/édition
- Dossier `static/` (optionnel) : CSS, images
- `blog.db` : base de données SQLite (générée automatiquement)
- `README.md` : Ce fichier

## Extensions possibles
- Authentification utilisateur (connexion/inscription)
- Catégories ou tags pour les articles
- Commentaires sur les articles
- Date et heure de dernière modification
- Recherche d'articles
- Pagination de la liste des articles