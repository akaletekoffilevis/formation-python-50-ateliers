# Formation Python - 50 Ateliers

Ce dépôt contient un cours complet de programmation Python organisé en 50 présentations PowerPoint, conçu pour apprendre Python du niveau débutant au niveau avancé.

## Structure du dépôt

```
.
├── presentations/               # Toutes les 50 présentations PowerPoint
│   ├── 01_Cest_quoi_informatique.pptx
│   ├── 02_Fonctionnement_ordinateur.pptx
│   ├── ... 
│   └── 50_Prochaines_etapes.pptx
│   └── apprendre_python_50_presentations.zip  # Archive ZIP de toutes les présentations
├── livraison/                   # Dossiers pour les projets étudiants à rendre
│   ├── 01_calculatrice_interactive/
│   │   └── README.md
│   ├── 02_gestionnaire_taches/
│   │   └── README.md
│   ├── 03_jeu_cartes/
│   │   └── README.md
│   ├── 04_pendu/
│   │   └── README.md
│   ├── 05_meteo_app/
│   │   └── README.md
│   ├── 06_blog_flask/
│   │   └── README.md
│   └── 07_projet_final/
│       └── README.md
├── generate_*.py               # Scripts pour générer les présentations (exclu de git)
└── README.md                   # Ce fichier
```

## Contenu du cours

Les 50 présentations couvrent :

### Partie 1 : Fondamentaux de l'informatique
- Présentation de l'informatique et son histoire
- Fonctionnement d'un ordinateur (CPU, RAM, stockage)
- Systèmes d'exploitation
- Internet et le Web
- Bases de la programmation et des algorithmes

### Partie 2 : Introduction à Python
- Installation de Python et VS Code
- Utilisation du terminal/ligne de commande
- Premier programme : "Hello World"
- Variables, types de données, chaînes de caractères
- Opérations mathématiques
- Entrée/sortie avec input()/print()

### Partie 3 : Structures de contrôle
- Instructions conditionnelles (if/elif/else)
- Boucles while et for
- Exercices pratiques sur les conditions et les boucles

### Partie 4 : Structures de données
- Listes : manipulation, méthodes, slicing
- Dictionnaires : création, accès, méthodes
- Tuples et ensembles (sets)
- Exercices combinant listes, boucles et dictionnaires

### Partie 5 : Programmation procédurale
- Définition et utilisation de fonctions
- Paramètres des fonctions (valeur par défaut, *args, **kwargs)
- Portée des variables (variables locales et globales)

### Partie 6 : Gestion des erreurs et fichiers
- Gestion des exceptions avec try/except/else/finally
- Lecture et écriture de fichiers
- Utilisation du contexte `with`
- Module pathlib pour la manipulation de chemins

### Partie 7 : Modules et organisation du code
- Création et utilisation de modules
- Instructions d'import (`import`, `from ... import`)
- Blocs `if __name__ == "__main__"`
- Organisation de projets Python

### Partie 8 : Programmation Orientée Objet (POO)
- Classes et objets
- Constructeur `__init__`
- Attributs et méthodes d'instance
- Héritage et polymorphisme
- Encapsulation et propriétés

### Partie 9 : Expressions avancées
- List comprehensions, dict comprehensions, set comprehensions
- Fonctions lambda
- Fonctions map(), filter(), reduce()
- Tri personnalisé avec sorted() et key=

### Partie 10 : Qualité du code et bonnes pratiques
- Gestion avancée des exceptions
- Débogage avec breakpoint() et assertions
- Tests unitaires avec unittest et pytest
- Bonnes pratiques Python (PEP 8, DRY, KISS)
- Documentation avec les docstrings

### Partie 11 : Projets pratiques et applications réelles
- Bibliothèques utiles : random, datetime, math, os, sys, json
- Mini-projets : Gestionnaire de tâches, Jeu de cartes, Le pendu
- Introduction à Git et GitHub
- Environnements virtuels et gestion des dépendances avec pip
- APIs HTTP avec le module requests
- Bases de données avec SQLite
- Développement web avec Flask
- Projet final : application complète à réaliser

### Partie 12 : Prochaine étapes
- Ressources pour continuer l'apprentissage
- Communautés de développeurs Python
- Suggestions de spécialisations (Data Science, développement web, jeux vidéo, etc.)

## Utilisation

### Visualiser les présentations
Les fichiers PowerPoint sont disponibles dans le dossier `presentations/`. Vous pouvez les ouvrir avec :
- Microsoft PowerPoint
- LibreOffice Impress
- Google Slides (via importation)
- Toute autre application compatible avec le format .pptx

Une archive ZIP contenant toutes les présentations est également disponible :
`presentations/apprendre_python_50_presentations.zip`

### Générer les présentations soi-même
Les scripts `generate_*.py` permettent de régénérer les présentations. Ils nécessitent :
- Python 3.x
- La bibliothèque `python-pptx` (installable avec `pip install python-pptx`)

**Note** : Ces scripts sont volontairement exclus du suivi Git pour éviter de surcharger le dépôt avec du code générateur plutôt que le contenu pédagogique lui-même.

### Projets étudiants
Le dossier `livraison/` contient des structures préparées pour que les étudiants puissent rendre leurs projets étape par étape. Chaque sous-dossier correspond à un projet spécifique du cours et contient un README détaillé avec :
- Les objectifs du projet
- Les fonctionnalités requises
- Les concepts à appliquer
- Les fichiers attendus
- Des suggestions d'extensions

## Licence

Ce projet est sous licence MIT - voir le fichier `LICENSE` pour plus de détails.

## Auteur

Créé avec ❤️ pour faciliter l'apprentissage de la programmation Python.

---
*Dernière mise à jour : Mai 2026*