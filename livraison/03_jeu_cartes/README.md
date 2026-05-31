# Projet : Jeu de Cartes (Bataille)

## Objectif
Implémenter le jeu classique de la "Bataille" en utilisant la Programmation Orientée Objet (POO).

## Fonctionnalités requises
- Deux joueurs (humain vs ordinateur ou deux humains)
- Un paquet standard de 52 cartes
- Distribution aléatoire des cartes
- Chaque tour : chaque joueur révèle une carte, la plus forte gagne
- En cas d'égalité : "bataille" (cartes supplémentaires mises en jeu)
- Le jeu se termine quand un joueur n'a plus de cartes
- Affichage du gagnant à la fin

## Concepts POO à appliquer
- Classes : Carte, Paquet, Joueur, Jeu
- Attributs et méthodes d'instance
- Constructeur (__init__)
- Méthodes spéciales (__str__, __lt__ pour comparer les cartes)
- Encapsulation (attributs privés si nécessaire)
- Possiblement héritage (pour différents types de joueurs)

## Fichiers attendus
- `carte.py` : Classe Carte
- `paquet.py` : Classe Paquet
- `joueur.py` : Classe Joueur
- `jeu.py` : Classe Jeu (logique principale)
- `main.py` : Point d'entrée du jeu
- `README.md` : Ce fichier

## Extensions possibles
- Interface graphique simple (avec tkinter)
- Règles de variante (Bataille avec plusieurs paquets)
- Statistiques (nombre de tours, guerres, etc.)
- Option pour rejouer