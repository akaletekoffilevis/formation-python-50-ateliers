# Projet : Application Météo

## Objectif
Développer une application console qui récupère et affiche les données météorologiques en temps réel pour une ville donnée en utilisant une API publique.

## Fonctionnalités requises
- Demander le nom d'une ville à l'utilisateur
- Appeler une API météorologique (ex: OpenWeatherMap)
- Afficher la température actuelle, l'humidité, la description du temps
- Afficher une prévision simple sur les prochains jours
- Gérer les erreurs (ville introuvable, problème de connexion, etc.)
- Optionnel : sauvegarder l'historique des recherches

## Concepts à appliquer
- Requêtes HTTP avec le module `requests`
- Manipulation de données JSON
- Gestion d'erreurs (try/except)
- Fichiers de configuration (optionnel)
- Variables d'environnement pour la clé API (optionnel)

## Fichiers attendus
- `meteo.py` : Script principal
- `config.py` : Configuration et gestion de la clé API (optionnel)
- `README.md` : Ce fichier
- `requirements.txt` : Liste des dépendances (doit contenir `requests`)

## Extensions possibles
- Interface graphique simple
- Sauvegarde des favoris
- Affichage graphique des températures
- Alertes météo
- Support de plusieurs unités (Celsius/Fahrenheit)