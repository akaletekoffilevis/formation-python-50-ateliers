# Python Cheat Sheet - Référence Rapide

## Syntaxe de Base
```python
# Commentaire sur une ligne
"""
Commentaire
sur plusieurs lignes
"""

# Variables
x = 10               # entier
y = 3.14             # flottant
nom = "Alice"        # chaîne
est_actif = True     # booléen

# Opérateurs
+ - * / // % **      # arithmétiques
== != < > <= <=      # comparaison
and or not           # logique
+= -= *= /=          # assignment composé
```

## Structures de Contrôle
```python
# Conditionnel
if condition:
    # code
elif autre_condition:
    # code
else:
    # code

# Boucle bornée
for i in range(10):          # 0 à 9
    print(i)
    
for element in ma_liste:
    print(element)

# Boucle conditionnelle
while condition:
    # code
    if condition_de_sortie:
        break
    if condition_de_skip:
        continue
```

## Fonctions
```python
def nom_fonction(param1, param2=valeur_defaut):
    """
    Docstring décrivant la fonction
    """
    return resultat

# Appel
result = nom_fonction(arg1, arg2=valeur)

# Paramètres spéciaux
def fonc(a, b=0, *args, **kwargs):
    # a: obligatoire
    # b: défaut
    # args: tuple des arguments positionnels supplémentaires
    # kwargs: dict des arguments nommés supplémentaires
```

## Structures de Données
### Listes
```python
lst = [1, 2, 3]
lst[0]           # premier élément
lst[-1]          # dernier élément
lst[1:4]         # slice (indices 1,2,3)
lst.append(x)    # ajouter à la fin
lst.insert(i, x) # insérer à l'index i
lst.remove(x)    # supprimer première occurrence de x
lst.pop()        # supprimer et retourner dernier élément
lst.pop(i)       # supprimer et retourner élément à l'index i
lst.sort()       # trier en place
lst.reverse()    # inverser en place
len(lst)         # longueur
x in lst         # test d'appartenance
lst.count(x)     # nombre d'occurrences de x
```

### Dictionnaires
```python
d = {"cle": valeur, "nom": "Alice"}
d["cle"]         # accès à la valeur
d.get("cle", default)  # accès avec valeur par défaut
d["nouvelle_cle"] = valeur  # ajouter/modifier
del d["cle"]     # supprimer une clé
cle in d         # test d'existence de clé
d.keys()         # vue des clés
d.values()       # vue des valeurs
d.items()        # vue des paires (clé, valeur)
d.update(other_dict)  # fusionner avec autre dict
```

### Ensembles (Sets)
```python
s = {1, 2, 3}
s.add(x)         # ajouter un élément
s.remove(x)      # retirer un élément (erreur si absent)
s.discard(x)     # retirer un élément (pas d'erreur)
s.pop()          # retirer et retourner un élément arbitraire
x in s           # test d'appartenance
s1 | s2          # union
s1 & s2          # intersection
s1 - s2          # différence
s1 ^ s2          # différence symétrique
```

### Tuples (immuables)
```python
t = (1, 2, 3)
t[0]             # accès
a, b, c = t      # dépaquetage
```

## Chaînes de Caractères
```python
s = "Bonjour le monde"
s[0]             # premier caractère
s[-1]            # dernier caractère
s[1:5]           # slice
s.upper()        # en majuscules
s.lower()        # en minuscules
s.title()        # première lettre de chaque mot en majuscule
s.strip()        # supprimer espaces début/fin
s.split(" ")     # séparation en liste selon séparateur
"-".join(liste)  # joint liste avec séparateur
len(s)           # longueur
s.startswith("Bon")  # teste début
s.endswith("e")      # teste fin
s.find("monde")      # retourne index ou -1
s.replace("ancien", "nouveau")
```

## Formatage de Chaînes
```python
# f-strings (Python 3.6+)
nom = "Alice"
age = 25
f"Je m'appelle {nom} et j'ai {age} ans"

# format()
"Je m'appelle {0} et j'ai {1} ans".format(nom, age)
"Je m'appelle {nom} et j'ai {age} ans".format(nom=nom, age=age)

# % ancien (à éviter)
"Je m'appelle %s et j'ai %d ans" % (nom, age)
```

## Fichiers
```python
# Lecture
with open("fichier.txt", "r", encoding="utf-8") as f:
    contenu = f.read()          # tout le fichier
    ligne = f.readline()        # une ligne
    lignes = f.readlines()      # toutes les lignes dans une liste
    for ligne in f:             # itération ligne par ligne (mémoire efficace)
        print(ligne.strip())

# Écriture
with open("fichier.txt", "w", encoding="utf-8") as f:
    f.write("texte à écrire")
    f.writelines(["ligne1\n", "ligne2\n"])

# Ajout
with open("fichier.txt", "a", encoding="utf-8") as f:
    f.write("texte à ajouter")
```

## Gestion des Erreurs
```python
try:
    # code susceptible de lever une exception
    resultat = 10 / 0
except ZeroDivisionError as e:
    print(f"Erreur de division par zéro: {e}")
except (ValueError, TypeError) as e:
    print(f"Erreur de valeur ou type: {e}")
except Exception as e:  # attrape tout (à utiliser avec précaution)
    print(f"Erreur inattendue: {e}")
else:
    # exécuté si aucune exception n'a été levée
    print("Tout s'est bien passé")
finally:
    # toujours exécuté, idéal pour nettoyage
    print("Nettoyage des ressources")
    
# Lever une exception
raise ValueError("Message d'erreur personnalisé")
```

## Modules et Import
```python
import math                     # importe tout le module
from math import sqrt, pi      # importe des éléments spécifiques
from math import *              # importe tout (à éviter)
import math as m               # alias
from . import module_local      # import relatif (dans un package)

# Utilisation
resultat = math.sqrt(16)
```

## Programmation Orientée Objet
```python
class MaClasse:
    """Docstring de la classe"""
    
    # Attribut de classe
    attribut_classe = "valeur partagée"
    
    def __init__(self, param1, param2="defaut"):
        """Constructeur appelé lors de la création d'instance"""
        self.attribut_instance = param1
        self.autre_attribut = param2
        
    def methode(self, arg):
        """Méthode d'instance"""
        return self.attribut_instance + arg
        
    @classmethod
    def methode_de_classe(cls):
        """Méthode de classe (reçoit la classe en premier argument)"""
        return cls.attribut_classe
        
    @staticmethod
    def methode_statique():
        """Méthode statique (ne reçoit ni self ni cls)"""
        return "Ne dépend pas de l'instance"
        
    def __str__(self):
        """Représentation lisible de l'objet"""
        return f"MaClasse({self.attribut_instance})"
        
    def __repr__(self):
        """Représentation technique de l'objet"""
        return f"MaClasse({self.attribut_instance!r})"
        
    # Opérateurs personnalisés
    def __add__(self, other):
        return MaClasse(self.attribut_instance + other.attribut_instance)
```

## Décorateurs
```python
def mon_decorateur(func):
    def wrapper(*args, **kwargs):
        # code avant l'appel de func
        resultat = func(*args, **kwargs)
        # code après l'appel de func
        return resultat
    return wrapper

@mon_decorateur
def ma_fonction():
    return "Hello"
```

## Gestionnaire de Contexte (with)
```python
class MonContexte:
    def __enter__(self):
        # code d'initialisation
        return self  # ce qui sera lié à 'as'
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        # code de nettoyage
        # retourner True pour supprimer l'exception, False pour la propager
        return False

with MonContexte() as ctx:
    # code utilisant ctx
    pass
```

## Expressions Avancées
### List Comprehensions
```python
# [expression for item in iterable if condition]
carres = [x**2 for x in range(10)]                    # [0,1,4,9,16,25,36,49,64,81]
pairs = [x for x in range(10) if x % 2 == 0]          # [0,2,4,6,8]
pairs_au_carre = [x**2 for x in range(10) if x % 2 == 0]  # [0,4,16,36,64]
```

### Dict Comprehensions
```python
# {key_expr: value_expr for item in iterable if condition}
carres_dict = {x: x**2 for x in range(5)}  # {0:0, 1:1, 2:4, 3:9, 4:16}
```

### Set Comprehensions
```python
# {expression for item in iterable if condition}
voisins_du_chr = {ord(c) for c in "hello" if c != "l"}  # {104, 101, 111}
```

### Expressions Génératrices (paresseuses)
```python
# (expression for item in iterable if condition)  # retourne un générateur
somme = sum(x**2 for x in range(1000000))  # efficace mémoire
```

### Fonctions Lambda
```python
# lambda arguments: expression
double = lambda x: x * 2
# Équivalent à:
def double(x):
    return x * 2

# Utilisation avec map/filter/sorted
liste = [1, 5, 2, 8, 3]
pairs = list(filter(lambda x: x % 2 == 0, liste))
doubles = list(map(lambda x: x * 2, liste))
trie = sorted(liste, key=lambda x: -x)  # ordre décroissant
```

## Bibliothèques Standards Utiles
```python
import random           # nombres aléatoires
random.randint(1, 6)    # entier entre 1 et 6 inclus
random.choice([1,2,3])  # élément aléatoire d'une liste
random.shuffle(liste)   # mélange une liste sur place

import datetime         # dates et heures
now = datetime.datetime.now()
aujourd'hui = datetime.date.today()
delta = datetime.timedelta(days=1)
hier = aujourd'hui - delta

import json             # manipulation JSON
chaine_json = json.dumps(obj, indent=2)
obj = json.loads(chaine_json)

import math             # fonctions mathématiques
math.sqrt(16)           # 4.0
math.pi                 # 3.14159...
math.factorial(5)       # 120

import os               # interaction avec le système d'exploitation
os.getcwd()             # répertoire courant
os.listdir(".")         # liste des fichiers du répertoire
os.path.join("a", "b")  # "a/b" (indépendant du OS)
os.path.exists("fichier")  # teste existence
os.makedirs("dossier/sous-dossier", exist_ok=True)  # crée arbre de dossiers

import sys              # interaction avec l'interpréteur
sys.version             # version de Python
sys.argv                # arguments de ligne de commande
sys.exit()              # quitter le programme
```

## Bonnes Pratiques (PEP 8 Extraits)
- Indentation : 4 espaces (jamais de tabulations)
- Longueur maximale de ligne : 79 caractères
- Espaces autour des opérateurs : `x = 1 + 2`
- Pas d'espaces après `(` ou avant `)`: `func(arg1, arg2)`
- Nommage :
  - fonctions et variables : `snake_case`
  - constantes : `MAJUSCULE_SNAKE_CASE`
  - classes : `PascalCase`
- Une importation par ligne
- Les imports en haut du fichier, dans l'ordre :
  1. imports standards
  2. imports tiers
  3. imports locaux
- Docstrings pour tous les modules, fonctions, classes et méthodes publiques
- Utiliser `is None` / `is not None` plutôt que `== None`
- Utiliser `if not liste:` plutôt que `if len(liste) == 0:` pour tester la vacuité