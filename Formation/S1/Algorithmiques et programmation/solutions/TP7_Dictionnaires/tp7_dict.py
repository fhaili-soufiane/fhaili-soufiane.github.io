# TP7 : Les dictionnaires — Solutions

# Ex1 : Trier un dictionnaire par clé / par valeur
def ex1():
    d = {'b': 3, 'a': 1, 'c': 2}
    print("Par clé:", dict(sorted(d.items())))
    print("Par valeur:", dict(sorted(d.items(), key=lambda x: x[1])))

# Ex2 : Compter les mots d'une phrase
def ex2():
    phrase = "le chat le chien le chat"
    d = {}
    for mot in phrase.split():
        d[mot] = d.get(mot, 0) + 1
    print(d)

# Ex3 : Fusionner des dictionnaires
def ex3():
    d1 = {'a': 1, 'b': 2}
    d2 = {'b': 3, 'c': 4}
    d3 = {'d': 5}
    result = {}
    for d in (d1, d2, d3):
        result.update(d)
    print(result)

# Ex4 : Supprimer les dictionnaires en doublon d'une liste
def ex4():
    liste = [{'a': 1}, {'b': 2}, {'a': 1}, {'c': 3}]
    result = []
    for d in liste:
        if d not in result:
            result.append(d)
    print(result)

# Ex5 : Comparer deux dictionnaires
def ex5():
    d1 = {'a': 1, 'b': 2}
    d2 = {'b': 2, 'a': 1}
    print(d1 == d2)

# Ex6 : Fréquence des caractères
def ex6():
    chaine = "hello world"
    d = {}
    for c in chaine:
        d[c] = d.get(c, 0) + 1
    print(d)

# Ex7 : Intersection de deux dictionnaires
def ex7():
    d1 = {'a': 1, 'b': 2, 'c': 3}
    d2 = {'b': 2, 'c': 4, 'd': 5}
    print({k: v for k, v in d1.items() if k in d2 and d2[k] == v})


if __name__ == "__main__":
    for i in range(1, 8):
        print(f"\n--- Ex{i} ---")
        globals()[f"ex{i}"]()
