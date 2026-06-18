# TP8 : Les ensembles (sets) — Solutions

# Ex1 : Vérifier si un ensemble est un sous-ensemble
def ex1():
    a = {1, 2, 3}
    b = {1, 2, 3, 4, 5}
    print(a.issubset(b))

# Ex2 : Union de tous les ensembles d'une liste
def ex2():
    lst = [{1, 2}, {2, 3}, {3, 4}]
    result = set()
    for s in lst:
        result |= s
    print(result)

# Ex3 : Intersection de tous les ensembles d'une liste
def ex3():
    lst = [{1, 2, 3}, {2, 3, 4}, {2, 3}]
    result = lst[0]
    for s in lst[1:]:
        result &= s
    print(result)

# Ex4 : Vérifier si deux ensembles sont disjoints
def ex4():
    a = {1, 2, 3}
    b = {4, 5, 6}
    print(a.isdisjoint(b))

# Ex5 : Vérifier si deux ensembles sont égaux
def ex5():
    a = {1, 2, 3}
    b = {3, 2, 1}
    print(a == b)

# Ex6 : Générer tous les sous-ensembles d'un ensemble
def ex6():
    s = {1, 2, 3}
    lst = list(s)
    result = []
    for i in range(1 << len(lst)):
        subset = {lst[j] for j in range(len(lst)) if i & (1 << j)}
        result.append(subset)
    print(result)

# Ex7 : Différence symétrique de deux ensembles
def ex7():
    a = {1, 2, 3}
    b = {3, 4, 5}
    print(a ^ b)

# Ex8 : Vérifier si un ensemble est vide
def ex8():
    print(len(set()) == 0)

# Ex9 : Trier une liste d'ensembles par leur taille
def ex9():
    lst = [{1, 2, 3}, {1}, {1, 2}, {1, 2, 3, 4}]
    print(sorted(lst, key=len))


if __name__ == "__main__":
    for i in range(1, 10):
        print(f"\n--- Ex{i} ---")
        globals()[f"ex{i}"]()
