# TP5 : Les listes — Solutions

# Ex1 : Plus grand nombre d'une liste
def ex1():
    lst = [3, 7, 2, 9, 5]
    print(max(lst))

# Ex2 : Supprimer les doublons
def ex2():
    lst = [1, 2, 2, 3, 4, 4, 5]
    print(list(set(lst)))

# Ex3 : Membre commun entre deux listes
def ex3():
    l1 = [1, 2, 3]
    l2 = [4, 5, 3]
    print(bool(set(l1) & set(l2)))

# Ex4 : Permutations d'une liste
def ex4():
    from itertools import permutations
    lst = [1, 2, 3]
    print(list(permutations(lst)))

# Ex5 : Listes circulairement identiques
def ex5():
    l1 = [8, 8, 12, 12, 8]
    l2 = [8, 8, 8, 12, 12]
    l3 = [1, 8, 8, 12, 12]
    def circ_eq(a, b):
        if len(a) != len(b):
            return False
        return str(b) in str(a * 2)
    print("L1 == L2 :", circ_eq(l1, l2))
    print("L1 == L3 :", circ_eq(l1, l3))

# Ex6 : Deuxième plus petit nombre
def ex6():
    lst = [5, 2, 8, 1, 9, 3]
    print(sorted(set(lst))[1])

# Ex7 : Valeurs uniques
def ex7():
    lst = [1, 2, 2, 3, 4, 3, 5]
    print(list(set(lst)))

# Ex8 : Fréquence des éléments
def ex8():
    from collections import Counter
    lst = [1, 2, 2, 3, 3, 3]
    print(Counter(lst))

# Ex9 : Itérer deux listes simultanément
def ex9():
    a = [1, 2, 3]
    b = ['a', 'b', 'c']
    for x, y in zip(a, b):
        print(x, y)

# Ex10 : Supprimer toutes les occurrences d'un élément
def ex10():
    lst = [1, 2, 3, 2, 4, 2]
    val = 2
    print([x for x in lst if x != val])

# Ex11 : Concaténer deux listes par index
def ex11():
    a = ['a', 'b', 'c']
    b = [1, 2, 3]
    print([x + str(y) for x, y in zip(a, b)])

# Ex12 : Compter les valeurs uniques
def ex12():
    lst = [1, 2, 2, 3, 4, 4, 5]
    print(len(set(lst)))


if __name__ == "__main__":
    for i in range(1, 13):
        print(f"\n--- Ex{i} ---")
        globals()[f"ex{i}"]()
