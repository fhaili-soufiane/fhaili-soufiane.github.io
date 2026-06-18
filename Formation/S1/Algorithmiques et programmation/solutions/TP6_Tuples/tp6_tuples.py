# TP6 : Les tuples — Solutions

# Ex1 : Inverser un tuple
def ex1():
    t = (1, 2, 3, 4)
    print(t[::-1])

# Ex2 : Tuple imbriqué — imprimer des valeurs
def ex2():
    t1 = ([1, 2, 3], (5, 10, 15))
    print("Liste:", t1[0])
    print("Tuple:", t1[1])
    print("Éléments liste:", *t1[0])
    print("Éléments tuple:", *t1[1])

# Ex3 : Échanger deux tuples
def ex3():
    a = (1, 2)
    b = (3, 4)
    a, b = b, a
    print("a =", a, "b =", b)

# Ex4 : Trier un tuple de tuples par 2ème élément
def ex4():
    t1 = (('Marrakech', 1.3), ('Casa', 3.7), ('Rabat', 1.1), ('Agadir', 0.9))
    print(tuple(sorted(t1, key=lambda x: x[1])))

# Ex5 : Vérifier si tous les éléments d'un tuple sont égaux
def ex5():
    t = (5, 5, 5, 5)
    print(len(set(t)) == 1)

# Ex6 : Supprimer les tuples imbriqués d'un tuple
def ex6():
    t = (1, (2, 3), 4, (5, 6), 7)
    print(tuple(x for x in t if not isinstance(x, tuple)))

# Ex7 : Trier les tuples par leur élément max
def ex7():
    lst = [(1, 5), (3, 2), (7, 8), (2, 4)]
    print(sorted(lst, key=lambda t: max(t)))

# Ex8 : Générer toutes les combinaisons de paires de 2 tuples
def ex8():
    from itertools import product
    a = (1, 2)
    b = ('x', 'y')
    print(list(product(a, b)))


if __name__ == "__main__":
    for i in range(1, 9):
        print(f"\n--- Ex{i} ---")
        globals()[f"ex{i}"]()
