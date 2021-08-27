
lista1 = ['a', 'b', 'c']
lista2 = ['d', 'e', 'f', 'g', 'h']

a = [(x, y) for x in lista1 for y in lista2]
slownik = dict(zip(a, range(15)))
print(slownik)