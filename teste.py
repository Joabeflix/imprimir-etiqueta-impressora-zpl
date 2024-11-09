v = ['teste1', 'teste2', 'teste3', 'testando']
q = ['10', '5', '3', '2']

for tx, qt in zip(v, q):
    print(tx, qt)

    qtd = 1
    for x in range(int(qt)):
        print(f'{qtd} - - - OIII')
        qtd+=1






