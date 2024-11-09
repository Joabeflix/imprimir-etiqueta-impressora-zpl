import pandas as pd

def texto_via_planilha():

    planilha = pd.read_excel('planilha.xlsx')

    texto = planilha['texto']
    quantidade = planilha['quantidade']

    lista_texto = []; lista_quantidade = []

    for tx, qt in zip(texto, quantidade):
        lista_texto.append(tx); lista_quantidade.append(qt)


    return lista_texto, lista_quantidade

listat, listq = texto_via_planilha()
print(listat)
print(listq)
