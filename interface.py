import tkinter as tk 
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import win32print
import time
import ttkbootstrap as ttk
import pandas as pd


# Separador visual para uso no console
linha = '__' * 15

# Classe principal da aplicação
class app():
    def __init__(self, root):
    
        self.root = root


        # lista_impressoras = self.impressoras_no_sistema()
        lista_impressoras = ...

        # Variável para armazenar a impressora selecionada
        self.impressora_escolhida = tk.StringVar()

        # Configuração de rótulos (labels) da interface
        self.label = tk.Label(root, text="Selecionar impressora:").place(x=63, y=2)
        self.label = tk.Label(root, text="Código para imprimir").place(x=22, y=65)
        self.label = tk.Label(root, text="Quantidade").place(x=168, y=65)

        # Menu de opções para selecionar impressoras disponíveis
        # option_menu = ttk.OptionMenu(root, self.impressora_escolhida, lista_impressoras[0], *lista_impressoras, bootstyle="outline")

        # option_menu.place(x=50, y=25)

        # Botão que aciona a função de impressão com a quantidade de páginas informada
        # self.botao = ttk.Button(root, text="Imprimir", width=12, command=lambda: self.loop_impressao(qtd=self.qtd_pag(), funcao_usar='imprimir_normal'), bootstyle="success")
        self.botao = ttk.Button(root, text="Imprimir", width=12, command=..., bootstyle="success")
        self.botao.place(x=22, y=151)

        # Campo de entrada para o texto a ser impresso
        self.texto_imprimir_input = ttk.Entry(root, width=20, bootstyle="success")
        self.texto_imprimir_input.place(x=22, y=85)

        # Campo de entrada para a quantidade de cópias a serem impressas
        self.quantidade_pag = ttk.Entry(root, width=8, bootstyle="info")
        self.quantidade_pag.place(x=174, y=85)

        # Checkbutton para selecionar se o código de barras será incluído na impressão
        self.check_var_codigo_barras = tk.BooleanVar()
        self.check_codigo_barras = ttk.Checkbutton(root, text="Código de barras", variable=self.check_var_codigo_barras)
        self.check_codigo_barras.place(x=10, y=130)

        # Botão que aciona a função de impressão com planilha
        # self.botao_planilha = ttk.Button(root, text="Via planilha", width=12, command=lambda: self.loop_impressao(qtd=self.qtd_pag(), funcao_usar='imprimir_via_planilha'), bootstyle="primary")
        self.botao_planilha = ttk.Button(root, text="Via planilha", width=12, command=lambda:..., bootstyle="primary")
        self.botao_planilha.place(x=130, y=151)

        self.botao_selecionar_planilha = ttk.Button(root, text="Selecionar planilha", width=12)

        # Campo de entrada para o caminho do arquivo exel
        self.caminho_exel = ttk.Entry(root, width=33, bootstyle="success")
        self.caminho_exel.place(x=22, y=210)

        self.botao_selecionar_planilha = ttk.Button(root, text="Selecionar", width=12, command=lambda: self.selecionar_exel())
        self.botao_selecionar_planilha.place(x=75, y=247)


    def selecionar_exel(self):
        caminho_planilha = filedialog.askopenfilename()
        if caminho_planilha:
            self.caminho_exel.delete(0, tk.END)
            self.caminho_exel.insert(0, caminho_planilha)


if __name__ == "__main__":
    root = tk.Tk()
    meu_app = app(root)
    root.title('PrintEasy ZPL')
    root.geometry('256x300')
    root.mainloop()
print('com')
