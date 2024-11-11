import tkinter as tk 
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import win32print
import time
import ttkbootstrap as ttk
import pandas as pd


# Separador visual para uso no console
linha = '__________________________________________________________________'

# Classe principal da aplicação
class app():
    def __init__(self, root):
        
        # Inicializa a interface gráfica
        self.root = root

        # Lista de impressoras disponíveis, obtida pela função "impressoras_no_sistema"
        lista_impressoras = self.impressoras_no_sistema()

        # Variável para armazenar a impressora selecionada
        self.impressora_escolhida = tk.StringVar()

        # Configuração de rótulos (labels) da interface
        self.label = tk.Label(root, text="Selecionar impressora:").place(x=63, y=2)
        self.label = tk.Label(root, text="Código para imprimir").place(x=22, y=65)
        self.label = tk.Label(root, text="Quantidade").place(x=168, y=65)

        # Menu de opções para selecionar impressoras disponíveis
        option_menu = ttk.OptionMenu(root, self.impressora_escolhida, lista_impressoras[0], *lista_impressoras, bootstyle="outline")
        option_menu.place(x=50, y=25)

        # Botão que aciona a função de impressão com a quantidade de páginas informada
        self.botao = ttk.Button(root, text="Imprimir", width=12, command=lambda: self.loop_impressao(qtd=self.qtd_pag(), funcao_usar='imprimir_normal'), bootstyle="success")
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
        self.botao_planilha = ttk.Button(root, text="Via planilha", width=12, command=lambda: self.loop_impressao(qtd=self.qtd_pag(), funcao_usar='imprimir_via_planilha'), bootstyle="primary")
        self.botao_planilha.place(x=130, y=151)

        self.botao_selecionar_planilha = ttk.Button(root, text="Selecionar planilha", width=12)

        # Campo de entrada para o caminho do arquivo exel
        self.caminho_exel = ttk.Entry(root, width=33, bootstyle="success")
        self.caminho_exel.place(x=22, y=210)

        self.botao_selecionar_planilha = ttk.Button(root, text="Selecionar", width=12, command=self.selecionar_arquivo)
        self.botao_selecionar_planilha.place(x=75, y=247)


    def selecionar_arquivo(self):
        caminho_planilha = filedialog.askopenfilename()
        if caminho_planilha:
            self.caminho_exel.delete(0, tk.END)  # Limpa a entrada antes de inserir o novo caminho
            self.caminho_exel.insert(0, caminho_planilha)


    # Função para enviar o comando ZPL para a impressora
    def imprimir(self, comando_zpl, nome_impressora):
        """ Envia comandos ZPL para a impressora """
        try:
            # Conectar à impressora selecionada
            conectar_impressora = win32print.OpenPrinter(nome_impressora)
            try:
                # Cria um objeto de trabalho de impressão
                hJob = win32print.StartDocPrinter(conectar_impressora, 1, ("ZPL Print Job", None, "RAW"))
                try:
                    # Inicia uma nova página de impressão
                    win32print.StartPagePrinter(conectar_impressora)
                    # Envia o comando ZPL para a impressora
                    win32print.WritePrinter(conectar_impressora, comando_zpl.encode())
                    # Finaliza a página de impressão
                    win32print.EndPagePrinter(conectar_impressora)
                    # Finaliza o trabalho de impressão
                    win32print.EndDocPrinter(conectar_impressora)
                finally:
                    # Fecha a conexão com a impressora
                    win32print.ClosePrinter(conectar_impressora)
            except Exception as e:
                print(f"Erro ao criar ou enviar o trabalho para a impressora: {e}")
        except Exception as e:
            print(f"Erro ao abrir a impressora: {e}")

    # Função de teste para debug, apenas imprime no console
    def imprimir_testes(self, zpl_command, printer_name):
        print(f'')
        x = zpl_command

    # Função para retornar a lista de impressoras disponíveis no sistema
    def impressoras_no_sistema(self):
        impressoras_lista = []

        # Obtém todas as impressoras locais e de rede conectadas
        impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)

        # Lista de impressoras a serem ignoradas
        impressoras_ignorar = [
            'OneNote (Desktop)', 
            'Microsoft Print to PDF', 
            'Microsoft XPS Document Writer', 
            'HPBF32B6 (HP LaserJet Pro M428f-M429f)',
            'HP LaserJet Pro M428f-M429f PCL-6 (V4)',
            'Fax - HP LaserJet Pro M428f-M429f',
            'Brother DCP-L2540DW series Printer (Copiar 1)',
            'Brother DCP-L2540DW series Printer',
            'HP Deskjet 2510 series',
            'OneNote for Windows 10',
            'Enviar para o OneNote 16',
            'Brother DCP-L2540DW series',
            'AnyDesk Printer'
            ]
        
        # Print de diagnóstico no console para listagem das impressoras
        idc = 1
        print('LOGS:')
        print(linha)
        print('Lista de impressoras no sistema:'); print()

        # Filtrando as impressoras disponíveis, excluindo as da lista de ignorados
        for impressora in impressoras:
            x, y, nome, z = impressora[:4]
            if nome not in impressoras_ignorar:
                impressoras_lista.append(nome)
                print(f'{idc} - {nome}')
                idc += 1

        if impressoras_lista:
            return impressoras_lista
        
        # Caso não haja impressoras compatíveis, exibe uma mensagem
        impressoras_lista = ['Você não tem nenhuma impressora compatível']
        return impressoras_lista

    # Função para obter o nome da impressora escolhida
    def nome_impressora_(self):
        nome_impressora = self.impressora_escolhida.get()
        return nome_impressora

    # Função que retorna o conteúdo a ser impresso em formato ZPL
    def conteudo_imprimir_em_zpl(self, conteudo_imprimir):
        # Obtém o texto digitado para impressão
        texto_imprimir = conteudo_imprimir
        if texto_imprimir is None:
            return None

        # Define o código ZPL com o código de barras
        zpl_code_com_barras = f"""
        ^XA
        ^PW1024          ; Largura total da etiqueta (50mm + 50mm de cada coluna + 12 pontos de gap)
        ^LH0,0           ; Define a origem do layout no canto superior esquerdo
        ^CF0,45          ; Define a fonte e o tamanho do texto (ajuste conforme necessário)

        ; Coluna 1 - Centralizada
        ^FO-65,00        ; Posição inicial para a primeira coluna (deslocada 60 pontos para a esquerda, 10 pontos da parte superior)
        ^FB500,3,0,C,0   ; Define a caixa de texto de 500 pontos de largura (50 mm), até 3 linhas, centralizando o texto
        ^FD{texto_imprimir}^FS

        ; Código de Barras na Coluna 1
        ^FO60,50        ; Posição inicial para o código de barras da primeira coluna
        ^BY1,3,35        ; Largura do código de barras (1), altura (2), densidade (30)
        ^B3N,N,30,N,N    ; Código de barras tipo 39, sem texto abaixo, altura 30, com verificação de checksum
        ^FD>{texto_imprimir}^FS  ; Código de barras para a primeira coluna

        ; Coluna 2 - Centralizada (com separação de 3 mm, aproximadamente 12 pontos)
        ^FO355,00        ; Posição inicial para a segunda coluna (500 pontos da primeira coluna + 12 pontos de gap, 10 pontos da parte superior)
        ^FB500,3,0,C,0   ; Define a caixa de texto de 500 pontos de largura (50 mm), até 3 linhas, centralizando o texto
        ^FD{texto_imprimir}^FS

        ; Código de Barras na Coluna 2
        ^FO480,50        ; Posição inicial para o código de barras da segunda coluna
        ^BY1,3,35        ; Largura do código de barras (1), altura (2), densidade (30)
        ^B3N,N,30,N,N    ; Código de barras tipo 39, sem texto abaixo, altura 30, com verificação de checksum
        ^FD>{texto_imprimir}^FS  ; Código de barras para a segunda coluna

        ^XZ
        """

        # Define o código ZPL sem o código de barras
        zpl_code_sem_barras = f"""
        ^XA
        ^PW1024          ; Largura total da etiqueta (50mm + 50mm de cada coluna + 12 pontos de gap)
        ^LH0,0           ; Define a origem do layout no canto superior esquerdo
        ^CF0,45          ; Define a fonte e o tamanho do texto (ajuste conforme necessário)

        ; Coluna 1 - Centralizada
        ^FO-65,00        ; Posição inicial para a primeira coluna (deslocada 60 pontos para a esquerda, 10 pontos da parte superior)
        ^FB500,3,0,C,0   ; Define a caixa de texto de 500 pontos de largura (50 mm), até 3 linhas, centralizando o texto
        ^FD{texto_imprimir}^FS

        ; Coluna 2 - Centralizada (com separação de 3 mm, aproximadamente 12 pontos)
        ^FO355,00        ; Posição inicial para a segunda coluna (500 pontos da primeira coluna + 12 pontos de gap, 10 pontos da parte superior)
        ^FB500,3,0,C,0   ; Define a caixa de texto de 500 pontos de largura (50 mm), até 3 linhas, centralizando o texto
        ^FD{texto_imprimir}^FS
        ^XZ
        """

        # Verifica se o código de barras será usado
        usar_ou_nao_codigo_barras = self.check_var_codigo_barras.get()

        if texto_imprimir:
            if usar_ou_nao_codigo_barras:
                print(f'Imprimindo "{texto_imprimir}". (C/ Código de barras)')
                return zpl_code_com_barras
            else:
                print(f'Imprimindo "{texto_imprimir}". (S/ Código de barras)')
                return zpl_code_sem_barras

    # Função para retornar a quantidade de páginas escolhida para impressão
    def qtd_pag(self):
        qtd = self.quantidade_pag.get()
        return qtd

    # Função principal para iniciar a impressão de acordo com a quantidade de páginas
    def loop_impressao(self, qtd, funcao_usar):

        def imprimir_normal():
            texto_imprimir = self.texto_imprimir_input.get()
            # Obtém o conteúdo a ser impresso
            conteudo_imprimir_valor = self.conteudo_imprimir_em_zpl(conteudo_imprimir=texto_imprimir)

            # Verifica se o campo de texto de impressão está vazio
            if conteudo_imprimir_valor is None:
                messagebox.showerror('Erro de entrada', 'Você tem que digitar um código para imprimir!')

            else:
                # Define o intervalo de espera entre impressões
                delay_na_execucao = 3
                try:
                    qtd_numero = int(qtd)
                except ValueError:
                    qtd_numero = 1
                qtdd = 1
                for x in range(qtd_numero):
                    print(f'Imprimindo {qtdd}/{qtd_numero}')
                    qtdd += 1
                    time.sleep(delay_na_execucao)
                    self.imprimir(conteudo_imprimir_valor, self.nome_impressora_())

                print(linha); print('Impressão finalizada com sucesso!')
        def imprimir_via_planilha():

            lista_texto, lista_de_qtd = self.retorno_lista_planilha()

            delay_na_execucao = 3
            print(linha)
            print('!!!(Imprimindo via planilha)!!!')
            print(linha)
            for texto_imprimir, quantidade in zip(lista_texto, lista_de_qtd):
                qtdd = 1
                for i in range(int(quantidade)):
                    print(f'Imprimindo {qtdd}/{quantidade}')
                    qtdd += 1
                    time.sleep(delay_na_execucao)
                    self.imprimir(self.conteudo_imprimir_em_zpl(conteudo_imprimir=texto_imprimir), self.nome_impressora_())
                
            print(linha); print('Impressão finalizada com sucesso!')
        if funcao_usar == 'imprimir_normal':
            imprimir_normal()
        else:
            imprimir_via_planilha()
        

    def retorno_lista_planilha(self):
        caminho_planilha = self.caminho_exel.get()
        planilha = pd.read_excel(caminho_planilha)

        texto = planilha['texto']
        quantidade = planilha['quantidade']

        lista_texto = []; lista_quantidade = []

        for tx, qt in zip(texto, quantidade):
            lista_texto.append(tx); lista_quantidade.append(qt)

        return lista_texto, lista_quantidade

# Inicializa a interface gráfica e executa o loop principal do Tkinter
if __name__ == "__main__":
    root = tk.Tk()
    meu_app = app(root)
    root.title('PrintEasy ZPL')
    root.geometry('256x300')
    root.mainloop()
print('com')
