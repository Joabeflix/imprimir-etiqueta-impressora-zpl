import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32print
import time
import ttkbootstrap as ttk

# textos em variávis para eu usar como Layout para dados do console
linha = '__________________________________________________________________'

class app():
    def __init__(self, root):
        self.root = root
        #________________________________________________________________________________________________#
        # Variável que recebe a lista de impressoras da função "impressoras_no_sistema"
        lista_impressoras = self.impressoras_no_sistema()

        #________________________________________________________________________________________________#
        # Variável para manter a seleção da impressora
        impressora_escolhida = tk.StringVar()

        #________________________________________________________________________________________________#
        # Labels do programa
        self.label = tk.Label(root, text="Selecionar impressora:").place(x=63, y=2)
        self.label = tk.Label(root, text="Código para imprimir").place(x=22, y=65)
        self.label = tk.Label(root, text="Quantidade").place(x=168, y=65)

        #________________________________________________________________________________________________#
        # OptionMenu | Menu de opções que recebe a lista de impressoras
        option_menu = ttk.OptionMenu(root, impressora_escolhida, lista_impressoras[0], *lista_impressoras, bootstyle="outline")
        option_menu.place(x=50, y=25)

        #________________________________________________________________________________________________#
        # Botão executar o programa
        self.botao = ttk.Button(root, text="Imprimir", width=12, command=lambda: self.loop_impressao(self.qtd_pag()), bootstyle="success", )
        self.botao.place(x=77, y=210)

        #________________________________________________________________________________________________#
        # Entrada di texto para imprimir
        self.texto_imprimir_input = ttk.Entry(root, width=20, bootstyle="success")
        self.texto_imprimir_input.place(x=22, y=85)

        #________________________________________________________________________________________________#
        # Entrada da quantidade de páginas para imprimir
        self.quantidade_pag = ttk.Entry(root, width=8, bootstyle="info")
        self.quantidade_pag.place(x=174, y=85)

        #________________________________________________________________________________________________#
        # Criar uma BooleanVar para armazenar o estado do Checkbutton
        self.check_var_codigo_barras = tk.BooleanVar()
        # Criar o Checkbutton usando a BooleanVar
        self.check_codigo_barras = ttk.Checkbutton(root, text="Código de barras", variable=self.check_var_codigo_barras)
        self.check_codigo_barras.place(x=10, y=130)

    # Função para imprimir, ela recebe o comando (zpl)
    # e também recebe o nome da impressora escolhida

    def imprimir(self, zpl_command, printer_name):
        """ Envia comandos ZPL para a impressora """
        try:
            # Conecte-se à impressora
            conectar_impressora = win32print.OpenPrinter(printer_name)
            try:
                # Crie um objeto de trabalho
                hJob = win32print.StartDocPrinter(conectar_impressora, 1, ("ZPL Print Job", None, "RAW"))
                try:
                    # Inicie a página
                    win32print.StartPagePrinter(conectar_impressora)
                    # Envie o comando ZPL
                    win32print.WritePrinter(conectar_impressora, zpl_command.encode())
                    # Termine a página
                    win32print.EndPagePrinter(conectar_impressora)
                    # Termine o trabalho
                    win32print.EndDocPrinter(conectar_impressora)
                finally:
                    # Feche o objeto de trabalho
                    win32print.ClosePrinter(conectar_impressora)
            except Exception as e:
                print(f"Erro ao criar ou enviar o trabalho para a impressora: {e}")
        except Exception as e:
            print(f"Erro ao abrir a impressora: {e}")

    # Função de teste para usar no lugar da função principal de imprimir
    # Feita somente para testes e printar na tela p/ debugar
    def imprimir_testes(self, zpl_command, printer_name):
        print(f'')
        x = zpl_command

    # Função para retornar a lista de impressoras
    # no sistema usando o próprio win32
    def impressoras_no_sistema(self):

        impressoras_lista = []

        # Obtém todas as impressoras disponíveis
        impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)

        # Lista de impressoras para ignorar
        # para não fazer uma lista de impressora grande
        # sem impressoras que são possíveis de usar
        impressoras_ignorar = [
            'OneNote (Desktop)', 
            'Microsoft Print to PDF', 
            'Microsoft XPS Document Writer', 
            'HPBF32B6 (HP LaserJet Pro M428f-M429f)',
            'HP LaserJet Pro M428f-M429f PCL-6 (V4)',
            'Fax',
            'Fax - HP LaserJet Pro M428f-M429f',
            'Brother DCP-L2540DW series Printer (Copiar 1)',
            'Brother DCP-L2540DW series Printer',
            'HP Deskjet 2510 series',
            'OneNote for Windows 10',
            'Enviar para o OneNote 16',
            'Brother DCP-L2540DW series',
            'AnyDesk Printer'
            ]
        
        # Variável idc somente para armazenar o valor
        # e depois dar o print para mostrar no console as
        # impresssoras do sistema e fazer testes
        idc = 1

        # Fazendo a lista de impressoras e
        # Também printando no console
        print('LOGS:')
        print(linha)
        print('Lista de impressoras no sistema:'); print()

        for impressora in impressoras:
            x, y, nome, z = impressora[:4]

            if nome in impressoras_ignorar:
                pass
            else:
                impressoras_lista.append(nome)
                print(f'{idc} - {nome}')
                idc+=1

        if impressoras_lista:
            return impressoras_lista
        
        impressoras_lista = ['Você não tem nenhuma impressora compatível']
        return impressoras_lista

    # Função para retornar o nome da impressora escolhida pelo usuário
    def nome_impressora_(self):
        nome_impressora = self.impressora_escolhida.get()
        return nome_impressora

    #________________________________________________________________________________________________#
    # Função para retorna o conteúdo que vamos imprimir
    # O conteúdo já vem em código zpl
    def conteudo_imprimir(self):

        # Verificando se campo para o usuário preencher
        # "texto_imprimir" está vazio. Se estiver 
        # vai retornar None e vai parar em um verificador
        # que fica ao chamar a funcao dentro da funcao
        # "loop_impressao".
        texto_imprimir = self.texto_imprimir_input.get()
        if texto_imprimir is None:
            return None

    #________________________________________________________________________________________________#
        # Código ZPL com o código de barras
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

    #________________________________________________________________________________________________#
        # Código ZPL sem o código de barras
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

    #________________________________________________________________________________________________#
    # Verificador se vai usar imprimir o código de barras
        usar_ou_nao_codigo_barras = self.check_var_codigo_barras.get()

        if texto_imprimir:

            if usar_ou_nao_codigo_barras:
                print(linha)
                print(f'Imprimindo "{texto_imprimir}". (C/ Código de barras)')
                return zpl_code_com_barras
            else:
                print(linha)
                print(f'Imprimindo "{texto_imprimir}". (S/ Código de barras)')
                return zpl_code_sem_barras

    #________________________________________________________________________________________________#
    # Função simples para retornar a quantidade 
    # de páginas que o usuário escolheu.
    def qtd_pag(self):
        qtd = self.quantidade_pag.get()
        return qtd

    #________________________________________________________________________________________________#
    # FUNÇÃO CHAVE QUE ATIVA TODAS!!!
    #
    # Loop para imprimir a quantidade de vezes selecionada pelo usuário
    # Melhorar isso aqui kkkkkkk, foi feito na pressa total
    def loop_impressao(self, qtd):

        # Verificando o conteúdo para imprimir
        conteudo_imprimir_valor = self.conteudo_imprimir()

        # Se o campo do texto para imprimir for None, vamos mostrar um erro na tela
        if conteudo_imprimir_valor is None:
            messagebox.showerror('Erro de entrada', 'Você tem que digitar um código para imprimir!')

        else:
            
            # Tempo entre uma execução e outra
            delay_na_execucao = 3
            try:
                qtd_numero = int(qtd)
            except ValueError:
                qtd_numero = 1
            qtdd = 1
            for x in range(qtd_numero):
                print(f'Imprimindo {qtdd}/{qtd_numero}')
                qtdd+=1
                time.sleep(delay_na_execucao)
                self.imprimir(conteudo_imprimir_valor, self.nome_impressora_())

            print(linha); print('Impressão finalizada com sucesso!')


    # Código antigo, deixei aqui para testar posteriormente alguma coisa
    #________________________________________________________________________________________________#
    """"
        try:
            if qtd == '0':

                for i in range(1):
                    time.sleep(delay_na_execucao)
                    imprimir_testes(conteudo_imprimir(), nome_impressora_())

            for i in range(int(qtd)):
                time.sleep(delay_na_execucao)
                imprimir_testes(conteudo_imprimir(), nome_impressora_())

        except ValueError:

            for i in range(1):
                time.sleep(delay_na_execucao)
                print('Dentro de ValueError')
                imprimir_testes(conteudo_imprimir(), nome_impressora_())     
    """  



if __name__ == "__main__":
    root = tk.Tk()
    meu_app = app(root)
    root.title('PrintEasy ZPL')
    root.geometry('255x250')
    root.mainloop()

