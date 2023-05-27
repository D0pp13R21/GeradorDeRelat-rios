import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyexcel as pe
import os
from datetime import datetime
from tkinter import ttk

def converter_ods_para_xlsx(caminho_ods):
    # Obter o diretório do arquivo do programa
    diretorio_programa = os.path.dirname(os.path.abspath(__file__))
    nome_arquivo_xlsx = os.path.join(diretorio_programa, 'planilha.xlsx')

    # Converter arquivo .ods para .xlsx
    pe.save_book_as(file_name=caminho_ods, dest_file_name=nome_arquivo_xlsx)

    return nome_arquivo_xlsx

def abrir_arquivo(caminho_arquivo):
    try:
        if os.name == 'nt':
            os.startfile(caminho_arquivo)
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, caminho_arquivo])
    except OSError:
        pass

def selecionar_arquivo_ods():
    # Abrir janela de seleção de arquivo .ods
    janela_arquivo = tk.Tk()
    janela_arquivo.withdraw()  # Ocultar a janela principal

    caminho_ods = filedialog.askopenfilename(title="Selecionar arquivo .ods")

    if caminho_ods:
        # Converter .ods para .xlsx
        caminho_xlsx = converter_ods_para_xlsx(caminho_ods)
        label_arquivo_selecionado.config(text=f"Arquivo selecionado: {caminho_xlsx}")
    else:
        label_arquivo_selecionado.config(text="Nenhum arquivo selecionado")

def gerar_relatorio():
    # Verificar se o arquivo .ods foi selecionado
    caminho_xlsx = label_arquivo_selecionado.cget("text")
    if caminho_xlsx.startswith("Arquivo selecionado: "):
        caminho_xlsx = caminho_xlsx.split(": ")[1]

        # Lendo a planilha
        planilha = pd.read_excel(caminho_xlsx)

        # Convertendo a coluna 'Data da Ocorrência' para o formato de data
        planilha['Data da Ocorrência'] = pd.to_datetime(planilha['Data da Ocorrência'], format='%d/%m/%Y')

        # Obtendo a opção selecionada pelo usuário
        periodo_selecionado = combo_periodo.get()

        # Definindo o período com base na opção selecionada
        if periodo_selecionado == '24 horas':
            offset = pd.DateOffset(days=1)
        elif periodo_selecionado == '48 horas':
            offset = pd.DateOffset(days=2)
        elif periodo_selecionado == '72 horas':
            offset = pd.DateOffset(days=3)
        elif periodo_selecionado == 'Última semana':
            offset = pd.DateOffset(weeks=1)
        elif periodo_selecionado == 'Último mês':
            offset = pd.DateOffset(months=1)
        elif periodo_selecionado == 'Último ano':
            offset = pd.DateOffset(years=1)
        elif periodo_selecionado == 'Personalizado':
            mes_selecionado = combo_mes.get()
            ano_selecionado = combo_ano.get()
            if not mes_selecionado or not ano_selecionado:
                messagebox.showinfo("Informação", "Selecione um mês e um ano para o período personalizado.")
                return
            data_inicio = pd.Timestamp(year=int(ano_selecionado), month=int(mes_selecionado), day=1)
            data_fim = data_inicio + pd.DateOffset(months=1) - pd.DateOffset(days=1)
            offset = None

        # Gerando o relatório para o período selecionado
        if offset:
            data_inicio = pd.Timestamp.now() - offset
            data_fim = pd.Timestamp.now()
        ocorrencias_periodo = planilha[
            (planilha['Data da Ocorrência'] >= data_inicio) &
            (planilha['Data da Ocorrência'] <= data_fim)
        ]

        # Obtendo o número total de ocorrências no período
        total_ocorrencias = len(ocorrencias_periodo)

        # Gerando o relatório
        relatorio = f"Relatório de Ocorrências\n\n"
        relatorio += f"Período: {data_inicio.strftime('%d/%m/%Y %H:%M:%S')} - {data_fim.strftime('%d/%m/%Y %H:%M:%S')}\n"
        relatorio += f"Total de ocorrências: {total_ocorrencias}\n\n"

        if total_ocorrencias > 0:
            for tipo_ocorrencia, grupo in ocorrencias_periodo.groupby('Tipo de Ocorrência'):
                total_tipo_ocorrencia = len(grupo)
                bairros = ', '.join(grupo['Bairro'])
                
                relatorio += f"Tipo de Ocorrência: {tipo_ocorrencia}\n"
                relatorio += f"Total de ocorrências desse tipo: {total_tipo_ocorrencia}\n"
                relatorio += f"Bairros: {bairros}\n\n"
            
            # Salvando o relatório em um arquivo de texto
            nome_arquivo_relatorio = f"relatorio_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
            with open(nome_arquivo_relatorio, 'w') as arquivo_saida:
                arquivo_saida.write(relatorio)
            
            # Exibindo a mensagem de sucesso e abrindo o arquivo
            messagebox.showinfo("Informação", "Relatório gerado com sucesso.")
            abrir_arquivo(nome_arquivo_relatorio)

        else:
            messagebox.showinfo("Informação", "Nenhum arquivo selecionado.")

# Criando a janela principal
janela = tk.Tk()
janela.title("Gerador de Relatório")
janela.geometry("400x250")

# Criando um botão para selecionar o arquivo .ods
botao_selecionar_arquivo = tk.Button(janela, text="Selecionar Arquivo .ods", command=selecionar_arquivo_ods)
botao_selecionar_arquivo.pack()

# Rótulo para exibir o arquivo selecionado
label_arquivo_selecionado = tk.Label(janela, text="Nenhum arquivo selecionado")
label_arquivo_selecionado.pack()

# Criando um rótulo
label_periodo = tk.Label(janela, text="Selecione o período:")
label_periodo.pack()

# Criando uma caixa de combinação (combobox)
opcoes_periodo = ['24 horas', '48 horas', '72 horas', 'Última semana', 'Último mês', 'Último ano', 'Personalizado']
combo_periodo = ttk.Combobox(janela, values=opcoes_periodo)
combo_periodo.pack()

# Criando caixas de combinação para mês e ano
frame_personalizado = tk.Frame(janela)

label_mes = tk.Label(frame_personalizado, text="Mês:")
label_mes.pack(side=tk.LEFT)
combo_mes = ttk.Combobox(frame_personalizado, values=[str(i) for i in range(1, 13)])
combo_mes.pack(side=tk.LEFT)

label_ano = tk.Label(frame_personalizado, text="Ano:")
label_ano.pack(side=tk.LEFT)
combo_ano = ttk.Combobox(frame_personalizado, values=[str(i) for i in range(2000, 2031)])
combo_ano.pack(side=tk.LEFT)

frame_personalizado.pack()

# Criando um botão para gerar o relatório
botao_gerar_relatorio = tk.Button(janela, text="Gerar Relatório", command=gerar_relatorio)
botao_gerar_relatorio.pack()

# Iniciando o loop principal da janela
janela.mainloop()
