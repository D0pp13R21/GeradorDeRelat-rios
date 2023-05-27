# Extrair Dados de Planilha - Gerador de Relatório

Este programa permite extrair dados de uma planilha no formato ODS, convertê-la para o formato XLSX e gerar um relatório com base em um período selecionado.

## Funcionalidades

- Selecionar um arquivo ODS por meio de um buscador.
- Converter o arquivo ODS para XLSX.
- Gerar um relatório com base em um período selecionado.
- Salvar o relatório em um arquivo de texto.
- Exibir uma janela de confirmação após a geração do relatório.
- Abrir automaticamente o último relatório gerado.

## Pré-requisitos

Certifique-se de ter as seguintes bibliotecas Python instaladas:

- pandas
- pyexcel
- tkinter

Você pode instalá-las usando o seguinte comando:

    pip install pandas pyexcel tkinter

## Como usar

1. Execute o programa `main.py`.
2. Clique no botão "Selecionar Arquivo .ods" e escolha o arquivo ODS que deseja extrair.
3. Selecione o período desejado na caixa de combinação.
4. Clique no botão "Gerar Relatório" para gerar o relatório.
5. Uma janela de confirmação será exibida para informar que o relatório foi gerado com sucesso.
6. Feche a janela de confirmação.
7. O arquivo de relatório será salvo no diretório do programa com o nome "relatorio_YYYYMMDDHHMMSS.txt".
8. O arquivo de relatório será aberto automaticamente.


