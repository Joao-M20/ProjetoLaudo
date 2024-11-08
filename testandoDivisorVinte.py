
#ORIGINALLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLL
#ORIGINALLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLL

from docx import Document
import pandas as pd
import math

# Caminho do arquivo Excel
excel_path = 'C:\\Users\\BurgerKing\\Downloads\\LAUDO - SUCATA BK.xlsx'

# Ler a planilha do Excel
df = pd.read_excel(excel_path, sheet_name='LAUDO 113')

# Número de linhas por documento
linhas_por_documento = 20

# Calcular quantos documentos são necessários (divisão arredondada para cima)
num_documentos = math.ceil(df.shape[0] / linhas_por_documento)

# Criar documentos dinamicamente com base no número de documentos necessários
for i in range(num_documentos):
    # Obter as linhas do DataFrame para o documento atual
    inicio = i * linhas_por_documento
    fim = inicio + linhas_por_documento
    dados_documento = df.iloc[inicio:fim]

    # Criar um novo documento Word
    doc = Document()

    
    # Adicionar a tabela com (número de linhas + 2 colunas adicionais)
    table = doc.add_table(rows=len(dados_documento) + 1, cols=len(dados_documento.columns) + 2)

    # Preencher a primeira coluna com os números de "Item" (1 a 20)
    for j, row in enumerate(dados_documento.values):
        # Coluna "Item"
        table.cell(j, 0).text = str(j + 1)
        
        # Coluna vazia
        table.cell(j, 1).text = ""
        
        # Preencher o restante das colunas com os dados do Excel
        for k, cell in enumerate(row):
            table.cell(j, k + 2).text = str(cell)  # k+2 para ajustar à nova estrutura

    # Salvar o documento com um nome único
    doc_path = f'C:\\Users\\BurgerKing\\Documents\\Equipamentos_Reparados_Parte{i+1}.docx'
    doc.save(doc_path)
    
    print(f"Documento {i+1} criado com as linhas de {inicio+1} a {fim}.")
