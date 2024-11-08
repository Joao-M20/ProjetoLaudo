from docx import Document
import pandas as pd
import math

# Caminho do arquivo Excel
excel_path = 'C:\\Users\\BurgerKing\\Downloads\\LAUDO - SUCATA BK.xlsx'

# Ler a planilha do Excel
df = pd.read_excel(excel_path, sheet_name='LAUDO 113', skiprows=1)

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
    
    # Criar uma tabela com exatamente 20 linhas e colunas + 2 adicionais
    table = doc.add_table(rows=linhas_por_documento, cols=len(dados_documento.columns) + 2)

    # Preencher a tabela com os dados
    for j in range(linhas_por_documento):
        # Preencher a coluna "Item" com números de 1 a 20
        table.cell(j, 0).text = str(j + 1)
        
        # Coluna vazia
        table.cell(j, 1).text = ""
        
        # Verificar se ainda há dados para preencher
        if j < len(dados_documento):
            # Preencher as colunas com dados reais
            for k, cell in enumerate(dados_documento.iloc[j]):
                table.cell(j, k + 2).text = str(cell)
        else:
            # Caso não haja mais dados, deixar as células vazias
            for k in range(len(dados_documento.columns)):
                table.cell(j, k + 2).text = ""

    # Salvar o documento com um nome único
    doc_path = f'C:\\Users\\BurgerKing\\Documents\\Equipamentos_Reparados_Parte{i+1}.docx'
    doc.save(doc_path)
    
    print(f"Documento {i+1} criado com as linhas de {inicio+1} a {fim}.")
