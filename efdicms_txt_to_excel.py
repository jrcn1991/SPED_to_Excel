import os
import pandas as pd
from xlsxwriter import Workbook

# Obter o diretório atual do script
script_dir = os.path.dirname(__file__)

# Função para processar o arquivo
def process_file_adjusted(file_path):
    with open(file_path, 'r', encoding='ISO-8859-1') as file:
        data = {}
        for line in file:
            # Ignorar a primeira coluna vazia
            elements = line.strip().split('|')[1:]
            reg_id = elements[0]
            if reg_id not in data:
                data[reg_id] = []
            data[reg_id].append(elements)
        return data

# Caminho para o arquivo .txt na pasta do projeto
file_name = 'SpedEFD-11035026000191-79637874-Remessa de arquivo original-jan2024.txt'
file_path = os.path.join(script_dir, 'Arquivo_Teste', file_name)

# Processar o arquivo com o ajuste
data_adjusted = process_file_adjusted(file_path)

# Criar um dicionário genérico para cabeçalhos ajustado
generic_headers_adjusted = {key: [f'cabeçalho{i+1}' for i in range(len(data_adjusted[key][0]))] for key in data_adjusted.keys()}

# Caminho para o arquivo Excel na pasta de saída
output_dir = os.path.join(script_dir, 'Output')
output_file_path = os.path.join(output_dir, 'Processed_Data_Adjusted.xlsx')

# Criar um DataFrame para cada conjunto de dados ajustado e salvar em um arquivo Excel
excel_writer_adjusted = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
for reg_id, values in data_adjusted.items():
    df_adjusted = pd.DataFrame(values, columns=generic_headers_adjusted[reg_id])
    df_adjusted.to_excel(excel_writer_adjusted, sheet_name=reg_id, index=False)

# Salvar o arquivo Excel ajustado
excel_writer_adjusted.close()

