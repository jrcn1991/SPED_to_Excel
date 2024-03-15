import pandas as pd

def excel_to_txt_corrected_final(excel_path, txt_path):
    # Ler o arquivo Excel
    xls = pd.ExcelFile(excel_path)

    # Criar um arquivo .txt para gravar os dados
    with open(txt_path, 'w', encoding='ISO-8859-1') as txt_file:
        # Processar cada aba do arquivo Excel
        for sheet_name in xls.sheet_names:
            # Ler a aba atual
            df = pd.read_excel(xls, sheet_name, dtype=str)

            # Converter cada linha da aba para o formato do arquivo .txt original
            for index, row in df.iterrows():
                # Tratar valores nulos (NaN) como strings vazias
                row = row.fillna('')

                # Recriar a linha com os separadores "|"
                line = '|' + '|'.join(row) + '|\n'
                line = line[:-2] + '\n'  # Remove o último '|' e mantém a quebra de linha
                txt_file.write(line)

# Caminho para o arquivo Excel processado anteriormente
excel_path = r'C:\Users\jrcn0\OneDrive2\OneDrive\Documentos\GitHub\SPED_to_Excel\Output\Processed_Data_Adjusted.xlsx'

# Caminho para o novo arquivo .txt que será criado
txt_path = r'C:\Users\jrcn0\OneDrive2\OneDrive\Documentos\GitHub\SPED_to_Excel\Output\Converted_Back_To_Txt.txt'

# Chamar a função para converter o Excel de volta para .txt com as correções
excel_to_txt_corrected_final(excel_path, txt_path)
