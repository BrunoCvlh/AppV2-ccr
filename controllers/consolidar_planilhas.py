# consolidar_planilhas.py
import pandas as pd
import os
from pathlib import Path

def consolidar_planilhas(file_paths: list, output_file_name="planilha_consolidada.xlsx"):

    if not file_paths:
        return "ERRO: Nenhuma planilha fornecida para consolidação."

    all_data = []
    erros_leitura = []

    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path)
            all_data.append(df)
        except Exception as e:
            erros_leitura.append(f"Erro ao ler o arquivo '{file_path.split('/')[-1]}': {e}")
            continue

    if not all_data:
        return f"ERRO: Nenhuma planilha pôde ser lida. Erros: {'; '.join(erros_leitura)}"

    try:
        consolidated_df = pd.concat(all_data, ignore_index=True)

        downloads_path = str(Path.home() / "Downloads")
        output_path = os.path.join(downloads_path, output_file_name)

        consolidated_df.to_excel(output_path, index=False)
        return output_path
    except Exception as e:
        return f"ERRO: Erro ao concatenar ou salvar as planilhas: {e}"