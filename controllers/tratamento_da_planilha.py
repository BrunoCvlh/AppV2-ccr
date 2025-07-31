import pandas as pd
import os
from datetime import datetime

def tratar_planilha(file_path: str, competencia: str, nome_balancete: str):
    try:
        df = pd.read_excel(file_path, skiprows=6, header=None)

        df_tratado = df.iloc[:, [0, 3, 4]].copy()

        df_tratado.columns = ['Conta', 'Orçado', 'Realizado']

        if len(df_tratado) >= 2:
            df_tratado = df_tratado.iloc[:-2]
        elif len(df_tratado) == 1:
            df_tratado = df_tratado.iloc[:-1]
        else:
            pass

        # Extrai mês e ano da string de competência
        mes, ano = map(int, competencia.split('/'))
        # Cria um objeto datetime para o primeiro dia do mês e ano
        primeiro_dia_competencia = datetime(ano, mes, 1).strftime('%d/%m/%Y')
        
        df_tratado['Competência'] = primeiro_dia_competencia

        df_tratado['Balancete'] = nome_balancete

        directory, filename = os.path.split(file_path)
        name, ext = os.path.splitext(filename)
        output_filename = f"{name}_tratado{ext}"
        output_path = os.path.join(directory, output_filename)

        df_tratado.to_excel(output_path, index=False)

        return output_path
    except FileNotFoundError:
        return f"ERRO: Arquivo não encontrado em '{file_path}'"
    except ValueError:
        return f"ERRO: Formato de 'competencia' inválido. Use 'MM/AAAA' (ex: '05/2025')."
    except Exception as e:
        return f"ERRO: Ocorreu um erro ao tratar o arquivo '{file_path}': {e}"