import pandas as pd
import os
from datetime import datetime
import re

def tratar_planilha(file_path, competencia):
    try:
        df = pd.read_excel(file_path, header=None) 
        
        balancete_info = str(df.iloc[2, 0]) 
        name_mapping = {
            "PLANO_DE_GESTÃO": "PGA",
            "BENEFICIO_DEFINIDO": "BD",
            "POSTALPREV": "PP"
        }

        # Verifica se a string tem pelo menos 86 caracteres (66 + 20) antes de tentar cortar
        if len(balancete_info) >= 86:
            # Corta 20 caracteres a partir do índice 66 (que é o 67º caractere)
            raw_extracted_name = balancete_info[66:86].strip()
            
            # Tenta encontrar uma correspondência no dicionário de mapeamento
            # A chave do dicionário deve ser exatamente como o nome é extraído *neste ponto*
            if raw_extracted_name in name_mapping:
                balancete_name = name_mapping[raw_extracted_name]
            else:
                # Se não houver mapeamento direto, aplica a lógica de limpeza genérica
                # 1. Remover o prefixo numérico e o hífen (ex: "96-")
                cleaned_balancete_name = re.sub(r'^\d+-', '', raw_extracted_name)
                
                # 2. Manter apenas caracteres em maiúsculas e underscores.
                final_balancete_name_parts = []
                for char in cleaned_balancete_name:
                    if char.isupper() or char == '_':
                        final_balancete_name_parts.append(char)
                balancete_name = "".join(final_balancete_name_parts).strip('_')
        else:
            balancete_name = ""

        df = df.iloc[4:].reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        df = df.iloc[:, [0, 3, 4]]
        df.columns = [df.columns[0], "Orçado", "Realizado"]
        
        try:
            competencia_date = datetime.strptime(competencia, "%m/%Y")
            primeiro_dia = competencia_date.strftime("%d/%m/%Y")
        except Exception:
            primeiro_dia = ""
        df["Data Competência"] = primeiro_dia
        
        if len(df) > 2:
            df = df.iloc[:-2].reset_index(drop=True)
        
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        
        if balancete_name:
            # Usa o nome mapeado/limpo como o nome do arquivo final
            tratado_file_name = f"{balancete_name}.xlsx"
        else:
            # Fallback se o balancete_name não for extraído ou mapeado/limpo
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            tratado_file_name = f"{name_without_ext}_tratada_{timestamp}.xlsx"
        
        tratado_path = os.path.join(downloads_folder, tratado_file_name)
        df.to_excel(tratado_path, index=False)
        return tratado_path
    except Exception as e:
        return f"ERRO: {e}"