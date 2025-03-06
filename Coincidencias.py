import pandas as pd

def comparar_planilhas_excel(caminho_arquivo_excel, nome_planilha_municipio, nome_planilha_h12,
                             nome_nova_planilha_coincidencias='Coincidencias', nome_nova_planilha_resumo='Nat. Rubr.',
                             nome_nova_planilha_incidencia='Incidencia', nome_nova_planilha_sem_equivalencia='Sem Equivalencia',
                             nome_nova_planilha_rubricas='Rubricas', nome_nova_planilha_quantitativo='Quantitativo',
                             nome_nova_planilha_coincidencias_unicas='Coincidencias Unicas'): # Novo nome da planilha
    """
    Compara planilhas Excel e gera 'Coincidencias', 'Resumo', 'Incidencia', 'Sem Equivalencia', 'Rubricas', 'Quantitativo' e 'Coincidencias Unicas'.

    Args:
        caminho_arquivo_excel (str): Caminho do arquivo Excel.
        nome_planilha_municipio (str): Nome da planilha 'MUNICIPIO'.
        nome_planilha_h12 (str): Nome da planilha 'H12'.
        nome_nova_planilha_coincidencias (str, opcional): Nome da planilha 'Coincidencias'.
        nome_nova_planilha_resumo (str, opcional): Nome da planilha 'Resumo'.
        nome_nova_planilha_incidencia (str, opcional): Nome da planilha 'Incidencia'.
        nome_nova_planilha_sem_equivalencia (str, opcional): Nome da planilha 'Sem Equivalencia'.
        nome_nova_planilha_rubricas (str, opcional): Nome da planilha 'Rubricas'.
        nome_nova_planilha_quantitativo (str, opcional): Nome da planilha 'Quantitativo'.
        nome_nova_planilha_coincidencias_unicas (str, opcional): Nome da planilha 'Coincidencias Unicas'. # Novo parâmetro
    """

    # 1. Leitura das planilhas
    try:
        df_municipio = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_planilha_municipio, header=1)
        df_h12 = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_planilha_h12)
    except Exception as e:
        print(f"Erro ao ler as planilhas: {e}")
        return

    # Verificação dos nomes das colunas da planilha MUNICIPIO
    print("Nomes das colunas da planilha MUNICIPIO:")
    print(df_municipio.columns)

    # Verificação dos nomes das colunas da planilha H12
    print("\nNomes das colunas da planilha H12:")
    print(df_h12.columns)

    # Listas para armazenar as linhas das novas planilhas
    linhas_nova_planilha_coincidencias = []
    linhas_nova_planilha_resumo = []
    linhas_nova_planilha_incidencia = []
    linhas_nova_planilha_sem_equivalencia = []
    linhas_nova_planilha_rubricas = []
    linhas_nova_planilha_quantitativo = []
    linhas_nova_planilha_coincidencias_unicas = [] # Nova lista para "Coincidencias Unicas"

    # 2. Iteração e Comparação
    for index_municipio, linha_municipio in df_municipio.iterrows():
        codigo_esocial_municipio_coluna_correta = 'CÓDIGO ESOCIAL'
        codigo_esocial_municipio = linha_municipio[codigo_esocial_municipio_coluna_correta]
        nat_rubr_coluna_correta_h12 = ' natRubr'
        coincidencias_h12 = df_h12[df_h12[nat_rubr_coluna_correta_h12] == codigo_esocial_municipio]

        # Dados fixos da planilha MUNICIPIO para a nova linha (para todas as planilhas)
        dados_municipio_fixos = {
            'COD - MUNICIPIO': linha_municipio['COD.'],
            'RUBRICA - MUNICIPIO': linha_municipio['RUBRICAS'],
            'NAT. RUBR. - MUNICIPIO': codigo_esocial_municipio,
            'INC. CP - MUNICIPIO': linha_municipio['INSS'],
            'INC IRRF - MUNICIPIO': linha_municipio['IR']
        }

        nova_linha_rubricas = { # Inicializa nova_linha_rubricas aqui
            'COD - MUNICIPIO': linha_municipio['COD.'],
            'RUBRICA - MUNICIPIO': linha_municipio['RUBRICAS'],
        }

        nova_linha_quantitativo = { # Inicializa nova_linha_quantitativo aqui
            'COD - MUNICIPIO': linha_municipio['COD.'],
            'RUBRICA - MUNICIPIO': linha_municipio['RUBRICAS'],
            'QUANT. COINCIDÊNCIAS': len(coincidencias_h12) # Conta o número de coincidências
        }

        if not coincidencias_h12.empty:  # Verifica se houve coincidências
            nova_linha_coincidencias = dados_municipio_fixos.copy()
            nova_linha_resumo = dados_municipio_fixos.copy()
            nova_linha_incidencia = dados_municipio_fixos.copy()

            # Check if there's only one coincidence
            if len(coincidencias_h12) == 1:
                nova_linha_coincidencias_unicas = dados_municipio_fixos.copy() # Create row for "Coincidencias Unicas"

            # Remove NAT. RUBR. - MUNICIPIO para planilha Resumo
            del nova_linha_resumo['INC. CP - MUNICIPIO']
            del nova_linha_resumo['INC IRRF - MUNICIPIO']

            # Remove NAT. RUBR. - MUNICIPIO para planilha Incidencia
            del nova_linha_incidencia['NAT. RUBR. - MUNICIPIO']

            contador_coincidencia = 1
            for index_h12, linha_h12 in coincidencias_h12.iterrows():
                # Dados da planilha H12 para colunas dinâmicas (todas as planilhas)
                nome_rubrica_h12_coluna_correta = ' NOME'
                nat_rubr_h12_coluna_correta = ' natRubr'

                nova_linha_coincidencias[f'RUBRICA - H12 - {contador_coincidencia}'] = linha_h12[nome_rubrica_h12_coluna_correta]
                nova_linha_coincidencias[f'NAT. RUBR. - H12 - {contador_coincidencia}'] = linha_h12[nat_rubr_h12_coluna_correta]
                nova_linha_coincidencias[f'INC. CP - H12 - {contador_coincidencia}'] = linha_h12[' codIncCP']
                nova_linha_coincidencias[f'INC IRRF - H12 - {contador_coincidencia}'] = linha_h12[' codIncIRRF']

                nova_linha_resumo[f'RUBRICA - H12 - {contador_coincidencia}'] = linha_h12[nome_rubrica_h12_coluna_correta]
                nova_linha_resumo[f'NAT. RUBR. - H12 - {contador_coincidencia}'] = linha_h12[nat_rubr_h12_coluna_correta]

                nova_linha_incidencia[f'RUBRICA - H12 - {contador_coincidencia}'] = linha_h12[nome_rubrica_h12_coluna_correta]
                nova_linha_incidencia[f'INC. CP - H12 - {contador_coincidencia}'] = linha_h12[' codIncCP']
                nova_linha_incidencia[f'INC IRRF - H12 - {contador_coincidencia}'] = linha_h12[' codIncIRRF']

                nova_linha_rubricas[f'RUBRICA - H12 - {contador_coincidencia}'] = linha_h12[nome_rubrica_h12_coluna_correta] # Adiciona rubrica H12 para planilha Rubricas

                # For "Coincidencias Unicas" sheet, add H12 data if it's the only coincidence
                if len(coincidencias_h12) == 1:
                    nova_linha_coincidencias_unicas[f'RUBRICA - H12 - {contador_coincidencia}'] = linha_h12[nome_rubrica_h12_coluna_correta]
                    nova_linha_coincidencias_unicas[f'NAT. RUBR. - H12 - {contador_coincidencia}'] = linha_h12[nat_rubr_h12_coluna_correta]
                    nova_linha_coincidencias_unicas[f'INC. CP - H12 - {contador_coincidencia}'] = linha_h12[' codIncCP']
                    nova_linha_coincidencias_unicas[f'INC IRRF - H12 - {contador_coincidencia}'] = linha_h12[' codIncIRRF']


                contador_coincidencia += 1

            linhas_nova_planilha_coincidencias.append(nova_linha_coincidencias)
            linhas_nova_planilha_resumo.append(nova_linha_resumo)
            linhas_nova_planilha_incidencia.append(nova_linha_incidencia)
            linhas_nova_planilha_rubricas.append(nova_linha_rubricas)
            if len(coincidencias_h12) == 1: # Append to "Coincidencias Unicas" list only if there's one coincidence
                linhas_nova_planilha_coincidencias_unicas.append(nova_linha_coincidencias_unicas)

        else:  # Se não houve coincidências, adiciona à planilha "Sem Equivalencia"
            nova_linha_sem_equivalencia = dados_municipio_fixos.copy()
            linhas_nova_planilha_sem_equivalencia.append(nova_linha_sem_equivalencia)

        linhas_nova_planilha_quantitativo.append(nova_linha_quantitativo) # Adiciona a linha para planilha Quantitativo


    # 3. Criação dos DataFrames das novas planilhas
    df_nova_planilha_coincidencias = pd.DataFrame(linhas_nova_planilha_coincidencias)
    df_nova_planilha_resumo = pd.DataFrame(linhas_nova_planilha_resumo)
    df_nova_planilha_incidencia = pd.DataFrame(linhas_nova_planilha_incidencia)
    df_nova_planilha_sem_equivalencia = pd.DataFrame(linhas_nova_planilha_sem_equivalencia)
    df_nova_planilha_rubricas = pd.DataFrame(linhas_nova_planilha_rubricas)
    df_nova_planilha_quantitativo = pd.DataFrame(linhas_nova_planilha_quantitativo)
    df_nova_planilha_coincidencias_unicas = pd.DataFrame(linhas_nova_planilha_coincidencias_unicas) # DataFrame for "Coincidencias Unicas"

    # 4. Escrita das novas planilhas em Excel
    try:
        with pd.ExcelWriter(caminho_arquivo_excel, engine='openpyxl', mode='a') as writer: # 'a' para adicionar ao arquivo existente
            df_nova_planilha_coincidencias.to_excel(writer, sheet_name=nome_nova_planilha_coincidencias, index=False)
            df_nova_planilha_resumo.to_excel(writer, sheet_name=nome_nova_planilha_resumo, index=False)
            df_nova_planilha_incidencia.to_excel(writer, sheet_name=nome_nova_planilha_incidencia, index=False)
            df_nova_planilha_sem_equivalencia.to_excel(writer, sheet_name=nome_nova_planilha_sem_equivalencia, index=False)
            df_nova_planilha_rubricas.to_excel(writer, sheet_name=nome_nova_planilha_rubricas, index=False)
            df_nova_planilha_quantitativo.to_excel(writer, sheet_name=nome_nova_planilha_quantitativo, index=False)
            df_nova_planilha_coincidencias_unicas.to_excel(writer, sheet_name=nome_nova_planilha_coincidencias_unicas, index=False) # Escreve "Coincidencias Unicas"
        print(f"Planilhas '{nome_nova_planilha_coincidencias}', '{nome_nova_planilha_resumo}', '{nome_nova_planilha_incidencia}', '{nome_nova_planilha_sem_equivalencia}', '{nome_nova_planilha_rubricas}', '{nome_nova_planilha_quantitativo}' e '{nome_nova_planilha_coincidencias_unicas}' criadas com sucesso em '{caminho_arquivo_excel}'.")
    except Exception as e:
        print(f"Erro ao escrever as novas planilhas em Excel: {e}")


# Exemplo de uso:
caminho_do_excel = 'Rubricas Equivalentes.xlsx' # Substitua pelo caminho do seu arquivo Excel
planilha_municipio_nome = 'MUNICIPIO' # Substitua pelo nome correto da planilha 'MUNICIPIO'
planilha_h12_nome = 'H12' # Substitua pelo nome correto da planilha 'H12'

comparar_planilhas_excel(caminho_do_excel, planilha_municipio_nome, planilha_h12_nome)