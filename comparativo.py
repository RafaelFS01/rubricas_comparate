import pandas as pd
from fuzzywuzzy import fuzz
import re

# --- DICIONÁRIO DIRECIONADO: NOMES PADRÃO (DICIONÁRIO) -> VARIAÇÕES (PLANILHA MUNICIPAL) ---
dicionario_rubricas_direcionado = {
    "Salário-família": [
        "SALARIO FAMILIA",
        "Salario Família",
        "Salário Fam.",
        "Sal. Família Compl.",
        "S Família",
        "SALARIO-FAMILIA",
        "SALARIO FAMILIAR"
    ],
    "Adicional Noturno": [
        "ADICIONAL NOTURNO",
        "Adicional da Noite",
        "Noturno",
        "Adic. Noturno",
        "Adicional Not.",
        "ADICIONAL NOTURNO 20%",
        "ADICIONAL NOTURNO 25%"
    ],
    "Horas Extras": [
        "HORA EXTRA",
        "Hora Extra",
        "HE",
        "H. Extra",
        "Extras",
        "Hora Extra 50%",
        "Hora Extra 100%",
        "HORA EXTRA 50%",
        "HORA EXTRA 100%"
    ],
    "Vale Refeição": ["vr", "vale r", "refeitorio", "ticket refeição", "v.r.", "vale-refeicao", "VALE REFEICAO"],
    "Vale Alimentação": ["vale a", "alimentacao", "vale-alimentacao", "VALE ALIMENTACAO"],
    "Contribuição Sindical": ["SINDI"],
    "Remuneração Mensal": ["remuneração mensal", "remuneração mensais", "REMUNERAÇÃO MENSAL"],
    "Desconto": ["desc.", "desc", "dcto", "DESCONTO"],
    "13º salário": ["13° sal.", "13 salario"],
    "Adiantamento Salarial": ["adiantamento", "adiant", "adiantar", "antecipação", "ADIANTAMENTO SALARIAL"],
    "Repouso Semanal Remunerado": ["dsr", "r.s.r.", "descanso semanal", "DSR", "RSR", "DESCANSO SEMANAL REMUNERADO"],
    "Salário": ["salario", "salário", "salários", "SALARIO", "SALARIOS", "Remuneração mensal", "vencimento"] # ADICIONEI "VENCIMENTOS" como variação de "Salário" para teste
    # Adicione mais entradas para o seu dicionário aqui!
}


def pre_processar_nome_rubrica_dicionario_v3(nome, dicionario, planilha_origem="municipal"):
    """
    Pré-processa o nome da rubrica usando dicionário DIRECIONADO (Nomes Padrão -> Variações Municipais).
    Retorna o nome pré-processado e o NOME PADRÃO do dicionário que foi usado para padronizar (se houver).
    Se não houver padronização pelo dicionário, retorna None para o nome padrão.
    """
    nome_padrao_encontrado = None # Inicializa variável para rastrear se houve padronização e qual nome padrão

    try:
        nome_preprocessado = nome.lower()

        if planilha_origem == "municipal":  # APLICA O DICIONÁRIO APENAS PARA NOMES MUNICIPAIS (VARIAÇÕES)
            for nome_padrao, lista_variacoes_municipais in dicionario.items():
                for variacao_municipal in lista_variacoes_municipais:
                    if variacao_municipal.lower() in nome_preprocessado: # Verifica se a variação MUNICIPAL está CONTIDA no nome
                        nome_preprocessado = nome_preprocessado.replace(variacao_municipal.lower(), nome_padrao.lower())  # Padroniza para NOME PADRÃO (do dicionário)
                        nome_padrao_encontrado = nome_padrao # Registra o NOME PADRÃO que foi usado para padronizar
                        break # Importante: Para após a primeira padronização para evitar múltiplas substituições indesejadas (em casos específicos)
                if nome_padrao_encontrado: # Se encontrou e padronizou, já pode sair do loop externo (dicionário)
                    break

        nome_preprocessado = re.sub(r'[^\w\s-]', '', nome_preprocessado)
        nome_preprocessado = re.sub(r'\d+%', '', nome_preprocessado)
        nome_preprocessado = re.sub(r'\d+', '', nome_preprocessado)
        nome_preprocessado = re.sub(r'\s+', ' ', nome_preprocessado).strip()

        return nome_preprocessado, nome_padrao_encontrado # Retorna TUPLA: (nome preprocessado, nome_padrao_encontrado)

    except Exception as e:
        print(f"Erro ao pré-processar nome da rubrica: {nome}. Erro: {e}")
        return "", None # Retorna também None em caso de erro


def find_rubric_equivalences_v2(municipal_rubric_file, dicionario_rubricas, limiar_similaridade=80, bonus_dicionario=15): # ADICIONEI 'bonus_dicionario' como parâmetro
    """
    Encontra rubricas equivalentes usando fuzzy matching e dicionário DIRECIONADO.
    APRIMORAMENTO: Aplica um "bônus" no score de similaridade se o nome padrão for sugerido pelo dicionário.
    """
    try:
        df_municipal = pd.read_excel(municipal_rubric_file)

        df_municipal_renamed = df_municipal.rename(columns={'COD.': 'municipal_code', 'RUBRICAS': 'municipal_name'})
        df_municipal_selected = df_municipal_renamed[['municipal_code', 'municipal_name']].dropna(subset=['municipal_name']).copy()

        # --- MODIFICAÇÃO IMPORTANTE: Recebe tupla da função de pré-processamento ---
        df_municipal_selected[['municipal_name_processed', 'nome_padrao_sugerido']] = df_municipal_selected['municipal_name'].apply(lambda nome: pd.Series(pre_processar_nome_rubrica_dicionario_v3(nome, dicionario_rubricas, planilha_origem="municipal")))

        standard_rubric_names = list(dicionario_rubricas.keys())
        standard_rubric_names_processed = [pre_processar_nome_rubrica_dicionario_v3(nome, dicionario_rubricas, planilha_origem="padrao")[0] for nome in standard_rubric_names] # Pega só o nome preprocessado (índice 0 da tupla)

        equivalencias_fuzzy = []
        matched_municipal_codes = set()

        for index_municipal, row_municipal in df_municipal_selected.iterrows():
            melhor_score = 0
            melhor_match_standard_name = None

            for standard_name_index, standard_name in enumerate(standard_rubric_names_processed):
                score_base = fuzz.token_set_ratio(row_municipal['municipal_name_processed'], standard_name)
                score_final = score_base # Começa com o score base

                # --- APLICAÇÃO DO BÔNUS DO DICIONÁRIO ---
                if row_municipal['nome_padrao_sugerido'] is not None and row_municipal['nome_padrao_sugerido'] == standard_rubric_names[standard_name_index]: # VERIFICA se dicionário SUGERIU e se o nome padrão ATUAL é o SUGERIDO
                    score_final += bonus_dicionario # Aplica o bônus se a condição for verdadeira

                if score_final > melhor_score: # Usa 'score_final' para comparar agora (com ou sem bônus)
                    melhor_score = score_final
                    melhor_match_standard_name = standard_rubric_names[standard_name_index]

            if melhor_score >= limiar_similaridade:
                equivalencias_fuzzy.append({
                    'municipal_name': row_municipal['municipal_name'],
                    'standard_name': melhor_match_standard_name,
                    'score_similaridade': melhor_score # Salva o score FINAL (com bônus, se houver)
                })
                matched_municipal_codes.add(row_municipal['municipal_code'])

        equivalencias_fuzzy_df = pd.DataFrame(equivalencias_fuzzy)
        output_final_equivalencias = equivalencias_fuzzy_df[['municipal_name', 'standard_name', 'score_similaridade']].copy()

        unmatched_rubrics_list = []
        for index_municipal, row_municipal in df_municipal_selected.iterrows():
            if row_municipal['municipal_code'] not in matched_municipal_codes:
                unmatched_rubrics_list.append(row_municipal['municipal_name'])

        output_final_nao_correspondidas = pd.DataFrame({'Rubricas Não Correspondidas': unmatched_rubrics_list})

        return output_final_equivalencias, output_final_nao_correspondidas

    except FileNotFoundError:
        print("Erro: Arquivo Excel não encontrado.")
        return pd.DataFrame(), pd.DataFrame()
    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo: {e}")
        return pd.DataFrame(), pd.DataFrame()


# --- Exemplo de uso ---
municipal_rubric_file = 'MUNICIPIO RUBRICAS.xlsx'
limiar_similaridade = 80
bonus_dicionario = 15 # Define o valor do bônus do dicionário
arquivo_saida_excel = 'Rubricas Equivalentes_V5_BonusDicionario.xlsx' # Nome arquivo saída - Versão V5 - Com Bônus Dicionário

rubricas_equivalentes, rubricas_nao_correspondidas_df = find_rubric_equivalences_v2(
    municipal_rubric_file, dicionario_rubricas_direcionado, limiar_similaridade, bonus_dicionario # PASSA o 'bonus_dicionario' agora
)

if not rubricas_equivalentes.empty:
    print("Rubricas Equivalentes Encontradas e Salvas em:", arquivo_saida_excel, " (Versão 5 - Com Bônus Dicionário)")
    with pd.ExcelWriter(arquivo_saida_excel) as writer:
        rubricas_equivalentes.to_excel(writer, sheet_name='Equivalências', index=False)
        rubricas_nao_correspondidas_df.to_excel(writer, sheet_name='Rubricas Não Correspondidas', index=False)
else:
    print("Nenhuma rubrica equivalente encontrada (Versão 5 - Com Bônus Dicionário) ou ocorreu um erro.")


if not rubricas_nao_correspondidas_df.empty:
    print("\n--- Rubricas da Planilha Municipal SEM Similaridade Encontrada (abaixo do limiar) - Lista no Console: ---")
    for nome_rubrica_nao_correspondida in rubricas_nao_correspondidas_df['Rubricas Não Correspondidas']:
        print(f"- {nome_rubrica_nao_correspondida}")
else:
    print("\n--- Todas as rubricas da Planilha Municipal encontraram similaridade acima do limiar. ---")