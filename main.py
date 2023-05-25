import os
import pandas as pd
import requests


# Função para consultar a API ViaCEP
def consultar_cep(cep):
    try:
        url = f'https://viacep.com.br/ws/{cep}/json/'
        response = requests.get(url)
        if response.status_code == 200:
            return response.json()
        else:
            return None
    except requests.exceptions.RequestException as e:
        print(f"Ocorreu um erro na consulta do CEP {cep}: {e}")
        return None


# Função principal para preencher os dados da API na planilha
def preencher_dados_api_na_planilha():
    # Obter o caminho absoluto para a planilha
    caminho_planilha = os.path.abspath("CEPS.xlsx")

    try:
        # Ler a planilha Excel, especificando que a coluna 'CEP' deve ser lida como strings
        planilha = pd.read_excel(caminho_planilha, dtype={'CEP': str})

        # Percorrer as linhas da planilha
        for index, row in planilha.iterrows():
            cep = row['CEP']

            # Validar o CEP
            if len(str(cep)) == 8 and str(cep).isdigit():
                # Consultar a API ViaCEP
                resultado = consultar_cep(cep)

                # Verificar se houve resultado
                if resultado is not None:
                    # Preencher os dados na planilha
                    planilha.at[index, 'Endereço'] = resultado.get('logradouro', '')
                    planilha.at[index, 'Bairro'] = resultado.get('bairro', '')
                    planilha.at[index, 'Cidade'] = resultado.get('localidade', '')
                    planilha.at[index, 'Estado'] = resultado.get('uf', '')
                    print(f"Preenchimento do cep {cep} - OK")
            else:
                # Se o CEP for inválido, preencher com valores vazios
                planilha.at[index, 'Endereço'] = ''
                planilha.at[index, 'Bairro'] = ''
                planilha.at[index, 'Cidade'] = ''
                planilha.at[index, 'Estado'] = ''

        # Salvar as alterações na planilha
        planilha.to_excel(caminho_planilha, index=False)

        print(f"Preenchimento {index + 1} dos dados concluído com sucesso!")

    except FileNotFoundError:
        print("Arquivo da planilha não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro durante o preenchimento dos dados: {e}")


# Executar a função principal
preencher_dados_api_na_planilha()