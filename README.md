# Automatização de preenchimento de dados da API ViaCEP em planilha Excel

Este repositório contém um código em Python para automatizar o preenchimento de dados em uma planilha Excel utilizando a API ViaCEP. O script percorre uma coluna de CEPs na planilha, consulta a API para obter os dados relacionados a cada CEP e preenche os resultados na planilha. Essa automatização é útil para agilizar a obtenção de informações completas sobre endereços a partir de CEPs.

## Pré-requisitos
- Python 3.x
- Bibliotecas: pandas, requests

Certifique-se de ter o Python instalado em sua máquina local juntamente com as bibliotecas mencionadas acima.

## Como usar
1. Clone este repositório em sua máquina local ou faça o download dos arquivos.

2. Certifique-se de ter uma planilha Excel com uma coluna de CEPs. O formato esperado da planilha é `.xlsx`, mas você pode ajustar o código conforme necessário para outros formatos suportados pelo `pandas`.

3. Abra o arquivo `preenchimento_api.py` em um editor de texto ou ambiente de desenvolvimento Python.

4. No arquivo `preenchimento_api.py`, atualize o caminho do arquivo da planilha Excel na linha `planilha = pd.read_excel('caminho/para/a/sua/planilha.xlsx')` para o caminho correto do seu arquivo Excel.

5. Execute o script Python `preenchimento_api.py`. Ele irá percorrer a coluna de CEPs na planilha, fazer consultas na API ViaCEP e preencher os dados retornados na planilha.

6. Após a execução do script, um novo arquivo Excel com o nome `completo.xlsx` será gerado, contendo a planilha original com os dados da API preenchidos.

7. Verifique o arquivo `completo.xlsx` para visualizar os resultados obtidos.

## Observações
- Certifique-se de ter uma conexão com a internet durante a execução do script, pois ele faz consultas na API ViaCEP para obter os dados dos CEPs.

- Caso a API ViaCEP não retorne dados para um CEP específico, será preenchido um dicionário vazio para esse CEP na planilha.

- É possível ajustar e personalizar o código conforme necessário para atender às suas necessidades específicas.
