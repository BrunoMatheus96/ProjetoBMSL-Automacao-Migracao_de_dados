# PLANEJAMENTO
"""
- [PLANILHA] - Pegar os dados das colunas 'ID', 'Funcionalidade', 'Step Name', 'Action' e 'Expected Result'

- [PLANILHA] - Correr a coluna 'Funcionalidade' e sempre que mudar o nome da funcionalidade, salvar o intervalo ID de cada funcionalidade. Ex: Funcionalidade 1: ID - 1 até 20

- [DOCUMENTO] - Pegar só o número no título

- Comparar os números do documento com os números da planilha, e quando encontrar os iguais, correr as colunas 'Action' e 'Expected Result' da planilha verificando os que se repetem em cada ID
e salvar esses dados em um dicionário, onde a chave é o número do documento e o valor é uma lista de dicionários, onde cada dicionário tem as chaves 'Action' e 'Expected Result' e os valores são as ações
e os resultados esperados que se repetem para cada ID.
Ex: {1: [{'Action': 'Ação 1', 'Expected Result': 'Resultado Esperado 1'}, {'Action': 'Ação 2', 'Expected Result': 'Resultado Esperado 2'}], 2: [{'Action': 'Ação 3', 'Expected Result': 'Resultado Esperado 3'}]}

- [DOCUMENTO] - Correr o documento da funcionalidade que tem os prints e verificar se os dados do dicionário estão presentes no documento, e quando encontrar, salvar o print correspondente a cada ação
e resultado esperado em um dicionário, onde a chave é o número do documento e o valor é uma lista de dicionários, onde cada dicionário tem as chaves 'Action', 'Expected Result' e 'Print' e os valores são as ações,
os resultados esperados e os prints correspondentes.

- [Documento] - Correr os outros documentos daquela funcionalidade, verificando se os steps repetidos estão presentes e adicionar os prints correspondentes a cada 'Action e 'Expected Result'
"""

# CÓDIGO