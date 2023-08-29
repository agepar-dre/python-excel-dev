## Scripts de Processamento de Dados do Excel
Este repositório contém scripts em Python para processamento de dados do Excel. Os scripts incluem funções personalizadas para manipular dados do Excel e verificar expressões de fórmulas em colunas.

# Arquivos
TstExcel.py
TstExcel.py é o script principal responsável pelo processamento de dados do Excel. Ele executa as seguintes tarefas:

Importa Módulos Necessários: Importa módulos como funcsexcel, pandas e numpy para lidar com o processamento de dados.

Define Funções de Utilidade:

identifica_numero(coluna): Identifica se um valor em uma coluna é um número.
listar_argumentos(funcao): Lista os argumentos necessários para uma função específica.
has_formula(x): Verifica se uma coluna possui uma fórmula com base nas informações em Teste.txt.
extract_lines(tdict): Extrai dados de Teste.txt e cria um dicionário.
Processamento de Dados:

Lê um arquivo do Excel chamado Teste_real.xlsx e cria um DataFrame.
Itera pelos dados extraídos de Teste.txt.
Constrói e avalia funções para colunas com fórmulas.
Compara as colunas modificadas com as originais e registra o status em um arquivo de saída chamado RelAngelo1.txt.
Trata exceções e escreve mensagens de erro no arquivo de saída.
funcsexcel.py
funcsexcel.py parece ser um módulo personalizado usado em TstExcel.py. Provavelmente, ele contém funções personalizadas (calcular_*) que são chamadas com base no conteúdo das colunas no arquivo do Excel. Os detalhes exatos das funções não são fornecidos, mas elas parecem ser essenciais para a manipulação de dados em TstExcel.py.

Teste.txt
Teste.txt é um arquivo de texto usado como fonte de dados para TstExcel.py. Ele contém informações sobre as colunas do Excel, incluindo se elas possuem fórmulas. Este script lê este arquivo para determinar quais colunas precisam de processamento especial.

Uso
Para usar esses scripts, siga estas etapas:

Coloque TstExcel.py, funcsexcel.py, Teste.txt e o arquivo do Excel (Teste_real.xlsx) no mesmo diretório.

Certifique-se de ter o Python instalado junto com as bibliotecas necessárias (pandas e numpy).

Execute TstExcel.py. Ele irá processar os dados, aplicar funções personalizadas e gerar um arquivo de saída chamado RelAngelo1.txt.

Notas Adicionais
A funcionalidade exata e o propósito de funcsexcel.py não são fornecidos nesta documentação. Para entender completamente seu papel, consulte a implementação específica das funções calcular_* dentro do arquivo.

Certifique-se de que os caminhos para os arquivos de entrada e saída estejam corretamente especificados dentro de TstExcel.py.
