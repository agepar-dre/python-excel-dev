import pandas as pd
import openpyxl
import re

def remove_lines(filename, target_line):
    with open(filename, 'r') as file:
        lines = file.readlines()

    with open(filename, 'w') as file:
        i = 0
        while i < len(lines):
            if lines[i].strip() == target_line:
                i += 3  # Pula a linha atual, anterior e posterior
            elif i > 0 and i + 2 < len(lines) and lines[i + 1].strip() == target_line:
                i += 3  # Pula a linha atual, anterior e posterior
            else:
                file.write(lines[i])
                i += 1

def verifica_formula(texto):
    if texto.startswith('='):
        if texto[1].isalpha():
            return True
        else:
            return False
    else:
        return False

def substituir_virgula_por_ponto_virgula(texto):
    padrao = r'\((.*?)\)'  # Padrão para encontrar o texto entre parênteses
    resultado = re.sub(padrao, lambda match: match.group().replace(',', ';'), texto)
    return resultado

def substituir_ponto_por_virgula(texto):
    padrao = r'\d+\.\d+'
    resultado = re.sub(padrao, lambda match: match.group().replace('.', ','), texto)
    return resultado

def remove_space(string):
    return "".join(string.split())

def remove_num(string):
    return re.sub(r'\d+', '', string)

def int_to_col_letter(n):
    result = ''
    while n >= 0:
        n, remainder = divmod(n, 26)
        result = chr(remainder + 97) + result
        n -= 1
    return result

def trocar_por_chave(val, dct):
    for chave, valor in dct.items():
        if valor == val:
            return chave

def extract_lines(filename, target_variable):
    lines_to_extract = []
    extracting = False

    with open(filename, 'r') as file:
        for line in file:
            if extracting:
                if '------------------------------' in line:
                    break
                lines_to_extract.append(line)
            elif f'Coluna {target_variable}' in line:
                extracting = True

    return lines_to_extract

def gerar_prompt(f):
    prompt = 'faça uma função em python para a formula de excel: ' + f + '\n'
    prompt += 'insira uma documentação da função indicando a correspondência de cada variável, preservando os nomes originais escritos na tabela abaixo, conforme o exemplo:\n'

    for x,y in dict_coluna.items():
        if  x == f[0:f.find('=')] or re.search(f'\W{x.upper()}\W', f):
            prompt += x+ ' : ' + y + '\n'
    
    prompt+= '''Exemplo:
    def função(inputs):
	    #Função que calcula a variável x
	    #Inputs
	    #x: coluna x
	    #Outputs
	    #a: coluna a
	    #b: coluna b
	    x = a + b
	    return x
        
        '''

    return prompt


REGEX_VAR = '(?<![\'":])([A-Z]{1,3}[0-9]{1,7})(?![\'":])'

path = r'c:\Users\est.angelo\Documents\Teste_real.xlsx'
path_txt = r"c:\Users\est.angelo\Documents\Teste.txt"
a = pd.read_excel(path, header=1)


# Carregar o arquivo Excel
wb = openpyxl.load_workbook(filename=path)
sheet_names = wb.sheetnames
name = sheet_names[0]
sheet_ranges = wb[name]
vals = sheet_ranges.values

# Criar o DataFrame a partir dos valores e usar a primeira linha como nomes das colunas
df = pd.DataFrame(vals)
df = df.dropna(axis=1, how='all')
df.drop(index=0, axis=0, inplace=True)
df.columns = a.columns

dict_coluna = {}

# Abrir o arquivo de texto para escrita
with open(path_txt, "w") as f:
    f.write('Dicionario de dados\n')
    colunas_i = []

    for col_name, col_values in df.items():
        colunas_i.append(col_name)

    for idx, c in enumerate(colunas_i):
        letra = int_to_col_letter(idx)
        dict_coluna[letra] = c
    
    for letra, coluna in dict_coluna.items():
        f.write(f"{letra}: {coluna}\n")
    
    f.write('-' * 30 + '\n') 
    f.write('Formulas:\n')

    # Iterar sobre as colunas e seus valores
    for col_name, col_values in df.items():
        written = set()
        have_l = False
        col_name_print = trocar_por_chave(col_name, dict_coluna)
        f.write(f"Coluna {col_name_print} ({col_name}):\n")

        # Iterar sobre os valores da coluna
        for cell_value in col_values:
            
            if isinstance(cell_value, str):
                
                cell_value = remove_space(cell_value)
                
                var = re.findall(REGEX_VAR, cell_value)
                
                for v in var:
                    if v in cell_value:
                        cell_value = cell_value.replace(v, remove_num(v))

                cell_value = substituir_ponto_por_virgula(cell_value)
                cell_value = substituir_virgula_por_ponto_virgula(cell_value)

                if cell_value not in written and verifica_formula(cell_value):
                    f.write(col_name_print + cell_value)
                    f.write('\n')
                    written.add(cell_value)
                    have_l = True

        if not have_l:
            f.write('Não há fórmulas nesta coluna.\n')
        f.write('-' * 30 + '\n')


remove_lines(path_txt, 'Não há fórmulas nesta coluna.')

# Fechar o arquivo de texto
print("Processamento concluído e o resultado foi salvo em 'Teste.txt'.\n\n\n")

for l in extract_lines(path_txt, 'eb'):
    print(gerar_prompt(l.strip()))