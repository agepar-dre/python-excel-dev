# Importando módulos necessários
import inspect
import funcsexcel as fxml  # Importa um módulo funcsexcel, presumivelmente personalizado
import pandas as pd
import numpy as np

# Função que identifica se um valor é um número
def identifica_numero(coluna):
    for valor in coluna:
        if valor == np.NaN:
            return False
        try:
            float(valor)
            return True
        except:
            return False

# Definindo uma função para listar os argumentos necessários para chamar uma função
def listar_argumentos(funcao):
    
    assinatura = inspect.signature(funcao)
    parametros = assinatura.parameters
    argumentos_necessarios = []

    # Iterando pelos parâmetros da função
    for parametro, info in parametros.items():
        if info.default == inspect.Parameter.empty:
            argumentos_necessarios.append(parametro)

    return argumentos_necessarios

# Definindo uma função para verificar se uma coluna tem uma fórmula
def has_formula(x):
    with open(r"c:\Users\est.angelo\Documents\Teste.txt", 'r') as file:
        v = False
        for line in file:
            s = f'Coluna {x} '
            if s in line:
                v = True
        return v

# Definindo uma função para extrair linhas de um arquivo e criar um dicionário com os dados
def extract_lines(tdict):
    extracting = False

    with open(r"c:\Users\est.angelo\Documents\Teste.txt", 'r') as file:
        for line in file:
            if extracting:
                if '------------------------------' in line:
                    break
                l = line.split(': ')
                tdict[l[0]] = l[1].replace('\n','')
            elif 'Dicionario de dados' in line:
                extracting = True

# Criando um dicionário para armazenar dados extraídos
dict_dados = {}
extract_lines(dict_dados)

# Criando um dicionário para armazenar informações de funções
dict_func = {}

# Iterando pelos dados extraídos
for i in dict_dados:
    if has_formula(i):
        stri = 'fxml.calcular_'+i

        # Listando os argumentos necessários para a função
        dfi = listar_argumentos(eval(stri))
        atts = []

        # Criando uma lista de argumentos formatados
        for a in dfi:
            atts.append(f'df_aux[dict_dados[\'{a}\']]')

        dict_func[i] = atts

# Caminho do arquivo Excel de entrada
path = r'c:\Users\est.angelo\Documents\Teste_real.xlsx'
a = pd.read_excel(path, header=1)
df = pd.DataFrame(a)
df_aux = df.copy()

# Abrindo um arquivo de saída para escrever resultados
with open(r"c:\Users\est.angelo\Documents\RelAngelo1.txt", 'w') as f:

    # Iterando pelos dados extraídos
    for x,y in dict_dados.items():
        col_want = y
        
        # Se a coluna não tiver uma fórmula, continue para a próxima
        if x not in dict_func:
            continue

        try:
            # Construindo uma string de função e avaliando-a
            attr_str = ', '.join(map(str, dict_func[x]))
            func_str = f'np.vectorize(fxml.calcular_{x})({attr_str})'
            df_aux[col_want] = eval(func_str)
        except Exception as exception:
            print(str(exception))
            f.write(x + ' : ' + 'algo errado (exception/error)\n')
            continue

        # Verificando se as colunas são iguais ou têm valores nulos
        if df_aux[col_want].astype(str).equals(df[col_want].astype(str)):
            f.write(x + ' : ' + 'tudo certo\n')
        elif identifica_numero(df_aux[col_want]) and identifica_numero(df[col_want]):
            if df_aux[col_want].astype(int, errors='ignore').equals(df[col_want].astype(int, errors='ignore')):
                f.write(x + ' : ' + 'tudo certo\n')
        elif (df[col_want].isnull().sum() + df_aux[col_want].isnull().sum()) == (len(df[col_want]) + len(df_aux[col_want])):
            f.write(x + ' : ' + 'tudo certo\n')
        else:
            f.write(x + ' : ' + 'algo errado\n')


print(df[dict_dados['dl']])
print(df_aux[dict_dados['dl']])