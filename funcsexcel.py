import numpy as np
import pandas as pd

path = r'c:\Users\est.angelo\Documents\Teste_real.xlsx'
a = pd.read_excel(path, header=1)
df = pd.DataFrame(a)

def vlookup(valor_ref, col_ref, col_ind, df_ref):

    indice = 0
    
    try:
        while indice <= len(df_ref[col_ref]):
            if valor_ref == df_ref.loc[indice,col_ref]:
                break
            indice+=1

        return df_ref.loc[indice,col_ind]
    except:
        return ''

# -------------------------------------------0------------------------------------------------

def calcular_cw(c, d, av):
    """
    Função que calcula a variável de onerosidade (cw) com base nas variáveis de entrada.
    
    Parâmetros:
    c (str): Valor da coluna C (001_aval_item_at).
    d (str): Valor da coluna D (002_aval_item_at_ag).
    av (str): Valor da coluna de avaliação onerosa (046_aval_oneros).
    
    Retorna:
    cw (str): Valor calculado da variável de onerosidade (115_ONEROSIDADE).
    """
    # Verifica se a avaliação onerosa está presente
    aval_onerosa = vlookup(f'{c}{d}', 'oneroso', 1, df)
    
    if aval_onerosa != "":
        cw = "ONEROSOS"
    else:
        cw = av
    
    return cw

# -------------------------------------------1------------------------------------------------

def calcular_cx(c, d):
    """
    Função que implementa a fórmula cx = IF(IFERROR(vlookup(_xlfn.CONCat($C;$D),oneroso,1,FALSE),"")<>"","atUALIZADOPARAONEROSOS","")

    Parâmetros:
    c (str): Valor da coluna c (001_aval_item_at).
    d (str): Valor da coluna d (002_aval_item_at_ag).

    Retorna:
    cx (str): Valor calculado conforme a fórmula.
    """

    # Implementação da fórmula
    concat_value = f"{c};{d}"
    lookup_result = vlookup(concat_value, 'oneroso', 1, df)

    if lookup_result != "":
        cx = "ATUALIZADO PARA ONEROSOS"
    else:
        cx = np.NaN

    return cx

# -------------------------------------------2------------------------------------------------

def calcular_cy(c, d, be):
    """
    Função que calcula o valor de fábrica unitário (cy) com base na fórmula fornecida.
    
    Parâmetros:
    c (float): Valor da coluna c (001_aval_item_at).
    d (float): Valor da coluna d (002_aval_item_at_ag).
    be (float): Valor da coluna be (055_aval_vlrfbc_unt).
    
    Retorno:
    cy (float): Valor calculado de VALOR_DE_FABRICA_UNITARIO (R$).
    """
    
    # Faz a busca usando vlookup e CONCat
    lookup_key = f"{c}_{d}"
    cy = vlookup(lookup_key, 'glosa', 56, df)
    if cy == '':
        cy = be
    return cy

# -------------------------------------------3------------------------------------------------

def calcular_cz(cy, bf):
    """
    Função que calcula o valor total da fábrica (cz) a partir do valor unitário de fábrica (CY) e da base de cálculo (BF).

    Parâmetros:
    cy (float): Valor unitário de fábrica (CY) em R$.
    bf (float): Base de cálculo (BF) para o cálculo do valor total.

    Retorno:
    cz (float): Valor total de fábrica (cz) em R$.
    """
    if cy == '' or cy == np.NaN:
        cy = 0
    cz = cy * bf
    return cz

# -------------------------------------------4------------------------------------------------

def calcular_da(c, d):
    """
    Calcula o valor da variável 'da' com base na fórmula do Excel.
    
    Essa função emula a fórmula de Excel:
    da = IF(IFERROR(vlookup(_xlfn.CONCat($C;$D),glosa,56,FALSE),"")<>"","VALORGLOSADO","")
    
    Parâmetros:
    c (str): Valor da coluna C (001_aval_item_at).
    d (str): Valor da coluna D (002_aval_item_at_ag).
    
    Retorna:
    da (str): Valor calculado da coluna da (119_FAB_MARCACAO_GLOSA_AGEPAR) com base na fórmula.
    """
    # Emulação da fórmula de Excel
    vlookup_result = vlookup(f'{c};{d}','glosa',56, df)
    
    if vlookup_result != "":
        da = "VALOR GLOSADO"
    else:
        da = np.NaN
    
    return da

# -------------------------------------------5------------------------------------------------

def calcular_db(cz, dc):
    """
    Função que calcula o custo adicional (db) a partir das variáveis cz e dc.
    
    Inputs:
    cz (float): Valor da coluna 118_VALOR_FABRICA_TOTAL (R$)
    dc (float): Valor da coluna 120_CUSTO_ADICIONAL (R$)
    
    Outputs:
    db (float): Custo adicional calculado usando a fórmula db = cz * dc
    """
    db = cz * dc
    return db

# -------------------------------------------6------------------------------------------------

def calcular_dc(bj, e, d):
    """
    Calcula a variável dc (121_CUSTO_ADICIONAL_CA_(%)) baseada em diferentes condições usando a tabela TabelaCA.

    Parâmetros:
    bj (float): Valor da variável 060_aval_ca_perc.
    e: 003_cont_gr
    d: 004_cont_munic

    Retorna:
    dc (float): Valor calculado da variável 121_CUSTO_ADICIONAL_CA_(%).
    """
    if bj == 0:
        dc = 0
    elif bj == d:
        dc = e
    else:
        dc = 0

    return dc

# -------------------------------------------7------------------------------------------------

def calcular_dd(cz, db, af, bg, bn, bs, ck, de):
    """
    Calcula o valor da variável dd com base em várias condições e variáveis.

    Parâmetros:
    cz (float): Valor total de fábrica (R$)
    db (float): Custo adicional (R$)
    af (float): Valor da coluna 030_cont_voc
    bg (float): Valor da coluna 057_aval_ea_vlr
    bn (str): Valor da coluna 064_aval_tp_ind_prec
    bs (float): Valor da coluna 069_aval_ind_prec_var
    ck (int): Valor da coluna 087_aval_analise
    de (float): Valor da coluna 123_JUROS_SOBRE_OBRAS_EM_ANDAMENTO_JOA_(%)

    Retorno:
    dd (float): Valor da coluna 122_JUROS_SOBRE_OBRAS_EM_ANDAMENTO_JOA (R$)
    """

    if bn == "IGP-M" and ck == 801:
        dd = cz * de
    elif bn == "IGP-M" and (ck == 802 or ck == 803 or ck == 804 or ck == 810 or ck == 811 or ck == 816):
        dd = af * bs * de
    elif de == 0:
        dd = 0
    else:
        dd = (cz + db + bg) * de

    return dd

# -------------------------------------------8------------------------------------------------

def calcular_de(bl):
    """
    Calcula o valor 'de' (123_JUROS_SOBRE_OBRAS_EM_ANDAMENTO_JOA_(%)) arredondando o valor de 'bl' (062_aval_joa_perc) para 8 casas decimais.

    Parâmetros:
    bl (float): Valor de 'bl' a ser arredondado.

    Retorna:
    float: Valor 'de' calculado após arredondar 'bl' para 8 casas decimais.
    """
    de = round(bl, 8)
    return de

# -------------------------------------------9------------------------------------------------

def calcular_df(be, bg, bm, bn, bs, bt, ck, cz, db, dd):
    """
    Calcula o valor df com base nas condições especificadas.

    Parâmetros:
    be (float): 055_aval_vlrfbc_unt
    bg (float): 057_aval_ea_vlr
    bm (float): 063_aval_vnr
    bn (str): 064_aval_tp_ind_prec
    bs (float): 069_aval_ind_prec_var
    bt (str): 070_aval_form_prec
    ck (int): 087_aval_analise
    cz (float): 118_VALOR_FABRICA_TOTAL (R$)
    db (float): 120_CUSTO_ADICIONAL (R$)
    dd (float): 122_JUROS_SOBRE_OBRAS_EM_ANDAMENTO_JOA (R$)

    Retorna:
    df (float): 124_VALOR_NOVO_DE_REPOSICAO_VNR_POS_GLOSA_REV_CA (R$)
    """
    if bt != "CCV - atualização do valor contábil" and bn == "IGP-M" and ck != 801:
        df = bm
    elif be == 0 and bn == "CONFORME CÁLCULO DE VAZÃO PROJETADA PARA 2030":
        df = bm
    else:
        try:
            df = (int(cz).astype + int(db) + int(dd) + int(bg)) * int(bs)
        except:
            df = np.NaN

    return df
    
# ------------------------------------------10------------------------------------------------

def calcular_dg(bt, cf):
    """
    Calcula o valor dg baseado na fórmula do Excel: dg = IF(BT="CCV-atualizaçãodovalorcontábil"; CF; 0)

    Parâmetros:
    bt (str): Valor da célula BT (070_aval_form_prec).
    cf (float): Valor da célula CF (082_aval_ccv).

    Retorna:
    dg (float): Valor calculado dg (124a_VALOR_CCV (R$)).
    """

    if bt == "CCV - atualização do valor contábil":
        dg = cf
    else:
        dg = 0

    return dg

# ------------------------------------------11------------------------------------------------

def calcular_dh(bm, bt, bu, cf, df, dg):
    """
    Calcula a depreciação conforme a fórmula do Excel.
    
    Inputs:
    bm : float - Valor da variável 063_aval_vnr (BM)
    bt : str - Valor da variável 070_aval_form_prec (BT)
    bu : float - Valor da variável 071_aval_dep_acum_reg_vlr (BU)
    cf : float - Valor da variável 082_aval_ccv (CF)
    df : float - Valor da variável 124_VALOR_NOVO_DE_REPOSICAO_VNR_POS_GLOSA_REV_CA (df)
    dg : float - Valor da variável 124a_VALOR_CCV (dg)
    
    Outputs:
    dh : float - Valor calculado da variável 125_DEPRECIACAO_POS_GLOSA_ACUMULADA (dh)
    """
    if bm != 0:
        dh = bu / bm * df
    elif cf != 0 and bt == "CCV - atualização do valor contábil":
        dh = bu / cf * dg
    else:
        dh = 0
    
    return dh

# ------------------------------------------12------------------------------------------------

def calcular_di(p, s, t):
    """
    Função que calcula a variável di com base na fórmula de Excel.

    Parâmetros:
    p (str): Valor da coluna P (014_cont_tipo).
    s (str): Valor da coluna S (017_cont_patrim).
    t (str): Valor da coluna T (018_cont_compl).

    Retorna:
    float: O valor calculado de di (126_qtde_total_BAR) ou 0 se houver um erro.
    """
    # Constrói a chave de pesquisa concatenando os valores de P, S e T.
    chave = f"{p}-{s}-{t}"

    # Tente encontrar a chave na tabela 'baixas' e retornar o valor da coluna 2 se encontrado.
    resultado = vlookup(chave, 'baixas', 2, df)
    if resultado == '':
        return 0
    else:
        return resultado

# ------------------------------------------13------------------------------------------------

def calcular_dj(p, s, t):
    """
    Calcula a variável dj (127_qtde_total_baixas) utilizando a fórmula do Excel:
    dj = IFERROR(vlookup(_xlfn.CONCat($P;"-";$S;"-";$T), baixas, 3, FALSE), 0)

    Parâmetros:
    p (str): Valor correspondente a 014_cont_tipo.
    s (str): Valor correspondente a 017_cont_patrim.
    t (str): Valor correspondente a 018_cont_compl.

    Retorna:
    dj (float): Valor calculado da variável dj (127_qtde_total_baixas).
    """

    # Simula a função vlookup do Excel usando um dicionário baixas
    baixas = "valor_014_cont_tipo"

    chave = f"{p}-{s}-{t}"

    # Constrói a chave de pesquisa usando os valores p, s e t
    chave_pesquisa = vlookup(chave, baixas, 3, df)

    # Tenta encontrar a chave no dicionário baixas, caso contrário, retorna 0
    if baixas == chave_pesquisa:
        dj = chave_pesquisa
        return dj
    else:
        return 0

# ------------------------------------------13------------------------------------------------

def calcular_dk(di, dj):
    """
    Calcula o valor de dk baseado na fórmula de Excel: dk = IF(DI<>0; ROUND(DJ/DI;8), 0)

    Parâmetros:
    di (float): Valor de DI (126_qtde_total_BAR)
    dj (float): Valor de DJ (127_qtde_total_baixas)

    Retorna:
    dk (float): Valor calculado de dk (128_fator_qtde_baixas)
    """

    if di != 0:
        try:
            dk = round(float(dj) / float(di), 8)
        except:
            dk = 0
    else:
        dk = 0
    
    return dk

# ------------------------------------------14------------------------------------------------

def calcular_dl(p, s, t):
    """
    Calcula o valor dl usando a fórmula Excel: dl = vlookup(_xlfn.CONCat($P;"-";$S;"-";$T), baixas, 4, FALSE)
    
    Parâmetros:
    p (str): Valor da coluna P (014_cont_tipo)
    s (str): Valor da coluna S (017_cont_patrim)
    t (str): Valor da coluna T (018_cont_compl)
    baixas (dict): Dicionário contendo os dados da tabela 'baixas', onde a chave é a combinação de P-S-T e o valor é a coluna 4
    
    Retorna:
    dl (str): Valor calculado de acordo com a fórmula, ou uma string vazia se não for possível calcular
    """
    chave = f"{p}-{s}-{t}"
    
    dl = vlookup(chave, 'baixas', 4, df)

    if dl == '':
        dl = np.NaN
    
    return dl

# ------------------------------------------15------------------------------------------------

def calcular_dm(y, di, dj):
    """
    Calcula a variável dm baseada na fórmula do Excel: dm = IFERROR((y / di) * (di - dj), 0).
-
    Parâmetros:
    y (float): Valor da coluna '023_cont_qtd'.
    di (float): Valor da coluna '126_qtde_total_BAR'.
    dj (float): Valor da coluna '127_qtde_total_baixas'.

    Retorna:
    float: Valor calculado da variável '130_Qtd_resultado_baixas'.
    """
    if di != 0:
        try:
            dm = (y / di) * (di - dj)
        except:
            dm = 0
    else:
        dm = 0
    return dm

# ------------------------------------------16------------------------------------------------

def calcular_dn(dk, df):
    """
    Calcula o valor de dn (131_VNR_efetivo_baixa (R$)) utilizando a fórmula dn = dk * df.
    
    Parâmetros:
    dk (float): Valor de dk (128_fator_qtde_baixas).
    df (float): Valor de df.
    
    Retorna:
    float: O valor calculado de dn.
    """
    dn = dk * df
    return dn

# ------------------------------------------17------------------------------------------------

def calcular_do(dk, dg):
    """
    Função que calcula a variável 'do' com base na fórmula do Excel: do = dk * dg
    
    Inputs:
    dk (float): Valor da variável dk (128_fator_qtde_baixas)
    dg (float): Valor da variável dg (131a_CCV_efetivo_baixa (R$))
    
    Outputs:
    do (float): Resultado do cálculo da variável 'do'
    """
    do = dk * dg
    return do

# ------------------------------------------17------------------------------------------------

def calcular_dp(dk, dh):
    """
    Função que calcula a variável dp (Depreciação Efetiva Baixas) com base nas variáveis dk e dh.

    Inputs:
    dk (float): Valor da variável dk (128_fator_qtde_baixas)
    dh (float): Valor da variável dh (alguma descrição de dh)

    Outputs:
    dp (float): Valor calculado da variável dp (132_Deprec_efetivo_baixas (R$))
    """
    dp = dk * dh
    return dp

# ------------------------------------------18------------------------------------------------

def calcular_dq(df, dn):
    """
    Calcula a diferença entre df e dn.

    Esta função calcula a diferença entre duas variáveis: df (124_VALOR_NOVO_DE_REPOSICAO_VNR_POS_GLOSA_REV_CA)
    e dn (133_VRN_resultado_baixas).

    Args:
        df (float): Valor de df (124_VALOR_NOVO_DE_REPOSICAO_VNR_POS_GLOSA_REV_CA).
        dn (float): Valor de dn (133_VRN_resultado_baixas).

    Returns:
        float: Valor da diferença entre df e dn (dq).
    """
    dq = df - dn
    return dq

# ------------------------------------------19-----------------------------------------------

def calcular_dr(dg, do):
    """
    Calcula a variável dr utilizando a fórmula dr = dg - DO.

    Parâmetros:
    dg (float): Valor da variável dg (124a_VALOR_CCV (R$)).
    do (float): Valor da variável DO (133a_CCV_resultado_baixas (R$)).

    Retorna:
    float: Valor da variável dr.
    """
    dr = dg - do
    return dr

# ------------------------------------------20-----------------------------------------------

def calcular_ds(dh, dp):
    """
    Calcula a variável ds (Diff_DS) com base na fórmula: ds = dh - ds
    
    Parâmetros:
    dh (float): Valor da coluna '125_DEPRECIACAO_POS_GLOSA_ACUMULADA (R$)'
    ds (float): Valor da coluna '134_Deprec_resultado_baixas (R$)'
    
    Retorno:
    float: Resultado do cálculo ds = dh - ds
    """
    resultado_ds = dh - dp
    return resultado_ds

# ------------------------------------------21-----------------------------------------------

def calcular_dt(dq, dr, ds):
    """
    Calcula a variável 'Residual_resultado_baixas' (R$) com base nas fórmulas do Excel.
    
    Parâmetros:
    dq (float): Valor da coluna '133_VRN_resultado_baixas (R$)'.
    dr (float): Valor da coluna '133a_CCV_resultado_baixas (R$)'.
    ds (float): Valor da coluna '134_Deprec_resultado_baixas (R$)'.
    
    Retorna:
    float: Valor calculado da coluna '135_Residual_resultado_baixas (R$)'.
    """

    dq = np.nan_to_num(dq)
    dr = np.nan_to_num(dr)
    ds = np.nan_to_num(ds)

    if dq - ds > 0:
        dt = dq - ds
    elif dr - ds > 0:
        dt = dr - ds
    else:
        dt = 0
    return dt

# ------------------------------------------22-----------------------------------------------

def calcular_du(cb, dq, dr, dt):
    """
    Função que calcula a nova taxa mensal (du) com base nas variáveis fornecidas.

    Parâmetros:
    cb (float): Valor da coluna 078_aval_vu_reman_reg.
    dq (float): Valor da coluna 133_VRN_resultado_baixas (R$).
    dr (float): Valor da coluna 133a_CCV_resultado_baixas (R$).
    dt (float): Valor da coluna 135_Residual_resultado_baixas (R$).

    Retorna:
    float: O valor da nova taxa mensal (du) calculado de acordo com a fórmula.
    """

    if cb == 0:
        return 0
    elif dq != 0:
        return round((dt / cb / dq), 8)
    elif dr != 0:
        return round((dt / cb / dr), 8)
    else:
        return 0
    
# ------------------------------------------23-----------------------------------------------

def calcular_dv(bt, dq, dr, ds, du):
    """
    Calcula o valor da variável dv com base nas variáveis fornecidas.

    :param bt: Valor da célula BT (070_aval_form_prec)
    :type bt: str
    :param dq: Valor da célula dq (133_VRN_resultado_baixas)
    :type dq: float
    :param dr: Valor da célula DR (133a_CCV_resultado_baixas)
    :type dr: float
    :param ds: Valor da célula DS (134_Deprec_resultado_baixas)
    :type ds: float
    :param du: Valor da célula DU (136_nova_taxa_MENSAL)
    :type du: float
    :return: Valor calculado da variável dv (137_DEPREC_dez_2020)
    :rtype: float
    """

    dq = np.nan_to_num(dq)
    dr = np.nan_to_num(dr)
    ds = np.nan_to_num(ds)
    du = np.nan_to_num(du)
    
    if bt != "CCV - atualização do valor contábil":
        if (((du * dq) * 60) + ds) > dq:
            return dq
        else:
            return (((du * dq) * 60) + ds)
    else:
        if (((du * dr) * 60) + ds) > dr:
            return dr
        else:
            return (((du * dr) * 60) + ds)
        
# ------------------------------------------23-----------------------------------------------

def calcular_dw(bt, dq, dr, dv):
    """
    Calcula o valor de dw com base na fórmula do Excel.

    Args:
    bt (float): Valor da coluna BT (070_aval_form_prec).
    dq (float): Valor da coluna dq (133_VRN_resultado_baixas (R$)).
    dr (float): Valor da coluna DR (133a_CCV_resultado_baixas (R$)).
    dv (float): Valor da coluna DV (137_DEPREC_dez_2020 (R$)).

    Returns:
    float: O valor calculado de DW (138_RESIDUAL_NOVO_DEZ_2020 (R$)).
    """
    if bt != "CCV - atualização do valor contábil":
        dw = dq - dv
    else:
        dw = dr - dv
    return dw

# ------------------------------------------24-----------------------------------------------

def calcular_dy(dq, dx):
    """
    Função que calcula a variável dy usando a fórmula do Excel.

    Inputs:
    dq (float): Valor da célula dq (133_VRN_resultado_baixas (R$))
    dx (float): Valor da célula dx (140_VRN_indice (R$))

    Outputs:
    dy (float): Resultado do cálculo dy, conforme a fórmula dy = dq * dx
    """
    dy = dq * dx
    return dy

# ------------------------------------------25-----------------------------------------------

def calcular_dz(dr, dx):
    """
    Função que calcula dz com base na fórmula dz = DR * dx.

    Parâmetros:
    dr (float): Valor de DR (133a_CCV_resultado_baixas em R$).
    dx (float): Valor de dx (139_FatOR_INDICE).

    Retorna:
    float: O valor calculado de dz (140a_CCV_indice em R$).
    """
    dz = dr * dx
    return dz

# ------------------------------------------26-----------------------------------------------

def calcular_ea(dx, dv):
    """
    Função que calcula o valor de 'ea' com base na fórmula ea = dx * DV.

    Args:
        dx (float): O valor da coluna '139_FatOR_INDICE'.
        dv (float): O valor da coluna '137_DEPREC_dez_2020 (R$)'.

    Returns:
        float: O valor calculado de 'ea' (141_DEPREC_ACUMULADA_indice (R$)').
    """
    ea = dx * dv
    return ea

# ------------------------------------------27-----------------------------------------------

def calcular_eb(ea, ee):
    """
    Função que calcula a variável eb com base nos valores de ea e ee.

    Inputs:
    ea (float): Valor da coluna EA (141_DEPREC_ACUMULADA_indice (R$)).
    ee (float): Valor da coluna ee (144_INDICE_DE_APROVEITAMENTO_IA_%).

    Outputs:
    eb (float): Valor da coluna EB (141a_DEPREC_ACUMULADA_indice (R$)_DedUZIDA_DO_IA_(R$)).

    Parâmetros:
    - ea: O valor da coluna EA, que representa 141_DEPREC_ACUMULADA_indice (R$).
    - ee: O valor da coluna ee, que representa 144_INDICE_DE_APROVEITAMENTO_IA_%.

    Retorna:
    O valor calculado da coluna EB com base nos valores de EA e ee. Se EA for uma string vazia, o resultado será uma string vazia.
    Se ee for uma string vazia, o resultado será igual a EA. Caso contrário, o resultado será EA multiplicado por ee.
    """
    if ea == "":
        return ""
    elif ee == "" or "NaN":
        return ea
    else:
        return ea * ee

    
# ------------------------------------------28-----------------------------------------------

def calcular_ec(ea, bt, dy, dz):
    """
    Função que calcula a variável 142_DEPRECIACAO_ACUMULADA_RegULatORIA_%

    Inputs:
    ea (float): 141_DEPREC_ACUMULADA_indice (R$)
    bt (str): 070_aval_form_prec
    dy (float): 140_VRN_indice (R$)
    dz (float): 140a_CCV_indice (R$)

    Outputs:
    ec (float): 142_DEPRECIACAO_ACUMULADA_RegULatORIA_%
    """
    if ea != 0:
        if bt != "CCV - atualização do valor contábil":
            ec = round(ea / dy, 8)
        else:
            ec = round(ea / dz, 8)
    else:
        ec = 0

    return ec

# ------------------------------------------29-----------------------------------------------

def calcular_ed(bt, dy, dz, ea):
    """
    Calcula a variável ed (143_VALOR_DE_MERCADO_EM_USO_VMU_R$) com base na fórmula do Excel.

    Parâmetros:
    bt (str): Valor da célula BT (070_aval_form_prec).
    dy (float): Valor da célula dy (140_VRN_indice (R$)).
    dz (float): Valor da célula DZ (140a_CCV_indice (R$)).
    ea (float): Valor da célula EA (141_DEPREC_ACUMULADA_indice (R$)).

    Retorna:
    float: Valor calculado da variável ed (143_VALOR_DE_MERCADO_EM_USO_VMU_R$).
    """
    if bt != "CCV - atualização do valor contábil":
        ed = dy - ea
    else:
        ed = dz - ea

    return ed

# ------------------------------------------29-----------------------------------------------

def calcular_eg(dy, ee):
    """
    Função que calcula a Parcela Não Aproveitável do VNR (Valor de Referência Nacional).

    Inputs:
    dy (float): Valor da dy (140_VRN_indice (R$)).
    ee (float): Valor do Índice de Aproveitamento IA (%) (144_INDICE_DE_APROVEITAMENTO_IA_%).

    Outputs:
    eg (float): Valor da Parcela Não Aproveitável do VNR (146_PARCELA_NAO_APROVEITAVEL_DO_VNR (R$)).
    """

    if ee == "NaN" or "":
        eg = ""
        return eg
    else:
        eg = dy - (dy * (ee / 100))
        return eg

# ------------------------------------------30-----------------------------------------------

def calcular_eh(eg, dy):
    """
    Função que calcula a diferença entre dy e eg, ou retorna dy se eg for uma string vazia.

    Inputs:
    eg (str): Valor da célula eg correspondente à coluna 146_PARCELA_NAO_APROVEITAVEL_DO_VNR (R$).
    dy (float): Valor da célula dy correspondente à coluna 140_VRN_indice (R$).

    Outputs:
    eh (float): Resultado do cálculo da fórmula. Se eg for uma string vazia, retorna dy; caso contrário, retorna dy - eg.
    """

    if eg == "" or "NaN":
        eh = dy
        return eh
    else:
        eh = dy - float(eg)  # Converte eg para float se for um número
        return eh

# ------------------------------------------31-----------------------------------------------

def calcular_ei(ed, ee):
    """
    Função que calcula a variável EI (PARCELA_NAO_APROVEITAVEL_DO_VMU) com base nos valores de ed e ee.

    Inputs:
    ed (float): Valor de mercado em uso (VMU) em R$ (143_VALOR_DE_MERCADO_EM_USO_VMU_R$).
    ee (float): Índice de aproveitamento (IA) em porcentagem (144_INDICE_DE_APROVEITAMENTO_IA_%).

    Outputs:
    EI (float): Parcela não aproveitável do VMU em R$ (148_PARCELA_NAO_APROVEITAVEL_DO_VMU (R$)).
    
    Formula:
    EI = ed if ee == '' else ed - (ed * (ee / 100))
    """
    if ee == '':
        return ed
    else:
        ee = float(ee)  # Certifique-se de que ee seja um número
        return ed - (ed * (ee / 100))

# ------------------------------------------32-----------------------------------------------

def calcular_ej(ed, ei):
    """
    Calcula o Valor de Mercado em Uso Deduzido do IA (R$) com base nos valores de ed e EI.

    Parâmetros:
    ed (float): Valor de Mercado em Uso (R$) - 143_VALOR_DE_MERCADO_EM_USO_VMU_R$
    ei (float): Parcela Não Aproveitável do VMU (R$) - 148_PARCELA_NAO_APROVEITAVEL_DO_VMU (R$)

    Retorno:
    float: O Valor de Mercado em Uso Deduzido do IA (R$) calculado de acordo com a fórmula:
           ej = IF(EI="", ed, ed - EI)
    """
    if ei == "" or "NaN":
        return ed
    else:
        return ed - ei

# ------------------------------------------33-----------------------------------------------

def calcular_ek(cb, ec):
    """
    Função que calcula a taxa de depreciação remanescente anual.

    Args:
        cb (float): Valor da coluna 078_aval_vu_reman_reg.
        ec (float): Valor da coluna 142_DEPRECIACAO_ACUMULADA_RegULatORIA_%.

    Returns:
        float: A taxa de depreciação remanescente anual, arredondada para 8 casas decimais.
    """
    if cb - 60 > 0:
        taxa = round((1 - ec) / ((cb - 60) / 12), 8)
    else:
        taxa = 0
    return taxa

# ------------------------------------------34-----------------------------------------------

def calcular_el(ek):
    """
    Calcula o valor de EL com base no valor de EK de acordo com a fórmula do Excel.

    Args:
        ek (float): Valor de EK (150_taxa_de_depreciacao_remanescente_anual).

    Returns:
        float: Valor calculado de EL (151_VIDA_UTIL_RegULatORIA_[ANOS]).
    """
    if ek != 0:
        el = round(1 / ek, 8)
    else:
        el = 0
    return el

# ------------------------------------------35-----------------------------------------------

def calcular_em(at, dl):
    """
    Função que calcula a elegibilidade com base nos valores de at e dl.

    Inputs:
    - at (str): A coluna at (044_aval_eleg) da tabela.
    - dl (str): A coluna dl (129_BAIXAS) da tabela.

    Outputs:
    - EM (str): A coluna EM (152_ELegIBILIDADE) calculada com base nas condições.

    Exemplo de Uso:
    EM = calcula_elegibilidade("044_aval_eleg", "129_BAIXAS")
    """

    if dl == "TOTAL":
        em = "NÃO ELEGÍVEIS"
    elif at == 0:
        em = ""
    else:
        em = at

    return em

# ------------------------------------------36-----------------------------------------------
