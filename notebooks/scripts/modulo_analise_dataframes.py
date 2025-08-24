import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet


def getValores_Col(
        letraCol: str, 
        sheet: Worksheet):
    
    coluna_escolhida = \
        [cell[0].value for cell in sheet[f'{letraCol}7':f'{letraCol}199']]
    return coluna_escolhida


def setListaDeListas(array: list):
    listaDividida = \
    [array[i:i+6] for i in range(0, len(array), 6)]
    return listaDividida


def turnIntoDf(
        listaDividida: list, 
        nomesColunas = ['C1','C2','C3','C4','C5','C6']):
    
    novo_df = pd.DataFrame(
        listaDividida, 
        columns=nomesColunas)
    return novo_df


# Essa função basicamente faz todas as três acima, porém de uma  só vez:
def reorganizeColumn(
        letraCol: str, 
        cabecalho: list, 
        sheet_excel: Worksheet):
    
    arrayVls = getValores_Col(letraCol, sheet_excel)
    listaDividida = setListaDeListas(arrayVls)
    return turnIntoDf(listaDividida, cabecalho)


def concatenarDfs(
        dfA: pd.DataFrame, 
        dfB: pd.DataFrame):
    return pd.concat(
        [dfA, dfB], axis=1)


def buildColunaPreenchida(
        vlHeaderColuna: str,
        vlToFillColuna: str, 
        df_usado: pd.DataFrame):
    colunaCriada = [vlToFillColuna for i in range(len(df_usado))]
    return pd.DataFrame(
        colunaCriada, 
        columns=[vlHeaderColuna])


def merge_total_governo_privado_tpArmazem_dt_registro(
        dfA: pd.DataFrame, 
        dfB: pd.DataFrame, 
        dfC: pd.DataFrame,
        tp_armazens_typed: str,
        vl_coluna_armazens: str,
        tp_armazens_typed2: str,
        desc_semestre_ano: str,
        df_base: pd.DataFrame):
    
    join1 = concatenarDfs(dfA, dfB)
    join2 = concatenarDfs(join1, dfC)
    
    df_ctg_armazens = buildColunaPreenchida(
        tp_armazens_typed, 
        vl_coluna_armazens, 
        df_base)
    
    join3 = concatenarDfs(join2, df_ctg_armazens)
    df_semestre_ano = buildColunaPreenchida(
        tp_armazens_typed2, 
        desc_semestre_ano, 
        df_base)
    
    join4 = concatenarDfs(join3, df_semestre_ano)
    return join4