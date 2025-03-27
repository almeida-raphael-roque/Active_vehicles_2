import pandas as pd
from datetime import timedelta

today = pd.Timestamp.today().date()
yesterday = today - timedelta(days=1)
yesterday_str = yesterday.strftime('%Y-%m-%d')

path = r"C:\Users\raphael.almeida\OneDrive - Grupo Unus\analise de dados - Arquivos em excel\CAMPANHA_RANKING_ATIVACOES.xlsx"

df_ativ = pd.read_excel(path, sheet_name='ATIVAÇÕES')
df_cancel = pd.read_excel(path, sheet_name='CANCELAMENTOS')


def contar_placas(status, empresa):
    return len(
            df_ativ[
                (df_ativ['status']==status)&
                (df_ativ['empresa']==empresa)
            ]
        )


lista_empresas = ['Segtruck', 'Stcoop', 'Viavante']
lista_status = ['ATIVO', 'NOVO', 'RENOVAÇÃO', 'MIGRAÇÃO', 'REATIVAÇÃO', 'CANCELADO'] #talvez incluir reativação 


for empresax in lista_empresas:

    for statusx in lista_status:
        nome_variavel = f'{statusx}_{empresax}'

        if statusx == 'CANCELADO':
            globals()[nome_variavel] = len(df_cancel[
                (df_cancel['empresa']==empresax)&
                (df_cancel['data_cancelamento']==yesterday) #atenção à essa condição (verificar)
                ])
        elif statusx == 'ATIVO':
            globals()[nome_variavel] = len(df_ativ[df_ativ['empresa']==empresax])
        else:
            globals()[nome_variavel] = contar_placas(statusx, empresax)


#movimentações gerais --------------------------------------------------------------------------------
geral_novas = globals()['NOVO_Viavante'] + globals()['NOVO_Stcoop'] + globals()['NOVO_Segtruck']
geral_renovadas = globals()['RENOVAÇÃO_Viavante'] + globals()['RENOVAÇÃO_Stcoop'] + globals()['RENOVAÇÃO_Segtruck']
geral_reativacao = globals()['REATIVAÇÃO_Viavante'] + globals()['REATIVAÇÃO_Stcoop'] + globals()['REATIVAÇÃO_Segtruck']
geral_migracao = globals()['MIGRAÇÃO_Viavante'] + globals()['MIGRAÇÃO_Stcoop'] + globals()['MIGRAÇÃO_Segtruck']
geral_canceladas = globals()['CANCELADO_Viavante'] + globals()['CANCELADO_Stcoop'] + globals()['CANCELADO_Segtruck']


#placas gerais ---------------------------------------------------------------------------------------
geral_hoje = globals()['ATIVO_Segtruck']+globals()['ATIVO_Stcoop']+globals()['ATIVO_Viavante']
geral_segtruck = globals()['ATIVO_Segtruck']
geral_stcoop = globals()['ATIVO_Stcoop']
geral_viavante = globals()['ATIVO_Viavante']


#tabela dataframe ------------------------------------------------------------------------------------
índices = ['Novas', 'Renovadas', 'Reativadas', 'Migração', 'Canceladas', 'Total Empresa']
tabela = {
    'Segtruck': [globals()['NOVO_Segtruck'], globals()['RENOVAÇÃO_Segtruck'], globals()['REATIVAÇÃO_Segtruck'], globals()['MIGRAÇÃO_Segtruck'], globals()['CANCELADO_Segtruck'], geral_segtruck],
    'Stcoop': [globals()['NOVO_Stcoop'], globals()['RENOVAÇÃO_Stcoop'], globals()['REATIVAÇÃO_Stcoop'], globals()['MIGRAÇÃO_Stcoop'], globals()['CANCELADO_Stcoop'], geral_stcoop],
    'Viavante': [globals()['NOVO_Viavante'], globals()['RENOVAÇÃO_Viavante'], globals()['REATIVAÇÃO_Viavante'], globals()['MIGRAÇÃO_Viavante'], globals()['CANCELADO_Viavante'], geral_viavante],
    'Total': [geral_novas, geral_renovadas, geral_reativacao, geral_migracao, geral_canceladas, geral_hoje]
}

tabela_df = pd.DataFrame(tabela, index=índices)

tabela_df.applymap(lambda x: f'{x:,.0f}'.replace(',','.') if isinstance(x, (int,float)) else x)

tabela_df





