import pandas as pd
from openpyxl import load_workbook

caminho_arquivo = r'wm_task_(10).xlsx'
caminho_arquivo_procx = r'Gerencias_panda.xlsx'
caminho_arquivo_procx2 = r'Solicitações.xlsx'

df = pd.read_excel(caminho_arquivo, engine='openpyxl')
df_gerencias = pd.read_excel(caminho_arquivo_procx, engine='openpyxl')
df_outras_gerencias = pd.read_excel(caminho_arquivo_procx2, header=4, usecols='B:L')

def excluir_colunas(df):
    df = df.drop(columns=['Atribuída a', 'Aberto por', 'Ouvidoria'], inplace=True)
    return df

def abreviar_data(df):
    df['Criada em'] = df['Criada em'].dt.date

def ajustar_coluna_local(df):
    if 'Local' in df.columns:
        df[['Parte1', 'Parte2', 'Parte3','Parte4']] = df['Local'].str.split('/', expand=True, n=3)
   
    return df
def ajustar_imovel_impactado(df):
    df['Imóvel Impactado'] = df['Imóvel Impactado'].fillna('')
    df.loc[df['Imóvel Impactado'] == '', 'Imóvel Impactado'] = df['Parte2']
    df.loc[df['Imóvel Impactado'] == 'RJ', 'Imóvel Impactado'] = df['Parte4']

def apagar_cols(df):
    def apagar(*cols):
        for col in cols:
            df.drop(columns=[col], inplace=True)
    return apagar

def renomear_col(df):
    df.rename(columns={'Parte3':'Local'}, inplace=True)   

def add_locais(df,df3):
    df['Local'] = df['Local'].fillna('')
    df.loc[df['Local'] == '', 'Local'] = df3['Local']
    
def add_gerencias(df, df2):
    df = pd.merge(df, df2, on='Local', how='left')
    return df 
    
excluir_colunas(df)
abreviar_data(df)
ajustar_coluna_local(df)
ajustar_imovel_impactado(df)
apagar = apagar_cols(df)
apagar('Parte1', 'Parte2', 'Local','Parte4')
renomear_col(df)
add_locais(df, df_outras_gerencias)
df_com_gerencias = add_gerencias(df,df_gerencias)

df_com_gerencias.to_excel('Solicitações Operacionais DDC.xlsx', index=False)


