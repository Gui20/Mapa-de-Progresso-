import pandas as pd
import sqlalchemy
from functools import reduce


def conect_sql_server():
    engine = sqlalchemy.create_engine(
        'mssql+pymssql://admin:Ioutility2020!@database-1.cz9ylvwjzzvp.sa-east-1.rds.amazonaws.com/metodo')
    return engine


# Tentando conectar ao servidor
try:
    e = conect_sql_server()
    furodirecional_df = pd.read_sql_table('furodirecional', e)
    ramal_df = pd.read_sql_table('ramal', e)
    instalacao_df = pd.read_sql_table('instalacao', e)
    vistoria_df = pd.read_sql_table('vistoria', e)
    ligacao_df = pd.read_sql_table('ligacao', e)
except sqlalchemy.except_.NoSuchModuleError:
    print("Deu erro ao conectar no servidor!")
except sqlalchemy.except_.ArgumentError:
    print("Não da pra analisar, possivelmente erro de pontuação da URL!")
except sqlalchemy.except_.OperationalError:
    print("Não foi possível conectar com o banco de dados!")
finally:
    e = conect_sql_server()
    furodirecional_df = pd.read_sql_table('furodirecional', e)
    ramal_df = pd.read_sql_table('ramal', e)
    instalacao_df = pd.read_sql_table('instalacao', e)
    vistoria_df = pd.read_sql_table('vistoria', e)
    ligacao_df = pd.read_sql_table('ligacao', e)


writer = pd.ExcelWriter(r'C:\Users\gcoel\Documents\Mapa de Progresso brabo\Geral_df.xlsx',
                        engine='openpyxl')


vistoria_df_novo = vistoria_df.rename(columns={'NNota': 'n_nota'})

dfs = {0: ramal_df, 1: ligacao_df, 2: instalacao_df, 3: vistoria_df_novo}

suffix = ('_ramal', '_ligacao', '_instalacao', '_vistoria')

for i in dfs:
    dfs[i] = dfs[i].add_suffix(suffix[i])


dfs[3] = dfs[3].rename(columns={'n_nota_vistoria': 'n_nota'})
dfs[0] = dfs[0].rename(columns={'n_nota_ramal': 'n_nota'})
dfs[1] = dfs[1].rename(columns={'n_nota_ligacao': 'n_nota'})
dfs[2] = dfs[2].rename(columns={'n_nota_instalacao': 'n_nota'})


def agg_df(dflist):
    temp = reduce(lambda left, right: pd.merge(left, right, on='n_nota'), dflist)

    return temp


df_final = agg_df(dfs.values())
df_final = df_final.drop_duplicates(subset=['n_nota']).reset_index(drop=True)

'''
dfs = [ramal_df, ligacao_df, instalacao_df, vistoria_df_novo]
df_final = reduce(lambda left, right: pd.merge(left, right, on='n_nota'), dfs)


dfs[0].to_excel(writer, 'Ramal')
dfs[1].to_excel(writer, 'Ligacao')
dfs[2].to_excel(writer, 'Instalacao')
dfs[3].to_excel(writer, 'Vistoria')
'''
df_final.to_excel(writer, 'Geral')
writer.save()
