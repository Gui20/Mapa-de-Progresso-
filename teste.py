import pandas as pd
from openpyxl import load_workbook
import sqlalchemy
from datetime import datetime
import locale


def conect_sql_server():
    engine = sqlalchemy.create_engine(
        'mssql+pymssql://admin:Ioutility2020!@database-1.cz9ylvwjzzvp.sa-east-1.rds.amazonaws.com/metodo')
    return engine


# Tentando conectar ao servidor
try:
    e = conect_sql_server()
    furodirecional_df = pd.read_sql_table('furodirecional', e)
    # ramal_df = pd.read_sql_table('ramal', e)
    # instalacao_df = pd.read_sql_table('instalacao', e)
    # vistoria_df = pd.read_sql_table('vistoria', e)
    # ligacao_df = pd.read_sql_table('ligacao', e)
except sqlalchemy.except_.NoSuchModuleError:
    print("Deu erro ao conectar no servidor!")
except sqlalchemy.except_.ArgumentError:
    print("Não da pra analisar, possivelmente erro de pontuação da URL!")
except sqlalchemy.except_.OperationalError:
    print("Não foi possível conectar com o banco de dados!")
finally:
    e = conect_sql_server()
    furodirecional_df = pd.read_sql_table('furodirecional', e)
    # ramal_df = pd.read_sql_table('ramal', e)
    # instalacao_df = pd.read_sql_table('instalacao', e)
    # vistoria_df = pd.read_sql_table('vistoria', e)
    # ligacao_df = pd.read_sql_table('ligacao', e)

'''
# Quantidade
qtd_aux = []
for v in furodirecional_df['Extensao_2']:
    try:
        z = float(v)
        if z < 0:
            v = 0
            qtd_aux.append(z)
        else:
            qtd_aux.append(z)
    except:
        qtd_aux.append(v)
        print("Algum erro de digitação na hora de cadastrar! Coluna -> [Quantidade (m)]")

qtd_df = pd.Series(qtd_aux)
'''

iden_furo_aux = []
for v in furodirecional_df['IdentificacaoFormulario']:
    try:
        string = v
        for index, char in enumerate(string):
            if char == '(':
                string_alt = string[index:]
                start = string[:index]
                break
        split1 = ' + '.join(string_alt.split('à')[0].split('+'))
        split2 = ' + '.join(string_alt.split('à')[1].split('+'))
        end = ' à '.join([split1, split2])

        final = start + ' '+end
        iden_furo_aux.append(final)
    except:
        iden_furo_aux.append(v)
iden_furo_df = pd.Series(iden_furo_aux)
