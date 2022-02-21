import pandas as pd
from openpyxl import load_workbook
import sqlalchemy
from datetime import datetime
import locale
from functools import reduce
import time

inicio = time.time()

locale.setlocale(locale.LC_ALL, 'pt_pt.UTF-8')


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

# Criando um book com o formato que deve ser o Mapa de Progresso
book = load_workbook(r'C:\Users\gcoel\PycharmProjects\AcessandoSQL\Mapa de Progresso - Modelo - Oficial.xlsx')

# Selecionando local e arquivo a ser criado com o formato do book
writer = pd.ExcelWriter(r'C:\Users\gcoel\Documents\Mapa de Progresso brabo\Mapa de Progresso - Oficial - MERGE.xlsx',
                        engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

'''
-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- Mapa de Progresso Rede -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
 Concatenando as colunas a partir da tabela FURODIRECIONAL 

 '''

# Fase
fase_mr_aux = []
for v in furodirecional_df['Condominio']:
    z = v.upper()
    if "RIC" in z:
        fase_mr_aux.append(z[:8])
    elif "LL4" in z:
        fase_mr_aux.append(z[:8])
    else:
        fase_mr_aux.append("")

fase_mr_df = pd.Series(fase_mr_aux)

# RIC / LL4
aux = []
for v in furodirecional_df['Pressao_2']:
    aux.append(v.split(" ")[0])

pressao_df = pd.Series(aux)

# Material
mat_aux = []
for i in range(len(furodirecional_df['Desc1_3_1'])):
    v = 'TUBO PE'
    mat_aux.append(v)
mat_df = pd.Series(mat_aux)

# Diâmetro
diam_df = pd.to_numeric(furodirecional_df['DN_2'], downcast='integer')

# Furo
furo_aux = []
for i in range(len(mat_aux)):
    furo_aux.append(1)
furo_df = pd.Series(furo_aux)

# Estaca Inicial
es_inicial_df = pd.to_numeric(furodirecional_df['EstacaInicial1_2'], downcast='integer')

# Estaca Inicial (Complemento)
es_inicial_complemento_df = pd.to_numeric(furodirecional_df['EstacaInicial2_2'])

# Estaca Final
es_final_aux = []
for x in furodirecional_df['EstacaFinal1_2']:
    try:
        x = float(x)
        es_final_aux.append(x)
    except ValueError:
        es_final_aux.append(x)

es_final_df = pd.Series(es_final_aux)

# Estaca Final (Complemento)
es_final_complemento_df = pd.to_numeric(furodirecional_df['EstacaFinal2_2'])

# Nº. Relatório Furo Direcional
n_rel_df = pd.to_numeric(furodirecional_df['REF'])

# RDO
rdo_df = pd.to_numeric(furodirecional_df['NRDO'])

# Quantidade
qtd_aux = []
for v in furodirecional_df['Extensao_2']:
    if v == "":
        qtd_aux.append(v)
    else:
        try:
            z = float(v)
            if z < 0:
                z = 0
                qtd_aux.append(z)
            else:
                qtd_aux.append(z)
        except:
            qtd_aux.append(v)
            print("Algum erro de digitação na hora de cadastrar! Coluna -> [Quantidade (m)]")

qtd_df = pd.Series(qtd_aux)

# Furo 80%
furo_aux = []

for v in furodirecional_df['Data']:
    try:
        furo_aux.append(datetime.strptime(v, '%d/%m/%Y').strftime('%d/%b/%Y'))
    except:
        furo_aux.append(v)
        print(
            "Os valores não estão sendo salvos na base de dados com um padrão! -> Coluna U - Furo 80% - Mapa de Progresso Rede")
furo_df = pd.Series(furo_aux)

# Mês
l_aux = []
for v in furodirecional_df['Data'].tolist():
    l_aux.append(int(v[3:5]))

data_df = pd.Series(l_aux)

# Semana
d_aux = []
d1 = datetime.strptime('28/08/2020', '%d/%m/%Y')

for data in furodirecional_df['Data']:
    # Data final
    d2 = datetime.strptime(data, '%d/%m/%Y')
    diff = abs((d2 - d1).days)
    d_aux.append(diff // 7 + 1)
semana_df = pd.Series(d_aux)

# Série vazia para os campos de preenchimento manual
vazio_aux = []
for i in range(len(mat_aux)):
    vazio_aux.append("")
vazio_df = pd.Series(vazio_aux)

# Identificação Furo
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
        final = start + ' ' + end
        iden_furo_aux.append(final)
    except:
        iden_furo_aux.append(v)
iden_furo_df = pd.Series(iden_furo_aux)

# Criação de uma lista com todos os dataframes a serem concatenados
dataframes = [furodirecional_df['Municipio'], furodirecional_df['tipo'], furodirecional_df['Rua_2'],
              furodirecional_df['Projeto'], furodirecional_df['TU'], furodirecional_df['PEP'],
              fase_mr_df, pressao_df, furodirecional_df['Metodo_4'], mat_df, diam_df, furo_df,
              es_inicial_df, es_inicial_complemento_df, es_final_df, es_final_complemento_df, n_rel_df, rdo_df,
              qtd_df, vazio_df, furo_df, vazio_df, vazio_df, data_df,
              semana_df, iden_furo_df]

# Função que concatena todos os dataframes
mapa_progresso_rede_df = pd.concat(dataframes, axis=1, join="outer")

# Escrevendo no Excel pré-formatado
mapa_progresso_rede_df.to_excel(writer, 'Mapa de Progresso Rede', startrow=7, startcol=0, header=False, index=False)

'''
-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-  Mapa de Progresso Ramal + Ligação  -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 Concatenando as colunas a partir da tabela RAMAL, INSTALAÇÃO, VISTORIA E LIGAÇÃO
'''

# Renomeia a coluna n_nota pra conseguir aplicar o merge
vistoria_df_novo = vistoria_df.rename(columns={'NNota': 'n_nota'})

# Aplicando um sufixo em todos os dataframes para fácil identificação
dfs = {0: ramal_df, 1: ligacao_df, 2: instalacao_df, 3: vistoria_df_novo}
suffix = ('_ramal', '_ligacao', '_instalacao', '_vistoria')
for i in dfs:
    dfs[i] = dfs[i].add_suffix(suffix[i])

# renomeando as colunas n_nota antes de aplicar o merge
dfs[3] = dfs[3].rename(columns={'n_nota_vistoria': 'n_nota'})
dfs[0] = dfs[0].rename(columns={'n_nota_ramal': 'n_nota'})
dfs[1] = dfs[1].rename(columns={'n_nota_ligacao': 'n_nota'})
dfs[2] = dfs[2].rename(columns={'n_nota_instalacao': 'n_nota'})


# mesclando os 4 dataframes
def agg_df(dflist):
    temp = reduce(lambda left, right: pd.merge(left, right, on='n_nota'), dflist)
    return temp


df_final = agg_df(dfs.values())

# Removendo as linhas duplicadas
df_final = df_final.drop_duplicates(subset=['n_nota']).reset_index(drop=True)

# Nota
nota_df = pd.to_numeric(df_final['n_nota'])

# Número
num_df = pd.to_numeric(df_final['numero_endereco_ramal'])

# RIC / LL4
ric_aux = []
for v in df_final['local_atividade_ramal']:
    if "RIC" in v:
        v = "RIC"
        ric_aux.append(v)
    elif "LL4" in v:
        v = "LL4"
        ric_aux.append(v)
    else:
        # posso alterar pra aparecer exatamente o que tem nos espaços em branco
        ric_aux.append(" ")

ric_df = pd.Series(ric_aux)

# Condomínio / Fase
fase_aux = []
cond_aux = []
for v in df_final['local_atividade_ramal']:
    x = v.upper()
    if "RIC" in x:
        fase_aux.append(x[:8])
        cond_aux.append("")
    elif "LL4" in x:
        fase_aux.append(x[:8])
        cond_aux.append("")
    else:
        fase_aux.append("")
        cond_aux.append(x)

fase_df = pd.Series(fase_aux)
cond_df = pd.Series(cond_aux)

# Método
tat_aux = []
for i in df_final['n_relatorio_ramal'].index:
    tat_aux.append('Tatuzinho')

tat_df = pd.Series(tat_aux)

# Material
mat_aux = []
for v in df_final['rede_distribuicao_material_ramal']:
    if "PE" in v:
        x = "Tubo PE"
        mat_aux.append(x)
    else:
        x = "Tubo PE"
        mat_aux.append(x)

mat_df = pd.Series(mat_aux)

# Furo 80%
furo_r_aux = []

for v in df_final['data_info_gerais_ramal']:
    try:
        furo_r_aux.append(datetime.strptime(v, '%d/%m/%Y').strftime('%d/%b/%Y'))
    except:
        furo_r_aux.append(v)
        print("Os valores não estão sendo salvos na base de dados com um padrão! "
              "Coluna O - Furo 80% - Mapa de Progresso Ramal + Ligação")
furo_r_df = pd.Series(furo_r_aux)

# mês
s = []
for v in df_final['data_info_gerais_ramal']:
    s.append(int(v[3:5] + v[6:]))
s_df = pd.Series(s)

# Semana
d_ramal_aux = []
# Data inicial
d1_ramal = datetime.strptime('28/08/2020', '%d/%m/%Y')

for data_ramal in df_final['data_info_gerais_ramal']:
    # Data final
    d2_ramal = datetime.strptime(data_ramal, '%d/%m/%Y')
    diff = abs((d2_ramal - d1_ramal).days)
    d_ramal_aux.append(diff // 7 + 1)

semana_ramal_df = pd.Series(d_ramal_aux)

# Quantidade Ligação (un)
qtd_ligacao_aux = []
for v in df_final['data_info_gerais_ligacao']:
    if v == "":
        v = 0
        qtd_ligacao_aux.append(int(v))
    else:
        v = 1
        qtd_ligacao_aux.append(int(v))

qtd_r_df = pd.Series(qtd_ligacao_aux)

# Vistoria Data
vist_data_aux = []

for v in df_final['Data_vistoria']:
    if v == "":
        vist_data_aux.append(v)
    else:
        try:
            vist_data_aux.append(datetime.strptime(v, '%d/%m/%Y').strftime('%d/%b/%Y'))
        except:
            vist_data_aux.append(v)
            print("Os valores não estão sendo salvos na base de dados com um padrão! "
                  "Coluna W - Vistoria (15%)- Mapa de Progresso Ramal + Ligação")
vist_data_df = pd.Series(vist_data_aux)

# Interna Data
int_data_aux = []

for v in df_final['data_info_gerais_instalacao']:
    if v == "":
        int_data_aux.append(v)
    else:
        try:
            int_data_aux.append(datetime.strptime(v, '%d/%m/%Y').strftime('%d/%b/%Y'))
        except:
            int_data_aux.append(v)
            print("Os valores não estão sendo salvos na base de dados com um padrão! "
                  "Coluna X - Interna (45%) - Mapa de Progresso Ramal + Ligação")
int_data_df = pd.Series(int_data_aux)

# Ligação Data
lig_data_aux = []

for v in df_final['data_info_gerais_ligacao']:
    if v == "":
        lig_data_aux.append(v)
    else:
        try:
            lig_data_aux.append(datetime.strptime(v, '%d/%m/%Y').strftime('%d/%b/%Y'))
        except:
            lig_data_aux.append(v)
            print("Os valores não estão sendo salvos na base de dados com um padrão! "
                  "Coluna Y - Ligação (20%) - Mapa de Progresso Ramal + Ligação")
lig_data_df = pd.Series(lig_data_aux)

# Mês ligação
s_lig_aux = []
for v in df_final['data_info_gerais_ligacao']:
    if v == '':
        s_lig_aux.append(v)
    else:
        s_lig_aux.append(int(v[3:5] + v[6:]))
s_lig_df = pd.Series(s_lig_aux)

# Semana Ligação
d_lig_aux = []
# Data inicial
d1_lig = datetime.strptime('28/08/2020', '%d/%m/%Y')

for data_lig in df_final['data_info_gerais_ligacao']:
    if data_lig == "":
        d_lig_aux.append("")
    else:
        # Data final
        d2_lig = datetime.strptime(data_lig, '%d/%m/%Y')
        diff = abs((d2_lig - d1_lig).days)
        d_lig_aux.append(diff // 7 + 1)

semana_lig_df = pd.Series(d_lig_aux)

# Vistoria
vist_aux = []
for value in df_final['Data_vistoria']:
    if value == "":
        vist_aux.append("")
    else:
        vist_aux.append("VL INT")
vist_df = pd.Series(vist_aux)

# Interna
inter_aux = []
for value in df_final['data_info_gerais_instalacao']:
    if value == "":
        inter_aux.append("")
    else:
        inter_aux.append("VL INT")
inter_df = pd.Series(inter_aux)

# Ramal
r_aux = []
for value in df_final['data_info_gerais_ramal']:
    if value == "":
        r_aux.append("")
    else:
        r_aux.append("VL INT")
r_df = pd.Series(r_aux)

# Ligação
lig_aux = []
for value in df_final['data_info_gerais_ligacao']:
    if value == "":
        lig_aux.append("")
    else:
        lig_aux.append("VL INT")
lig_df = pd.Series(lig_aux)

# Criação de lista com todos os dataframes a serem concatenados para Mapa de Ramal e Ligação
dataframes_2 = [nota_df, df_final['endereco_cliente_ramal'], num_df, df_final['cidade_ramal'], cond_df, ric_df,
                df_final['tipo_ramal_ramal'], fase_df, df_final['tipo_pacote_venda_ligacao'], tat_df, mat_df, vazio_df,
                df_final['n_relatorio_ramal'], df_final['servicos_qtd_ramal'], furo_r_df,
                furo_r_df, vazio_df, s_df, semana_ramal_df, vazio_df, df_final['n_relatorio_ligacao'],
                qtd_r_df, vist_data_df, int_data_df, lig_data_df,
                vazio_df, s_lig_df, semana_lig_df, vazio_df, vist_df, inter_df, r_df, lig_df]

# Função que concatena todos os dataframes
mapa_progresso_ramal_ligacao_df = pd.concat(dataframes_2, axis=1, join="outer")

# Escrevendo no Excel pré-formatado
mapa_progresso_ramal_ligacao_df.to_excel(writer, 'Mapa de Progresso Ramal+Ligação', startrow=7, startcol=0,
                                         header=False, index=False)

writer.save()

fim = time.time()

print("O tempo de execução é: {} s ".format(fim - inicio))
