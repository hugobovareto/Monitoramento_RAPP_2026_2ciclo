'''
Objetivo: gerar base de dados em formato de planilha para serem usadas como fonte de dados para a aplicação em Google Apps Script para monitoramento do 2] ciclo dos estudantes em RAPP.
1ª base: Estudantes (base geral de todos componentes de todos estudantes em RAPP, com rendimento e situação)
Colunas:
Matrícula (identificador único do estudante);
Nome;
Componente;
Cod_Inep;
Escola;
Série;
Direc;
Etapa_Ensino;
Rendimento;
Tempo_Prova;
Situação;
Enturmação.


São usadas as seguintes fontes de informação:
- Redash das avaliações de RAPP - Rendimento e Participação;
- Relatórios de Acompanhamento de Turmas e Progressão Parcial. 
- Dados finais do painel de monitoramento de RAPP - 1º ciclo


1) Redash das avaliações de RAPP:
Para informações de rendimento e tempo de prova, por componente para os estudantes.
O identificador único do estudante presente é o email, do qual será tirado a matrícula.
Exclusão de provas com tempo = 0.


2) Relatórios de Acompanhamento de Turmas e Progressão Parcial.
Para saber os estudantes enturmados e alguns que já estão aprovados (mesmo sem fazer a prova Plurall).
Considerar somente com os componentes da Formação Geral Básica da BNCC (sem Ensino Religioso):
- Arte;
- Biologia;
- Ciências;
- Educação Física;
- Filosofia;
- Física;
- Geografia;
- História;
- Língua Espanhola;
- Língua Portuguesa;
- Língua Inglesa;
- Matemática;
- Química;
- Sociologia.

3) Dados finais do painel de monitoramento de RAPP - 1º ciclo:
Para saber os estudantes aprovados no 1º ciclo e que devem ser retirados da base geral do 2º ciclo.

Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.


Tratamentos:
Estudante é considerado enturmado se aparece no Relatório de Acompanhamento de Turmas e Progressão Parcial.
Componente está aprovado se a nota (presente no redash for >= 60).
Estudante com 'SITUAÇÃO FINAL' = APROVADO no Relatório de Acompanhamento de Turmas e Progressão Parcial é considerado aprovado, mesmo que não tenha feito a prova Plurall.
A nota desse estudante será a nota em 'MÉDIA FINAL'.
Estudante com 'MÉDIA FINAL' >= 60 no Relatório de Acompanhamento de Turmas e Progressão Parcial é considerado aprovado, mesmo que não tenha feito a prova Plurall.
A nota desse estudante será a nota em 'MÉDIA FINAL'.
'''


# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re

# Carregar dados dos Relatórios de Acompanhamento de Turmas e Progressão Parcial para saber enturmação e nota dos aprovados
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Enturmados_RAPP_2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []


for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_enturmados = pd.concat(dfs, ignore_index=True)


# Considerar somente os componentes da Formação Geral Básica da BNCC (sem Ensino Religioso):
componentes_bncc = ['Arte',
                     'Biologia',
                     'Ciências',
                     'Educação Física',
                     'Filosofia',
                     'Física',
                     'Geografia',
                     'História',
                     'Língua Espanhola',
                     'Língua Portuguesa',
                     'Língua Inglesa',
                     'Matemática',
                     'Química',
                     'Sociologia']

df_enturmados = df_enturmados[df_enturmados['COMPONENTE CURRICULAR'].isin(componentes_bncc)]

# Criar a coluna 'ETAPA_RESUMIDA' no df_enturmados, com valores 'Ens. Fund. - Anos Finais' e 'Ensino Médio', de acordo com a Série
# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais',
    'TURMA I (6° E 7° ANOS)': 'Ens. Fund. - Anos Finais',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '1º Período (1ª Série)': 'Ensino Médio',
    '2º Período (2ª Série)': 'Ensino Médio',
    '3º Período (3ª Série)': 'Ensino Médio',
    '1° PERÍODO': 'Ensino Médio',
    '2° PERÍODO': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    '4° PERÍODO': 'Ensino Médio',
    '1º MÓDULO': 'Ensino Médio',
    '1º MODULO': 'Ensino Médio',
    '2° MÓDULO': 'Ensino Médio',
    '2º MODULO': 'Ensino Médio',
    '3º MÓDULO': 'Ensino Médio',
    '3º MODULO': 'Ensino Médio',
    '4º MÓDULO': 'Ensino Médio',
    '4º MODULO': 'Ensino Médio',
    'BLOCO A': 'Ensino Médio',
    'BLOCO B': 'Ensino Médio',
    'BLOCO C': 'Ensino Médio',
    'BLOCO D': 'Ensino Médio',
    '5º PERÍODO (8° E 9° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '4º PERÍODO (6° E 7° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '1º SEMESTRE': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    '3º SEMESTRE': 'Ensino Médio',
    '4º SEMESTRE': 'Ensino Médio',
    'UNICA': 'Ens. Fund. - Anos Finais'
}

df_enturmados['ETAPA_RESUMIDA'] = df_enturmados['SÉRIE'].map(mapeamento_etapa)

# Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'
# Garante que MÉDIA FINAL é numérica
df_enturmados['MÉDIA FINAL'] = pd.to_numeric(
    df_enturmados['MÉDIA FINAL'].str.replace(',', '.', regex=False),
    errors='coerce'
)

# Cria a prioridade (1 = APROVADO, 0 = demais)
df_enturmados['_prioridade'] = (
    df_enturmados['SITUAÇÃO FINAL']
    .eq('APROVADO')
    .astype(int)
)

# Ordena pelos critérios
df_enturmados = (
    df_enturmados
    .sort_values(
        by=['_prioridade', 'MÉDIA FINAL'],
        ascending=[False, False]
    )
    .drop_duplicates(
        subset=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
        keep='first'
    )
    .drop(columns='_prioridade')
    .reset_index(drop=True)
)


# Converter MATRÍCULA para string para garantir que a junção funcione corretamente
df_enturmados['MATRÍCULA'] = (
    df_enturmados['MATRÍCULA']
    .astype(str)
    .str.strip()
)


# Trocar nome da coluna 'ESTUDANTE' para 'NOME'
df_enturmados = df_enturmados.rename(columns={'ESTUDANTE': 'NOME'})

# Manter as colunas de interesse
df_enturmados = df_enturmados[['MATRÍCULA', 'NOME', 'COMPONENTE CURRICULAR', 'INEP ESCOLA', 'ESCOLA', 'SÉRIE', 'DIREC', 'ETAPA_RESUMIDA', 'SITUAÇÃO FINAL', 'MÉDIA FINAL']]


# Trocar o nome do dataframe (para aproveitar o código quando era feito merge entre df_rapp e df_enturmados)
df_final = df_enturmados


# Situação para os novos registros
df_final['Situação'] = 'Não Avaliado'

# Coluna de 'Enturmação'. Todos são 'Sim' porque todos vieram do df_enturmados
df_final['Enturmação'] = 'Sim'


# Estudante com 'SITUAÇÃO FINAL' = APROVADO é considerado aprovado, mesmo que não tenha feito a prova Plurall.
# A nota desse estudante será a nota que está em 'MÉDIA FINAL'.

# Máscara dos aprovados
mask = df_final['SITUAÇÃO FINAL'].eq('APROVADO')


# Converter 'MÉDIA FINAL' para numérico
df_final['MÉDIA FINAL'] = pd.to_numeric(
    df_final['MÉDIA FINAL'].astype(str).str.replace(',', '.', regex=False),
    errors='coerce'
)

# Multiplicar por 10 para ficar na mesma escala de 'rendimento (%)'
df_final['MÉDIA FINAL'] = df_final['MÉDIA FINAL'] * 10


# Atualizar nota e situação do df_final para os estudantes aprovados no df_enturmados
df_final.loc[mask, 'rendimento (%)'] = df_final.loc[mask, 'MÉDIA FINAL']
df_final.loc[mask, 'Situação'] = 'Aprovado'


# Estudantes com 'MÉDIA FINAL' >= 60 são considerados aprovados, mesmo que não tenham feito a prova Plurall.
# A nota desse estudante é a nota que está em 'MÉDIA FINAL'.

# Identifica estudantes com MÉDIA FINAL maior ou igual a 60
mask = df_final['MÉDIA FINAL'].ge(60)

# Atualiza rendimento e situação
df_final.loc[mask, 'rendimento (%)'] = df_final.loc[mask, 'MÉDIA FINAL']
df_final.loc[mask, 'Situação'] = 'Aprovado'


# Excluir as colunas 'SITUAÇÃO FINAL' e 'MÉDIA FINAL' do df_final
df_final = df_final.drop(columns=['SITUAÇÃO FINAL', 'MÉDIA FINAL'])


# Carregar dados do Redash (para ter nota e tempo de prova do estudante)
df_redash = pd.read_csv(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026_2ciclo\SEEC-RN_-_Rendimento_e_participação_dos_alunos_p_provas_-_RAPP_-_Avaliações_em_andamento_2026_10_27.csv")

# Manter somente os valores que possuem '2º Ciclo' e '2º ciclo' na coluna 'name'
df_redash = df_redash[df_redash['name'].str.contains('2º ciclo', case=False, na=False)]

# Código retirado, pois estudantes que realizam a prova antes da possibilidade de agendamento aparecem sem tempo de prova, mas as notas devem ser mantidas
# Excluir valores que o tempo de prova foi 0 (zero) ou nulo
# df_redash = df_redash[
#     df_redash['tempo de prova'].notna() &
#     (df_redash['tempo de prova'].str.strip() != '') &
#     (df_redash['tempo de prova'] != '00:00:00')]

# Trocar valor 'Língua Ingles' para 'Língua Inglesa' na coluna 'prova'
df_redash.loc[df_redash['prova'].eq('Língua Ingles'), 'prova'] = 'Língua Inglesa'


# Criar a coluna 'MATRÍCULA' no df_redash, extraindo a matrícula do email do estudante
df_redash['MATRÍCULA'] = df_redash['email_aluno'].str.split('@').str[0]


# Matrículas no df_redash que não estão no df_final
df_matriculas_nao_encontradas = df_redash[
    ~df_redash['MATRÍCULA'].isin(df_final['MATRÍCULA'])]


# Pegar as informações do df_matriculas_nao_encontradas para trocar as matrículas no df_final nos casos que estiverem presentes no df_matriculas_nao_encontradas
# Criar mapa somente com as informações necessárias
df_mapa = df_matriculas_nao_encontradas[
    ['MATRÍCULA', 'aluno', 'prova']
].rename(columns={
    'aluno': 'NOME',
    'prova': 'COMPONENTE CURRICULAR',
    'MATRÍCULA': 'MATRÍCULA_NOVA'
})

# Fazer o match usando NOME + COMPONENTE CURRICULAR
df_final = df_final.merge(
    df_mapa,
    on=['NOME', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Trocar a matrícula SOMENTE quando houve match
df_final['MATRÍCULA'] = df_final['MATRÍCULA_NOVA'].combine_first(
    df_final['MATRÍCULA']
)

# Remover coluna auxiliar
df_final = df_final.drop(columns='MATRÍCULA_NOVA')


# De acordo com a matrícula (extraída do email) e componente, juntar as informações de nota e tempo de prova do Redash com a base geral de estudantes em RAPP
# Selecionar as colunas de interesse do df_redash
df_redash_merge = df_redash[['MATRÍCULA', 'prova', 'rendimento (%)', 'tempo de prova', 'quantidade_reaberturas']]

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_final["MATRÍCULA"] = (
    df_final["MATRÍCULA"]
    .astype("string")
    .str.strip()
)

df_redash_merge["MATRÍCULA"] = (
    df_redash_merge["MATRÍCULA"]
    .astype("string")
    .str.strip()
)

# Merge trazendo o rendimento do Redash com um nome temporário
df_merged = df_final.merge(
    df_redash_merge[['MATRÍCULA', 'prova', 'rendimento (%)', 'tempo de prova', 'quantidade_reaberturas']],
    left_on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    right_on=['MATRÍCULA', 'prova'],
    how='left',
    suffixes=('', '_redash')
)

# Substitui o rendimento do df_final pelo rendimento do Redash, somente quando houver valor no Redash
df_merged['rendimento (%)'] = df_merged['rendimento (%)_redash'].fillna(
    df_merged['rendimento (%)']
)

# Remove as colunas auxiliares
df_merged = df_merged.drop(
    columns=['prova', 'rendimento (%)_redash']
)


# Adiciona valores em Situação de acordo com o rendimento, exceto para os casos já aprovados vindos da Enturmação
condicoes = [
    df_merged['Situação'].eq('Aprovado'),
    df_merged['rendimento (%)'].isna(),
    df_merged['rendimento (%)'].ge(60)
]

valores = [
    'Aprovado',
    'Não Avaliado',
    'Aprovado'
]

df_merged['Situação'] = np.select(
    condicoes,
    valores,
    default='Não Aprovado'
)


# Reorganizar a ordem das colunas
df_merged = df_merged[
    ['MATRÍCULA',
     'NOME', 
     'COMPONENTE CURRICULAR',
     'INEP ESCOLA',
     'ESCOLA',
     'SÉRIE',
     'DIREC',
     'ETAPA_RESUMIDA',
     'rendimento (%)',
     'tempo de prova',
     'Situação',
     'Enturmação',
     'quantidade_reaberturas']
]

# Na coluna 'quantidade_reaberturas', substituir valores nulos por 0
df_merged['quantidade_reaberturas'] = df_merged['quantidade_reaberturas'].fillna(0).astype(int)



# Carregar dados dos estudantes do 1º ciclo que foram aprovados, para retirar da base geral do 2º ciclo
df_1ciclo = pd.read_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026_2ciclo\1ciclo_Relatorio_Monitoramento_componentes.xlsx")


# Retirar da base geral do 2º ciclo os estudantes que foram aprovados no 1º ciclo
# Filtra no df_1ciclo apenas os registros aprovados
df_aprovados = df_1ciclo.loc[
    df_1ciclo['SITUAÇÃO'].eq('Aprovado'),
    ['MATRÍCULA', 'COMPONENTE']
]


# Normaliza as chaves do df_merged
merged_norm = df_merged[['MATRÍCULA', 'COMPONENTE CURRICULAR']].copy()

merged_norm['MATRÍCULA'] = (
    merged_norm['MATRÍCULA']
    .astype('string')
    .str.strip()
    .str.replace(r'\.0$', '', regex=True)
)

merged_norm['COMPONENTE CURRICULAR'] = (
    merged_norm['COMPONENTE CURRICULAR']
    .astype('string')
    .str.strip()
)

# Normaliza as chaves do df_aprovados
aprovados_norm = df_aprovados[['MATRÍCULA', 'COMPONENTE']].copy()

aprovados_norm['MATRÍCULA'] = (
    aprovados_norm['MATRÍCULA']
    .astype('string')
    .str.strip()
    .str.replace(r'\.0$', '', regex=True)
)

aprovados_norm['COMPONENTE'] = (
    aprovados_norm['COMPONENTE']
    .astype('string')
    .str.strip()
)

# Identifica as combinações que devem ser excluídas
chaves_merged = pd.MultiIndex.from_frame(merged_norm)
chaves_aprovados = pd.MultiIndex.from_frame(aprovados_norm)

mask_excluir = chaves_merged.isin(chaves_aprovados)

# Exclui as linhas correspondentes
df_merged = df_merged.loc[~mask_excluir].copy()


# Exportar a base final em Excel para usar na aplicação em Google Apps Script
df_merged.to_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026_2ciclo\20261027_Monitoramento_RAPP.xlsx", index=False)








































##################################################################################
'''
Gerar uma planilha para alimentar a aplicação em Google Apps Script para monitoramento dos estudantes em RAPP, especificamente na página sobre o Agendamento das Avaliações.
Dados granulados por Turma, de acordo com Redash.

Colunas:
DIREC;
ESCOLA;
ID ESCOLA;
TURMA;
ID TURMA;
AGENDOU;
COMPONENTE;
SERIE;
ETAPA_RESUMIDA;
AGENDAMENTO


O 'AGENDAMENTO' apresenta a data de abertura, horário de abertura, data de fechamento e horário de fechamento.
'''

# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re


# Carregar os dados de agendamento (vindos do Redash)
df_agendamento = pd.read_csv(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026_2ciclo\[SEEC-RN]_Agendamento_de_Avaliações_por_turmas_2026_10_27.csv")


# Manter somente os valores que possuem '2º Ciclo' e '2º ciclo' na coluna 'Nome da Avaliação'
df_agendamento = df_agendamento[df_agendamento['Nome da Avaliação'].str.contains('2º ciclo', case=False, na=False)]


# Separar a informação do "Nome da Avaliação" para ter o Componente e a Série
# Excluir tudo a direita do último hífen (o texto é '2º ciclo')
df_agendamento['Nome da Avaliação'] = (
    df_agendamento['Nome da Avaliação']
    .str.rsplit('-', n=1)
    .str[0]
    .str.strip()
)

# Excluir tudo a esquerda do 2º hífen (o texto é 'RAPP - RN')
df_agendamento['Nome da Avaliação'] = (
    df_agendamento['Nome da Avaliação']
    .str.split('-', n=2)
    .str[2]
    .str.strip()
)

# Coluna de Componente com o que está a esquerda do hífen.
# Coluna de Série com o que está a direita do hífen
df_agendamento[['COMPONENTE', 'SERIE']] = (
    df_agendamento['Nome da Avaliação']
    .str.split('-', n=1, expand=True)
)

df_agendamento['COMPONENTE'] = df_agendamento['COMPONENTE'].str.strip()
df_agendamento['SERIE'] = df_agendamento['SERIE'].str.strip()


# Criar a coluna 'AGENDAMENTO' com a data e horário de abertura e fechamento
df_agendamento['AGENDAMENTO'] = (
    df_agendamento['Data de Abertura'] + ' ' + df_agendamento['Horário de Abertura'] + ' - ' + df_agendamento['Data de Fechamento'] + ' ' + df_agendamento['Horário de Fechamento']
)


# Excluir colunas que não serão usadas
df_agendamento = df_agendamento.drop(columns=[
    'Id Avaliação',
    'Nome da Avaliação',
    'CNPJ Escola',
    'Bloco',
    'Data de marcação do agendamento',
    'Data de Abertura',
    'Horário de Abertura',
    'Data de Fechamento',
    'Horário de Fechamento'])


# Trocar os nomes das colunas
df_agendamento = df_agendamento.rename(columns={
    'Id escola': 'ID ESCOLA',
    'Nome escola': 'ESCOLA',
    'Grupo': 'DIREC', 
    'Id Turma': 'ID TURMA',
    'Nome Turma': 'TURMA',
    'Agendou': 'AGENDOU'
})

# Criar coluna 'ETAPA_RESUMIDA' de acordo com os valores de Série
mapeamento_etapa = {
    '1ª série': 'Ensino Médio',
    '2ª série': 'Ensino Médio',
    '3ª série': 'Ensino Médio',
    '6º ano': 'Ens. Fund. - Anos Finais',
    '7º ano': 'Ens. Fund. - Anos Finais',
    '8º ano': 'Ens. Fund. - Anos Finais',
    '9º ano': 'Ens. Fund. - Anos Finais'
    }

df_agendamento['ETAPA_RESUMIDA'] = df_agendamento['SERIE'].map(mapeamento_etapa)


# Trocar ordem das colunas
df_agendamento = df_agendamento[
    ['DIREC', 'ESCOLA', 'ID ESCOLA', 'TURMA', 'ID TURMA', 'AGENDOU', 'COMPONENTE', 'SERIE', 'ETAPA_RESUMIDA','AGENDAMENTO']
]



############ LÓGICA DE CONTABILIZAÇÃO DE TURMAS E ESCOLAS AGENDADAS ############
'''
O dataframe é dividido em componentes e turmas por linhas.
Uma turma tem vários componentes diferentes, ou seja, várias linhas.
A turma para ser considerada 'Agendada' deve ter todos os componentes agendados.
Para ser considerada 'Não Agendada' basta ter 1 componente não agendado.
Escola 100% Agendada precisa de todas as turmas agendadas (em todos os componentes).
Escola 0% Agendada precisa de todas as turmas não agendadas (ou seja, cada turma com pelo menos 1 componente não agendado).
'''
# Nº de Turmas Agendadas
# Considerar uma turma agendada somente quando todos os registros (linhas) estiverem como AGENDOU = True
turmas_agendadas = (
    df_agendamento
    .groupby(['ID ESCOLA', 'ID TURMA'])['AGENDOU']
    .all()
)

# Quantidade de turmas agendadas
qtd_turmas_agendadas = turmas_agendadas.sum()

# Total de turmas
qtd_turmas_total = turmas_agendadas.shape[0]

# Percentual
perc_turmas_agendadas = qtd_turmas_agendadas / qtd_turmas_total * 100

print(f'Turmas agendadas: {qtd_turmas_agendadas}')
print(f'Total de turmas: {qtd_turmas_total}')
print(f'Percentual: {perc_turmas_agendadas:.2f}%')  

# Dataframe dessas turmas
# Recuperar os pares ID ESCOLA + ID TURMA das turmas agendadas
turmas_agendadas_ids = turmas_agendadas[turmas_agendadas].index

# Transformar o MultiIndex em DataFrame
df_ids_agendadas = turmas_agendadas_ids.to_frame(index=False)

# Fazer o filtro usando ESCOLA + TURMA
df_turmas_agendadas = df_agendamento.merge(
    df_ids_agendadas,
    on=['ID ESCOLA', 'ID TURMA'],
    how='inner'
)


# Situação de cada turma
# Cria uma tabela com a situação de cada turma
df_turmas = (
    df_agendamento
    .groupby(['ID ESCOLA', 'ID TURMA'])['AGENDOU']
    .all()
    .reset_index()
)

df_turmas = df_turmas.rename(columns={
    'AGENDOU': 'TURMA_AGENDADA'
})


# Escolas com 100% de Agendamento
escolas_100 = (
    df_turmas
    .groupby('ID ESCOLA')['TURMA_AGENDADA']
    .all()
)

# Quantidade de escolas com 100% de agendamento
qtd_escolas_100 = escolas_100.sum()

# Total de escolas
qtd_escolas_total = escolas_100.shape[0]

# Percentual
perc_escolas_100 = qtd_escolas_100 / qtd_escolas_total * 100

print(f'Escolas com 100% de agendamento: {qtd_escolas_100}')
print(f'Total de escolas: {qtd_escolas_total}')
print(f'Percentual: {perc_escolas_100:.2f}%')


# Dataframe dessas escolas
escolas_100_ids = escolas_100[escolas_100].index
df_escolas_100 = df_agendamento[
    df_agendamento['ID ESCOLA'].isin(escolas_100_ids)
]


# Escolas com 0% de Agendamento
escolas_0 = (
    df_turmas
    .groupby('ID ESCOLA')['TURMA_AGENDADA']
    .apply(lambda x: (~x).all())
)

# Quantidade de escolas com 0% de agendamento
qtd_escolas_0 = escolas_0.sum()

# Total de escolas
qtd_escolas_total = escolas_0.shape[0]

# Percentual
perc_escolas_0 = qtd_escolas_0 / qtd_escolas_total * 100

print(f'Escolas com 0% de agendamento: {qtd_escolas_0}')
print(f'Total de escolas: {qtd_escolas_total}')
print(f'Percentual: {perc_escolas_0:.2f}%')

# Dataframe dessas escolas
escolas_0_ids = escolas_0[escolas_0].index
df_escolas_0 = df_agendamento[
    df_agendamento['ID ESCOLA'].isin(escolas_0_ids)
]


# Salvar em Excel o dataframe
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026_2ciclo\20261027_Agendamento_RAPP.xlsx") as writer:
    df_agendamento.to_excel(writer, sheet_name='Agendamento', index=False)
    df_turmas_agendadas.to_excel(writer, sheet_name='Turmas Agendadas', index=False)
    df_escolas_100.to_excel(writer, sheet_name='Escolas 100% Agendadas', index=False)
    df_escolas_0.to_excel(writer, sheet_name='Escolas 0% Agendadas', index=False)





