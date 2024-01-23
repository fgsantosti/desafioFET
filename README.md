**Desafio de Programação: Conversor FET para Excel**

**Descrição do Problema:**

O Instituto Federal do Piauí (IFPI), campus Corrente, utiliza o FET, um software de agendamento automático de horários. No entanto, o FET, embora eficiente, apresenta uma limitação ao não permitir a exportação direta para o formato Excel. O FET gera os resultados finais em um arquivo .csv, e os usuários precisam converter esses dados para o formato Excel manualmente.

Seu desafio é criar uma ferramenta de conversão automática que permita aos usuários do FET converter facilmente os arquivos .csv para o formato Excel. Essa ferramenta deve ser prática e eficiente, garantindo que todos os dados importantes sejam mantidos durante a conversão.

**Requisitos do Sistema:**

1. **Entrada de Dados:**
   - O sistema deve aceitar como entrada arquivos .csv gerados pelo FET.
   - Os dados nos arquivos .csv incluirão informações sobre disciplinas, salas, professores e horários.

2. **Conversão para Excel:**
   - Desenvolva um algoritmo que converta os dados do formato .csv para o formato Excel (.xlsx).
   - Mantenha a estrutura original dos dados, incluindo as informações sobre disciplinas, salas, professores e horários.

3. **Interface de Usuário (opcional):**
   - Se desejar, crie uma interface de usuário que permita aos usuários selecionar e carregar facilmente os arquivos .csv para conversão.

4. **Validação e Tratamento de Erros:**
   - Implemente mecanismos de validação para garantir que os arquivos .csv estejam formatados corretamente antes da conversão.
   - Forneça mensagens de erro informativas em caso de problemas durante a conversão.

5. **Geração do Arquivo Excel:**
   - A saída do sistema deve ser um arquivo Excel (.xlsx) contendo os dados convertidos.
   - O arquivo Excel deve ser organizado de maneira clara e legível.

**Observações:**
- Os alunos terão acesso aos arquivos .csv gerados pelo FET, bem como ao arquivo PDF final gerado pelo sistema.
- A conversão deve preservar todas as informações essenciais e garantir que a estrutura do horário seja mantida.
- Os alunos são incentivados a usar bibliotecas de manipulação de dados, como Pandas para Python, para facilitar o processamento dos arquivos .csv e a criação do arquivo Excel.
- A eficiência do algoritmo de conversão e a facilidade de uso da interface (se implementada) serão critérios de avaliação importantes.

Este desafio permitirá que os alunos do curso de Análise e Desenvolvimento de Sistemas apliquem seus conhecimentos em manipulação de dados, algoritmos e interfaces de usuário enquanto abordam um problema prático enfrentado pelo IFPI campus Corrente.

**Conversor - Turmas**

```python
import pandas as pd
import os


# Leitura do arquivo CSV
csv_file = 'horarios.csv'
df = pd.read_csv(csv_file)
#df = pd.read_csv(csv_file, encoding='latin1', delimiter=';')

def conversor_excel_turmas(df):
    # Mapeamento dos dias da semana para uma ordem específica
    dias_da_semana_ordem = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira']

    # Mapeamento das horas do dia para uma ordem específica
    horas_dia_ordem = ['7 h', '8 h', '9 h', '10 h', '11 h', '12 h', '13 h', '14 h', '15 h', '16 h', '17 h', '18 h', '19 h', '20 h', '21 h']

    # Organizar os dados por turma, dia da semana e hora da aula
    turmas = df['Students Sets'].unique()

    for turma in turmas:
        # Filtrar o DataFrame pela turma
        turma_df = df[df['Students Sets'] == turma]
        # Organizar os dados por dia da semana e hora da aula
        turma_df = turma_df.sort_values(by=['Day', 'Hour'])

        # Criar um DataFrame vazio para o horário da turma
        horario_turma = pd.DataFrame(index=horas_dia_ordem, columns=dias_da_semana_ordem)

        # Preencher o DataFrame com os dados da turma
        for index, row in turma_df.iterrows():
            horario_turma.loc[row['Hour'], row['Day']] = f"{row['Subject']} - {row['Teachers']}"
        
        # Criar um nome para o arquivo Excel
        excel_file = f'horario_turma_{turma}.xlsx'

        # Salvar o horário da turma no arquivo Excel
        horario_turma.to_excel(excel_file)


def join_excel_turmas():    
    # Pasta contendo os arquivos .xlsx
    pasta_excel = '.'

    # Listar todos os arquivos .xlsx na pasta
    arquivos_excel = [arquivo for arquivo in os.listdir(pasta_excel) if arquivo.endswith('.xlsx')]

    # Criar um DataFrame vazio para consolidar todos os dados
    dados_consolidados = pd.DataFrame()

    # Iterar sobre os arquivos e concatenar os DataFrames
    for arquivo in arquivos_excel:
        caminho_arquivo = os.path.join(pasta_excel, arquivo)
        df_temp = pd.read_excel(caminho_arquivo)
        dados_consolidados = pd.concat([dados_consolidados, df_temp])

    # Salvar os dados consolidados em um único arquivo .xlsx
    excel_output_file = 'horarios_excel.xlsx'
    dados_consolidados.to_excel(excel_output_file, index=False)


conversor_excel_turmas(df)
join_excel_turmas()
```

**Conversor - Professor**

```python
import pandas as pd
import os
# Leitura do arquivo CSV
csv_file = 'horarios.csv'
df = pd.read_csv(csv_file)
#df = pd.read_csv(csv_file, encoding='latin1', delimiter=';')

def conversor_excel_professor(df):
    # Mapeamento dos dias da semana para uma ordem específica
    dias_da_semana_ordem = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira']

    # Mapeamento das horas do dia para uma ordem específica
    horas_dia_ordem = ['7 h', '8 h', '9 h', '10 h', '11 h', '12 h', '13 h', '14 h', '15 h', '16 h', '17 h', '18 h', '19 h', '20 h', '21 h']

    # Organizar os dados por turma, dia da semana e hora da aula
    professores = df['Teachers'].unique()

    for professor in professores:
        # Filtrar o DataFrame pela turma
        professor_df = df[df['Teachers'] == professor]
        # Organizar os dados por dia da semana e hora da aula
        professor_df = professor_df.sort_values(by=['Day', 'Hour'])

        # Criar um DataFrame vazio para o horário da turma
        horario_turma = pd.DataFrame(index=horas_dia_ordem, columns=dias_da_semana_ordem)

        # Preencher o DataFrame com os dados da turma
        for index, row in professor_df.iterrows():
            horario_turma.loc[row['Hour'], row['Day']] = f"{row['Subject']} - {row['Teachers']}"
        
        # Criar um nome para o arquivo Excel
        excel_file = f'horario_professor_{professor}.xlsx'

        # Salvar o horário da turma no arquivo Excel
        horario_turma.to_excel(excel_file)


def join_excel_professor():
    # Pasta contendo os arquivos .xlsx
    pasta_excel = '.'

    # Listar todos os arquivos .xlsx na pasta
    arquivos_excel = [arquivo for arquivo in os.listdir(pasta_excel) if arquivo.endswith('.xlsx')]

    # Criar um DataFrame vazio para consolidar todos os dados
    dados_consolidados = pd.DataFrame()

    # Iterar sobre os arquivos e concatenar os DataFrames
    for arquivo in arquivos_excel:
        caminho_arquivo = os.path.join(pasta_excel, arquivo)
        df_temp = pd.read_excel(caminho_arquivo)
        dados_consolidados = pd.concat([dados_consolidados, df_temp])

    # Salvar os dados consolidados em um único arquivo .xlsx
    excel_output_file = 'horarios_professores_excel.xlsx'
    dados_consolidados.to_excel(excel_output_file, index=False)

```
