import pandas as pd
import io
import subprocess

# Carregar a planilha enviada
input_file = 'dados_planilha.xlsx'

# Variáveis fixas para as colunas adicionais
dominio = "@alunosenai.mt"
office = "SENAI - Nova Mutum/MT"
criador = "Criado por Jeferson Silva"
destino = "OU=QUA.415.089 ASSISTENTE DE RECURSOS HUMANOS COM INFORMÁTICA,OU=CURSOS,OU=SENAINMT,OU=SENAI,OU=SFIEMT-EDU,DC=SESISENAIMT,DC=EDU"

# Ler a planilha
df = pd.read_excel(input_file)

# Ajustar os nomes das colunas conforme a estrutura da planilha
df.columns = ['TURMA', 'RA', 'ALUNO', 'CPF', 'USUÁRIO', 'SENHA', 'E-MAIL / OFFICE 365']

# Remover a linha que corresponde ao cabeçalho antigo, caso esteja presente nos dados
df = df[~df['RA'].astype(str).str.contains('RA', case=False)]
df = df[~df['ALUNO'].astype(str).str.contains('ALUNO', case=False)]

# Remover outras linhas em branco ou com dados faltantes (se necessário)
df = df.dropna(subset=['RA', 'ALUNO'])

# Função para gerar a senha a partir do RA
def gerar_senha(ra):
    return ra[2:]  # Mantém apenas os últimos 4 dígitos do RA

# Função para separar o primeiro nome e o restante
def separar_nome(nome_completo):
    partes = nome_completo.split(' ', 1)  # Limitar a divisão ao primeiro espaço
    primeiro_nome = partes[0].upper()  # Primeiro nome em maiúsculas
    sobrenome = partes[1].upper() if len(partes) > 1 else ''  # Sobrenome em maiúsculas
    return primeiro_nome, sobrenome

# Aplicar a função para gerar a senha e adicionar colunas extras
df['Senha'] = df['RA'].apply(lambda x: gerar_senha(str(x)))
df['Primeiro Nome'], df['Sobrenome'] = zip(*df['ALUNO'].apply(separar_nome))

# Criar o e-mail a partir do RA
df['E-mail'] = df['RA'].apply(lambda x: str(x).zfill(8) + dominio)

# Criar a coluna personalizada conforme solicitado
df['Descrição Completa'] = df['ALUNO'].str.upper() + ' - ' + office

# Montar a estrutura de saída conforme solicitado
df_final = pd.DataFrame()
df_final['Nome'] = df['ALUNO'].str.upper()  # Nome em maiúsculas
df_final['Dn'] = df['Descrição Completa']  # Descrição
df_final['PrimeiroNome'] = df['Primeiro Nome']  # Primeiro nome já em maiúsculas
df_final['Sobrenome'] = df['Sobrenome']  # Sobrenome já em maiúsculas
df_final['Conta'] = df['RA'].apply(lambda x: str(x).zfill(8))  # RA com 8 dígitos
df_final['Email'] = df['E-mail']
df_final['Desc'] = 0  # Fixo para 0
df_final['Office'] = office  # Office sem aspas
df_final['Dep'] = criador  # Criador sem aspas
df_final['OU'] = destino  # Destino sem aspas
df_final['Pass'] = df['Senha']  # Senha

# Criar um objeto StringIO para armazenar os dados em formato CSV
output = io.StringIO()

# Escrever o cabeçalho
header = 'Nome,Dn,PrimeiroNome,Sobrenome,Conta,Email,Desc,Office,Dep,OU,Pass\n'
output.write(header)

# Escrever os dados
for _, row in df_final.iterrows():
    line = (
        f'"{row["Nome"]}",'
        f'"{row["Dn"]}",'
        f'{row["PrimeiroNome"]},'
        f'{row["Sobrenome"]},'
        f'{row["Conta"]},'
        f'{row["Email"]},'
        f'{row["Desc"]},'
        f'"{row["Office"]}",'
        f'{row["Dep"]},'
        f'"{row["OU"]}",'
        f'"{row["Pass"]}"\n'
    )
    output.write(line)

# Obter o conteúdo CSV como string
csv_data = output.getvalue()

# Passar os dados CSV para o script PowerShell
script_powershell = 'arquivo_powershell.ps1'

try:
    result = subprocess.run([
        "powershell",
        "-ExecutionPolicy", "Bypass",
        "-File", script_powershell,
        "-Data", csv_data  # Passando os dados CSV
    ], capture_output=True, text=True)

    # Exibir a saída do script
    print("Saída do PowerShell:")
    print(result.stdout)

    # Se houver algum erro, ele será capturado
    if result.stderr:
        print("Erro ao executar o script PowerShell:")
        print(result.stderr)

except Exception as e:
    print(f"Ocorreu um erro ao tentar executar o script PowerShell: {e}")
