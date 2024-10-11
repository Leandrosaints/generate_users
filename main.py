import streamlit as st
import pandas as pd
import subprocess




def executar_comando(comando):
    try:
        result = subprocess.run(
            comando,
            capture_output=True, text=True, shell=True
        )
        return result.stdout, result.stderr
    except Exception as e:
        return "", str(e)



# Configuração da página
st.set_page_config(page_title="Gerador de Usuários", layout="wide")

# Estilização do container com CSS para bordas, background e formatação
st.markdown(
    """
    <style>
        h1 {
        font-family: "Source Sans Pro", sans-serif;
        font-weight: 700;
        text-align: center;
        color: rgb(49, 51, 63);
        padding: 1.25rem 0px 0px;
        margin: 0px;
        line-height: 1.2;
    }
      h5 {
        font-family: "Source Sans Pro", sans-serif;

        text-align: center;
        color: rgb(49, 51, 63);
        padding: 1.25rem 0px 1rem;
        margin: 0px;

    }
    .styled-container {
        background-color: #f9f9f9;  /* Cor de fundo suave */
        border: 1px solid #ddd;  /* Bordas sutis */
        border-radius: 8px;  /* Cantos arredondados */
        padding: 20px;  /* Espaçamento interno */
        box-shadow: 2px 2px 8px rgba(0, 0, 0, 0.1);  /* Sombra leve */
        margin-top: 20px;  /* Espaçamento do topo */
    }

    .st-emotion-cache-1r4qj8v {
        position: absolute;
        background:#7997a10f;
        color: rgb(49, 51, 63);
        inset: 0px;
        color-scheme: light;
        overflow: hidden;
    }
    .input-container {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
    }
    .input-container label {
        width: 150px;
        font-weight: bold;
        color: #333;
    }
    .stTextInput > div > input {
        border: 1px solid #ccc;
        padding: 5px;
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Seção principal dentro de um container estilizado
with st.container():
    st.markdown("<h1>GenServ - Gerador de Usuário Interno</h1>", unsafe_allow_html=True)
    st.markdown("<h5>Preencha as informações necessárias para gerar os usuários no servidor:</h5>",
                unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Configurações Gerais")
        st.markdown('<div class="input-container"><label>Domínio de E-mail</label></div>', unsafe_allow_html=True)
        dominio = st.text_input("", value="@alunosenai.mt", label_visibility="collapsed")

        st.markdown('<div class="input-container"><label>Office</label></div>', unsafe_allow_html=True)
        office = st.text_input("", value="SENAI - Nova Mutum/MT", label_visibility="collapsed")

        st.markdown('<div class="input-container"><label>Criador</label></div>', unsafe_allow_html=True)
        criador = st.text_input("", value="Criado por Jeferson Silva", label_visibility="collapsed")

        st.markdown('<div class="input-container"><label>Destino</label></div>', unsafe_allow_html=True)
        destino = st.text_input("",
                                value="OU=QUA.415.089 ASSISTENTE DE RECURSOS HUMANOS COM INFORMÁTICA,OU=CURSOS,OU=SENAINMT,OU=SENAI,OU=SFIEMT-EDU,DC=SESISENAIMT,DC=EDU",
                                label_visibility="collapsed")

    with col2:
        st.markdown("### Arquivo de Entrada")
        st.markdown('<div class="input-container"><label>Planilha usuarios (.xlsx)</label></div>',
                    unsafe_allow_html=True)
        input_file = st.file_uploader("", type="xlsx", label_visibility="collapsed")

        st.markdown('<div class="input-container"><label>Nome do Arquivo de Saída</label></div>',
                    unsafe_allow_html=True)

        output_file = 'resultado.csv'

        st.markdown('<div class="input-container"><label>Executar PowerShell</label></div>', unsafe_allow_html=True)
        execute = st.checkbox("Executar PowerShell após a geração?", value=False)
        # Disponibilizar o arquivo para download



    # Processamento
if input_file and st.button("Gerar Usuários"):
    # Ler a planilha
    df = pd.read_excel(input_file)

    # Ajustar os nomes das colunas conforme a estrutura da planilha
    df.columns = ['TURMA', 'RA', 'ALUNO', 'CPF', 'USUÁRIO', 'SENHA', 'E-MAIL / OFFICE 365']

    # Remover linhas indesejadas e processar os dados conforme necessário
    df = df[~df['RA'].astype(str).str.contains('RA', case=False)]
    df = df[~df['ALUNO'].astype(str).str.contains('ALUNO', case=False)]
    df = df.dropna(subset=['RA', 'ALUNO'])


    # Função para gerar a senha a partir do RA
    def gerar_senha(ra):
        return ra[2:]  # Remove os dois primeiros caracteres do RA


    # Função para separar o primeiro nome e o restante
    def separar_nome(nome_completo):
        partes = nome_completo.split(' ', 1)
        primeiro_nome = partes[0].upper()
        sobrenome = partes[1].upper() if len(partes) > 1 else ''
        return primeiro_nome, sobrenome


    # Aplicar funções para gerar dados adicionais
    df['Senha'] = df['RA'].apply(lambda x: gerar_senha(str(x)))
    df['Primeiro Nome'], df['Sobrenome'] = zip(*df['ALUNO'].apply(separar_nome))
    df['E-mail'] = df['RA'].apply(lambda x: str(x).zfill(8) + dominio)
    df['Descrição Completa'] = df['ALUNO'].str.upper() + ' - ' + office

    # Montar a estrutura de saída conforme solicitado
    df_final = pd.DataFrame()
    df_final['Nome'] = df['ALUNO'].str.upper()
    df_final['Dn'] = df['Descrição Completa']
    df_final['PrimeiroNome'] = df['Primeiro Nome']
    df_final['Sobrenome'] = df['Sobrenome']
    df_final['Conta'] = df['RA'].apply(lambda x: str(x).zfill(8))
    df_final['Email'] = df['E-mail']
    df_final['Desc'] = df['CPF']
    df_final['Office'] = office
    df_final['Dep'] = criador
    df_final['OU'] = destino
    df_final['Pass'] = df['Senha']

    # Salvar automaticamente o arquivo CSV
    df_final.to_csv(output_file, index=False)

    st.success(f"Arquivo {output_file} gerado com sucesso!")

    # Executar o script PowerShell se selecionado
    if execute:
        # Executar o script PowerShell se selecionado
        script_powershell = 'arquivo_powershell.ps1'
        comando = f'powershell -ExecutionPolicy Bypass -File {script_powershell}'
        stdout, stderr = executar_comando(comando)

        # Exibir a saída
        st.text("Saída do PowerShell:")
        st.text(stdout)

        if stderr:
            st.error("Erro ao executar o script PowerShell:")
            st.error(stderr)

    with open(output_file, 'rb') as file:
        btn = st.download_button(
            label="Baixar arquivo CSV",
            data=file,
            file_name=output_file,
            mime='text/csv'
        )

st.markdown("---")  # Linha divisória

# Adiciona uma barra de desenvolvedor no rodapé da página
st.markdown(
    """
    <div style='text-align: center; color: gray; margin-top: 10px;'>
        Desenvolvido por <strong>Saints Technology</strong> - 2024
    </div>
    """,
    unsafe_allow_html=True
)
