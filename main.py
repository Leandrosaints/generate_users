import streamlit as st
import pandas as pd
import subprocess

# Configurar a página do Streamlit
st.set_page_config(page_title="Gerenciamento de Planilhas", layout="wide")

# Estilo CSS personalizado para melhorar a aparência
st.markdown("""
    <style>
    .main {
        background-color:#959595;
        padding: 20px;
        border-radius: 8px;
    }
    .stButton > button {
        background-color: #007BFF;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 10px 20px;
        cursor: pointer;
    }
    .stButton > button:hover {
        background-color: #0056b3;
    }
    .legend {
        font-weight: bold;
        margin-bottom: 5px;
        margin-top: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# Título da aplicação
st.title("📊 Gerenciamento de Planilhas de Alunos")
st.subheader("Organize e processe suas planilhas de forma rápida e eficiente")

# Dividir a interface em três colunas
col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    st.header("Upload da Planilha")
    uploaded_file = st.file_uploader("Carregue a planilha de alunos (.xlsx)", type="xlsx")

with col2:
    st.header("Configurações")

    # Criar colunas internas para alinhar os campos de entrada lado a lado
    col2_1, col2_2 = st.columns(2)

    with col2_1:
        st.markdown('<div class="legend">Domínio de E-mail</div>', unsafe_allow_html=True)
        dominio = st.text_input("", value="@alunosenai.mt", key="dominio")

        st.markdown('<div class="legend">Criador</div>', unsafe_allow_html=True)
        criador = st.text_input("", value="Criado por Jeferson Silva", key="criador")

    with col2_2:
        st.markdown('<div class="legend">Office</div>', unsafe_allow_html=True)
        office = st.text_input("", value="SENAI - Nova Mutum/MT", key="office")

        st.markdown('<div class="legend">Destino OU</div>', unsafe_allow_html=True)
        destino = st.text_area("", value="OU=QUA.415.089 ASSISTENTE DE RECURSOS HUMANOS COM INFORMÁTICA,OU=CURSOS,OU=SENAINMT,OU=SENAI,OU=SFIEMT-EDU,DC=SESISENAIMT,DC=EDU", key="destino")

with col3:
    st.header("Ações")
    st.write("Escolha as ações para processar a planilha e executar scripts.")

    # Botão para processar a planilha
    if st.button("📥 Processar Planilha e Gerar CSV"):
        if uploaded_file is not None:
            # Ler a planilha enviada pelo usuário
            df = pd.read_excel(uploaded_file)

            # Ajustar os nomes das colunas conforme a estrutura da planilha
            df.columns = ['TURMA', 'RA', 'ALUNO', 'CPF', 'USUÁRIO', 'SENHA', 'E-MAIL / OFFICE 365']

            # Remover linhas desnecessárias
            df = df[~df['RA'].astype(str).str.contains('RA', case=False)]
            df = df[~df['ALUNO'].astype(str).str.contains('ALUNO', case=False)]
            df = df.dropna(subset=['RA', 'ALUNO'])

            # Funções para gerar senha e separar nome
            def gerar_senha(ra):
                return ra[2:]

            def separar_nome(nome_completo):
                partes = nome_completo.split(' ', 1)
                primeiro_nome = partes[0].upper()
                sobrenome = partes[1].upper() if len(partes) > 1 else ''
                return primeiro_nome, sobrenome

            # Aplicar funções e adicionar colunas
            df['Senha'] = df['RA'].apply(lambda x: gerar_senha(str(x)))
            df['Primeiro Nome'], df['Sobrenome'] = zip(*df['ALUNO'].apply(separar_nome))
            df['E-mail'] = df['RA'].apply(lambda x: str(x).zfill(8) + dominio)
            df['Descrição Completa'] = df['ALUNO'].str.upper() + ' - ' + office

            # Criar DataFrame final
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

            # Salvar o DataFrame final em um arquivo CSV
            output_file = 'resultado.csv'
            df_final.to_csv(output_file, index=False, quotechar='"')

            st.success('CSV gerado com sucesso!')

            # Link para download do arquivo CSV
            st.download_button(
                label="Baixar CSV gerado",
                data=open(output_file, 'rb').read(),
                file_name=output_file,
                mime='text/csv'
            )
        else:
            st.error("Por favor, carregue uma planilha válida.")

    # Botão para executar o script PowerShell
    if st.button("⚡ Executar Script PowerShell"):
        script_powershell = 'arquivo_powershell.ps1'

        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_powershell],
                capture_output=True, text=True
            )

            # Exibir a saída do script PowerShell
            st.text("Saída do PowerShell:")
            st.code(result.stdout)

            # Exibir erros, se houver
            if result.stderr:
                st.error(f"Erro ao executar o script PowerShell:\n{result.stderr}")
        except Exception as e:
            st.error(f"Ocorreu um erro ao tentar executar o script PowerShell: {e}")
