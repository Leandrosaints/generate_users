import csv
import sys
import subprocess
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QFileDialog, QGridLayout, QMessageBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt


def executar_comando(comando):
    try:
        result = subprocess.run(
            comando,
            capture_output=True, text=True, shell=True
        )
        return result.stdout, result.stderr
    except Exception as e:
        return "", str(e)
class ExcelProcessor(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('GenServ')
        self.setGeometry(200, 200, 600, 400)
        self.setStyleSheet(self.get_style())

        # Criar layout principal
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        # Adicionar título
        title = QLabel('GenServ - User Generation Tool for Servers')
        title.setFont(QFont('Arial', 16))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        main_layout.addWidget(title)

        # Criar layout grid para os campos de entrada
        layout = QGridLayout()
        main_layout.addLayout(layout)

        # Labels e campos de entrada
        font = QFont('Arial', 12)

        self.file_label = QLabel('Arquivo Excel:')
        self.file_label.setFont(font)
        self.file_input = QLineEdit(self)
        self.file_input.setFont(font)
        self.file_button = QPushButton('Escolher Arquivo')
        self.file_button.setFont(font)
        self.file_button.clicked.connect(self.open_file_dialog)

        self.domain_label = QLabel('Domínio:')
        self.domain_label.setFont(font)
        self.domain_input = QLineEdit('@alunosenai.mt')
        self.domain_input.setFont(font)

        self.office_label = QLabel('Office:')
        self.office_label.setFont(font)
        self.office_input = QLineEdit('SENAI - Nova Mutum/MT')
        self.office_input.setFont(font)

        self.creator_label = QLabel('Criador:')
        self.creator_label.setFont(font)
        self.creator_input = QLineEdit('Criado por Jeferson Silva')
        self.creator_input.setFont(font)

        self.dest_label = QLabel('Destino:')
        self.dest_label.setFont(font)
        self.dest_input = QLineEdit('OU=QUA.415.089 ASSISTENTE DE RECURSOS HUMANOS...')
        self.dest_input.setFont(font)

        self.process_button = QPushButton('Processar')
        self.process_button.setFont(font)
        self.process_button.clicked.connect(self.process_file)

        # Botão para executar o PowerShell
        self.powershell_button = QPushButton('Executar PowerShell')
        self.powershell_button.setFont(font)
        self.powershell_button.clicked.connect(self.run_powershell)
        self.label_info = QLabel("Desenvolvido por Saints Technology - 2024")
        self.label_info.setFont(font)
        self.label_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Adicionar widgets ao layout
        layout.addWidget(self.file_label, 0, 0)
        layout.addWidget(self.file_input, 0, 1)
        layout.addWidget(self.file_button, 0, 2)

        layout.addWidget(self.domain_label, 1, 0)
        layout.addWidget(self.domain_input, 1, 1, 1, 2)

        layout.addWidget(self.office_label, 2, 0)
        layout.addWidget(self.office_input, 2, 1, 1, 2)

        layout.addWidget(self.creator_label, 3, 0)
        layout.addWidget(self.creator_input, 3, 1, 1, 2)

        layout.addWidget(self.dest_label, 4, 0)
        layout.addWidget(self.dest_input, 4, 1, 1, 2)

        layout.addWidget(self.process_button, 5, 0, 1, 3)
        layout.addWidget(self.powershell_button, 6, 0, 1, 3)
        layout.addWidget(self.label_info, 7, 0, 1, 3)



        # Estilizar o contêiner
        self.setStyleSheet(self.get_style())

    def get_style(self):
        """Define o estilo CSS para os widgets."""
        return """
            QWidget {
                background-color: #f2f2f2;
                border: 1px solid #ccc;
                border-radius: 8px;
                padding: 10px;
            }
            QLabel {
               
                color: #333;
               
            }
            QLineEdit {
                background-color: white;
                border: 1px solid #ccc;
                padding: 5px;
                border-radius: 4px;
            }
            QPushButton {
                background-color: #5cb85c;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #4cae4c;
            }
        """

    def open_file_dialog(self):
        # Abrir diálogo para escolher o arquivo Excel
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Escolher Arquivo', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.file_input.setText(file_path)

    def process_file(self):
        # Definir variáveis a partir dos inputs
        input_file = self.file_input.text()
        dominio = self.domain_input.text()
        office = self.office_input.text()
        criador = self.creator_input.text()
        destino = self.dest_input.text()

        if not input_file:
            QMessageBox.warning(self, 'Erro', 'Por favor, selecione um arquivo Excel.')
            return

        try:
            # Ler o arquivo Excel e processar dados
            df = pd.read_excel(input_file)

            # Ajuste dos dados conforme sua lógica
            df.columns = ['TURMA', 'RA', 'ALUNO', 'CPF', 'USUÁRIO', 'SENHA', 'E-MAIL / OFFICE 365']

            df = df[~df['RA'].astype(str).str.contains('RA', case=False)]
            df = df.dropna(subset=['RA', 'ALUNO'])

            def gerar_senha(ra):
                return ra[2:]

            def separar_nome(nome_completo):
                partes = nome_completo.split()
                primeiro_nome = partes[0].upper()  # Primeiro nome
                sobrenome = ' '.join(partes[1:]).upper() if len(
                    partes) > 1 else ''  # Todos os restantes formam o sobrenome
                return primeiro_nome, sobrenome

            df['Senha'] = df['RA'].apply(lambda x: gerar_senha(str(x)))
            df['Primeiro Nome'], df['Sobrenome'] = zip(*df['ALUNO'].apply(separar_nome))
            df['E-mail'] = df['RA'].apply(lambda x: str(x).zfill(8) + dominio)
            df['Descrição Completa'] = df['ALUNO'].str.upper() + ' - ' + office

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

            output_file = 'resultado.csv'
            #df_final.to_csv(output_file, index=False)

            with open(output_file, 'w', encoding='utf-8') as file:

                header = 'Nome,Dn,PrimeiroNome,Sobrenome,Conta,Email,Desc,Office,Dep,OU,Pass\n'
                file.write(header)

                # Escrevendo os dados
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
                file.write(line)
            QMessageBox.information(self, 'Sucesso', f'Arquivo processado e salvo como {output_file}')

        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Erro ao processar o arquivo: {e}')


    def run_powershell(self):
        #ps_script_path = 'arquivo_powershell.ps1'  # Certifique-se de usar o caminho correto
        try:
            script_powershell = 'arquivo_powershell.ps1'
            comando = f'powershell -ExecutionPolicy Bypass -File {script_powershell}'
            stdout, stderr = executar_comando(comando)
        #print(stderr)

            QMessageBox.information(self, 'Erro', f'Erro ao executar o PowerShell: {str(stderr)}')
        except:
            QMessageBox.warning(self, 'Erro', "Aplicaçao encontrou um erro!")
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())
