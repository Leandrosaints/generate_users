import requests
from packaging import version
from PyQt6.QtWidgets import QMessageBox

# Versão atual da aplicação
CURRENT_VERSION = "1.0.0"

# URL do arquivo verify_version.json no GitHub (usando raw.githubusercontent.com)
VERSION_JSON_URL = "https://raw.githubusercontent.com/Leandrosaints/generate_users/b4665669c8aaaebbf5e509246c5fdd697cf414de/verify_version.json"


# Opção alternativa: usar a branch principal
# VERSION_JSON_URL = "https://raw.githubusercontent.com/Leandrosaints/generate_users/main/verify_version.json"

def check_for_updates(parent):
    """
    Verifica se há uma nova versão da aplicação disponível.
    parent: O widget pai para o QMessageBox (geralmente self na sua classe principal).
    """
    try:
        # Faz a requisição HTTP para o arquivo verify_version.json
        response = requests.get(VERSION_JSON_URL, timeout=5)
        response.raise_for_status()  # Levanta uma exceção se a requisição falhar

        # Lê o conteúdo do arquivo JSON
        remote_data = response.json()
        remote_version = remote_data.get("version")
        download_url = remote_data.get("download_url")
        changelog = remote_data.get("changelog", "Nenhuma descrição disponível.")

        if not remote_version:
            QMessageBox.warning(parent, "Erro",
                                "Não foi possível obter a versão remota do arquivo verify_version.json.")
            return

        # Compara as versões
        if version.parse(remote_version) > version.parse(CURRENT_VERSION):
            # Nova versão disponível
            message = (
                f"Uma nova versão ({remote_version}) está disponível!\n\n"
                f"Versão atual: {CURRENT_VERSION}\n"
                f"Changelog: {changelog}\n\n"
                f"Você pode baixar a nova versão em:\n{download_url}\n\n"
                f"Deseja abrir o link de download?"
            )
            reply = QMessageBox.question(
                parent,
                "Nova Versão Disponível",
                message,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                # Abre o link de download no navegador padrão
                import webbrowser
                webbrowser.open(download_url)
        else:
            # Nenhuma nova versão disponível
            QMessageBox.information(parent, "Atualização", "Você está usando a versão mais recente.")

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            QMessageBox.critical(
                parent,
                "Erro",
                "Arquivo verify_version.json não encontrado no GitHub.\n"
                "Verifique se o arquivo existe no repositório e se o URL está correto.\n"
                f"URL tentado: {VERSION_JSON_URL}"
            )
        else:
            QMessageBox.warning(parent, "Erro", f"Falha ao verificar atualizações: {e}")
    except requests.exceptions.RequestException as e:
        QMessageBox.warning(parent, "Erro", f"Falha ao conectar ao GitHub: {e}")
    except ValueError as e:
        QMessageBox.warning(parent, "Erro", f"Erro ao processar o arquivo verify_version.json: {e}")


'''# Exemplo de uso na sua classe principal (ExcelProcessor ou similar)
def __init__(self):
    super().__init__()
    # ... resto do seu código de inicialização ...

    # Verifica atualizações ao iniciar a aplicação
    self.check_for_updates()'''