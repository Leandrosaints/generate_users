import requests
from packaging import version
from PyQt6.QtWidgets import QMessageBox

# Versão atual da aplicação
CURRENT_VERSION = "1.0.0"

# URL do arquivo version.json no GitHub (usando raw.githubusercontent.com)
VERSION_JSON_URL = "https://raw.githubusercontent.com/Leandrosaints/generate_users/main/version.json"


def check_for_updates(parent):
    """
    Verifica se há uma nova versão da aplicação disponível.
    parent: O widget pai para o QMessageBox (geralmente self na sua classe principal).
    """
    try:
        # Faz a requisição HTTP para o arquivo version.json
        response = requests.get(VERSION_JSON_URL, timeout=5)
        response.raise_for_status()  # Levanta uma exceção se a requisição falhar

        # Verifica se a resposta é um JSON válido
        content_type = response.headers.get('Content-Type', '')
        if 'application/json' not in content_type:
            QMessageBox.warning(
                parent,
                "Erro",
                f"A resposta do servidor não é um JSON válido.\n"
                f"Tipo de conteúdo recebido: {content_type}\n"
                f"Conteúdo bruto (primeiros 500 caracteres): {response.text[:500]}"
            )
            return

        # Lê o conteúdo do arquivo JSON
        try:
            remote_data = response.json()
        except ValueError as json_error:
            QMessageBox.critical(
                parent,
                "Erro",
                f"Erro ao processar o arquivo version.json como JSON:\n{json_error}\n"
                f"Conteúdo bruto (primeiros 500 caracteres): {response.text[:500]}"
            )
            return

        # Verifica se as chaves esperadas estão presentes no JSON
        remote_version = remote_data.get("version")
        download_url = remote_data.get("download_url")
        changelog = remote_data.get("changelog", "Nenhuma descrição disponível.")

        # Verifica se os campos obrigatórios estão presentes
        if not remote_version:
            QMessageBox.warning(parent, "Erro", "Campo 'version' não encontrado no arquivo version.json.")
            return
        if not download_url:
            QMessageBox.warning(parent, "Erro", "Campo 'download_url' não encontrado no arquivo version.json.")
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
                "Arquivo version.json não encontrado no GitHub.\n"
                "Verifique se o arquivo existe no repositório e se o URL está correto.\n"
                f"URL tentado: {VERSION_JSON_URL}"
            )
        else:
            QMessageBox.warning(parent, "Erro", f"Falha ao verificar atualizações: {e}")
    except requests.exceptions.RequestException as e:
        QMessageBox.warning(parent, "Erro", f"Falha ao conectar ao GitHub: {e}")


