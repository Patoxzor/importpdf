import requests
import os
import sys
import subprocess
from urllib.request import urlretrieve
from config_json import ConfigManager
import tkinter as tk
from tkinter import messagebox

GITHUB_RELEASES_API_URL = 'https://api.github.com/repos/Patoxzor/importpdf/releases/latest'

def get_latest_release_info():
    response = requests.get(GITHUB_RELEASES_API_URL)
    return response.json() if response.ok else None

def download_and_install_update(url, version):
    # Define o nome do arquivo local para o novo executável
    local_filename = f'app-{version}.exe'

    # Baixa o novo executável
    urlretrieve(url, local_filename)

    # Obtem o caminho do executável atual
    current_app_path = os.path.abspath(sys.argv[0])

    # Renomeia o executável atual (por exemplo, para 'app_old.exe')
    old_app_path = f'{current_app_path}_old'
    os.rename(current_app_path, old_app_path)

    # Move o novo executável para o local do executável atual
    os.replace(local_filename, current_app_path)

    # Reinicia o aplicativo
    subprocess.Popen([current_app_path])

    # Sai do aplicativo atual para que a nova versão possa iniciar
    sys.exit()

def notify_user_of_update(latest_version):
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    response = messagebox.askyesno("Atualização Disponível", 
                                   f"Uma nova versão {latest_version} está disponível. Deseja atualizar agora?")
    if response:
        return True
    else:
        return False

def check_for_updates():
    current_version = ConfigManager.load_version()
    latest_release = get_latest_release_info()
    if latest_release:
        latest_version = latest_release['tag_name'].replace('v', '')
        if latest_version != current_version:
            print(f'Nova versão disponível: {latest_version}')
            assets = latest_release['assets']
            if assets:
                # Assumindo que o asset é um arquivo exe da nova versão
                browser_download_url = assets[0]['browser_download_url']
                download_and_install_update(browser_download_url, latest_version)

