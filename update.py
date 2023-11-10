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
    local_filename = f'Mixpdf - {version}.exe'

    urlretrieve(url, local_filename)

    current_app_path = os.path.abspath(sys.argv[0])
    old_app_path = f'{current_app_path}_old'

    if os.path.exists(old_app_path):
        os.remove(old_app_path)
    os.rename(current_app_path, old_app_path)
    os.replace(local_filename, current_app_path)
    ConfigManager.save_version(version)
    subprocess.Popen([current_app_path])
    sys.exit()

def notify_user_of_update(latest_version):
    root = tk.Tk()
    root.withdraw()
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
        latest_version = latest_release['tag_name'].replace('V', '')

        # Comparando a versão atual com a versão mais recente
        if latest_version != current_version:
            print(f'Nova versão disponível: {latest_version}')
            assets = latest_release['assets']
            if assets:
                browser_download_url = assets[0]['browser_download_url']
                download_and_install_update(browser_download_url, latest_version)
        else:
            print("Você já está com a versão mais recente.")

