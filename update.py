import requests
import os
import sys
import subprocess
from urllib.request import urlretrieve
from config_json import ConfigManager
import time 

GITHUB_RELEASES_API_URL = 'https://api.github.com/repos/Patoxzor/importpdf/releases/latest'

def get_latest_release_info():
    response = requests.get(GITHUB_RELEASES_API_URL)
    if response.ok:
        return response.json()
    return None

def download_and_install_update(url, version):
    local_filename = f'Mixpdf - {version}.exe'
    urlretrieve(url, local_filename)
    
    # Caminho absoluto do aplicativo atual
    current_app_path = os.path.abspath(sys.argv[0])

    if current_app_path.endswith('.exe'):
        old_app_path = f'{current_app_path}_old'
        # Remove a versão antiga, se existir
        if os.path.exists(old_app_path):
            os.remove(old_app_path)
        os.rename(current_app_path, old_app_path)
        
        # Substitui o executável antigo pelo novo
        os.replace(local_filename, current_app_path)

        # Atualiza a versão registrada
        ConfigManager.save_version(version)

        # Atraso para garantir que o arquivo esteja liberado para execução
        time.sleep(2)

        # Reiniciando a aplicação com o caminho atualizado e shell=True
        subprocess.Popen([current_app_path], shell=True)
        sys.exit()
    else:
        print("Atualização baixada. Reinicie manualmente o aplicativo.")

def check_for_updates():
    current_version = ConfigManager.load_version()
    latest_release = get_latest_release_info()

    if latest_release:
        latest_version = latest_release['tag_name'].replace('V', '')
        assets = latest_release['assets']
        browser_download_url = None

        for asset in assets:
            if 'Mixpdf' in asset['name']:
                browser_download_url = asset['browser_download_url']

        if latest_version != current_version and browser_download_url:
            return latest_version, browser_download_url

    return None, None
