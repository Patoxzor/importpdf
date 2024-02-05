import requests
import os
import sys
import subprocess
from urllib.request import urlretrieve
from config_json import ConfigManager

GITHUB_RELEASES_API_URL = 'https://api.github.com/repos/Patoxzor/importpdf/releases/latest'

def get_latest_release_info():
    response = requests.get(GITHUB_RELEASES_API_URL)
    return response.json() if response.ok else None

def download_and_install_update(url, version):
    local_filename = f'Mixpdf - {version}.exe'
    urlretrieve(url, local_filename)

    current_app_path = os.path.abspath(sys.argv[0])

    # Verificar se o caminho atual é um arquivo .exe
    if current_app_path.endswith('.exe'):
        old_app_path = f'{current_app_path}_old'
        if os.path.exists(old_app_path):
            os.remove(old_app_path)
        os.rename(current_app_path, old_app_path)
        os.replace(local_filename, current_app_path)

        # Atualizando a versão registrada
        ConfigManager.save_version(version)

        # Reiniciando a aplicação
        subprocess.Popen([current_app_path])
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
