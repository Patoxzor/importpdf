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
    temp_filename = 'Mixpdf_new.exe'
    final_filename = 'Mixpdf.exe'

    urlretrieve(url, temp_filename)

    current_app_path = os.path.abspath(sys.argv[0])
    directory = os.path.dirname(current_app_path)  
    final_path = os.path.join(directory, final_filename)  # Caminho final do arquivo

    if current_app_path.endswith('.exe'):
        if os.path.exists(final_path):
            os.remove(final_path)  
        os.rename(temp_filename, final_path)  

        # Atualizando a versão registrada
        ConfigManager.save_version(version)

        # Reiniciando a aplicação
        subprocess.Popen([final_path], cwd=directory)
        sys.exit()
    else:
        # Atualização para ambiente de desenvolvimento, não reiniciar
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
