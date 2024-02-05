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
    local_filename = 'Mixpdf.exe'
    temp_filename = 'Mixpdf_new.exe'  
    urlretrieve(url, temp_filename)

    current_app_path = os.path.abspath(sys.argv[0])
    directory = os.path.dirname(current_app_path)
    final_path = os.path.join(directory, local_filename)

    if os.path.exists(final_path):
        os.remove(final_path) 
    os.rename(temp_filename, final_path)  

    # Atualizando a versão registrada
    ConfigManager.save_version(version)

def check_for_updates():
    current_version = ConfigManager.load_version()
    latest_release_info = get_latest_release_info()

    if latest_release_info:
        latest_version = latest_release_info['tag_name'].replace('V', '').replace('v', '')  # Remove 'V' ou 'v' da tag
        assets = latest_release_info['assets']
        browser_download_url = None

        for asset in assets:
            if 'mixpdf.exe' in asset['name'].lower():  # Procura por 'mixpdf.exe' no nome do asset, ignorando maiúsculas
                browser_download_url = asset['browser_download_url']
                break  # Sai do loop assim que encontrar o primeiro match

        if latest_version != current_version and browser_download_url:
            download_and_install_update(browser_download_url, latest_version)
        else:
            print("Você já está na versão mais recente ou o arquivo não foi encontrado.")

check_for_updates()
