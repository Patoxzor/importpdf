from config_json import ConfigManager

def increment_version(version_str):
    # Divide a string da versão em major, minor, patch
    major, minor, patch = map(int, version_str.split('.'))
    # Incrementa o número do patch
    patch += 1
    # Retorna a nova string de versão
    return f'{major}.{minor}.{patch}'

def main():
    # Carrega a versão atual
    current_version = ConfigManager.load_version()
    print(f'Versão atual: {current_version}')
    
    # Incrementa a versão
    new_version = increment_version(current_version)
    print(f'Nova versão: {new_version}')
    
    # Salva a nova versão no arquivo de configuração
    ConfigManager.save_version(new_version)
    print(f'Versão atualizada com sucesso para {new_version}.')

if __name__ == '__main__':
    main()
