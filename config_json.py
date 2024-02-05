import json
import os

class ConfigManager:
    @staticmethod
    def save_to_file(data, filename="config.json"):
        try:
            with open(filename, 'w') as file:
                json.dump(data, file)
        except IOError as e:
            print(f"Erro ao salvar no arquivo {filename}: {e}")

    @staticmethod
    def load_from_file(filename="config.json"):
        try:
            with open(filename, 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return {}
        except json.JSONDecodeError as e:
            print(f"Erro ao carregar dados de {filename}: {e}")
            return {}
        except IOError as e:
            print(f"Erro ao abrir o arquivo {filename}: {e}")
            return {}

    @staticmethod
    def update_data(filename, key, value):
        data = ConfigManager.load_from_file(filename)
        data[key] = value
        ConfigManager.save_to_file(data, filename)

    @staticmethod
    def save_color_preferences(data, filename="config.json"):
        ConfigManager.update_data(filename, 'color_preferences', data)

    @staticmethod
    def load_color_preferences(filename="config.json"):
        default_colors = {'background': '#ffffff', 'buttons': '#0000ff', 'button_frame': '#ffffff'}
        data = ConfigManager.load_from_file(filename)
        color_preferences = data.get('color_preferences', default_colors)

        for key, value in default_colors.items():
            color_preferences.setdefault(key, value)
        return color_preferences

    @staticmethod
    def save_version(version, filename="config.json"):
        ConfigManager.update_data(filename, 'version', version)

    @staticmethod
    def load_version(filename="config.json"):
        data = ConfigManager.load_from_file(filename)
        return data.get('version', '1.0.2')

config = ConfigManager()
config.save_version("1.0.2")
print(config.load_version())
