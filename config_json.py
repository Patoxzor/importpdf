import json

class ConfigManager:
    @staticmethod
    def save_to_file(data, filename="config.json"):
        with open(filename, 'w') as file:
            json.dump(data, file)

    @staticmethod
    def load_from_file(filename="config.json"):
        try:
            with open(filename, 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return {}
        
    @staticmethod
    def save_color_preferences(data, filename="config.json"):
        current_data = ConfigManager.load_from_file(filename)
        current_data['color_preferences'] = data
        ConfigManager.save_to_file(current_data, filename)

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
        current_data = ConfigManager.load_from_file(filename)
        current_data['version'] = version
        ConfigManager.save_to_file(current_data, filename)

    @staticmethod
    def load_version(filename="config.json"):
        data = ConfigManager.load_from_file(filename)
        return data.get('version', '1.0.0')