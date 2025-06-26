import json
import logging
from pathlib import Path

class ConfigManager:
    """Manages configuration loading, saving, and defaults"""
    
    def __init__(self, base_dir=None):
        if base_dir is None:
            # Use current working directory if no base_dir provided
                    base_dir = Path.cwd()
        self.base_dir = Path(base_dir)
        self.configs_dir = self.base_dir / "configurations"
        self.configs_dir.mkdir(exist_ok=True)
    
    def get_default_config(self):
        return {
            "power_company": "Xcel",
            "telecom_providers": [
                "Proposed MetroNet", "Lightower", "Comcast", "Verizon", 
                "AT&T", "CATV", "Telephone Company", "Fiber"
            ],
            "power_keywords": ["Primary", "Secondary", "Neutral", "Secondary Drip Loop", "Riser", "Transformer"],
            "telecom_keywords": {
                "Proposed MetroNet": ["MetroNet", "MNT", "Proposed MNT"]
            },
            "output_settings": {
                "header_row": 3,
                "data_start_row": 4,
                "worksheet_name": "Consumers pg1"
            },
            "processing_options": {
                "use_geocoding": False,
                "open_output": False,
                "debug_mode": False,
                "use_decimal_format": False,
                "use_qc_routing": False
            },
            "manual_routes_options": {
                "use_manual_routes": False,
                "clear_existing_routes": False
            },
            "column_mappings": []
        }
    
    def get_config_file_path(self, config_name):
        """Get file path for configuration"""
        if config_name == "Default":
            return self.base_dir / "pole_mapper_config.json"
        else:
            return self.configs_dir / f"{config_name}.json"
    
    def get_available_configs(self):
        """Get list of available configurations"""
        configs = ["Default"]
        try:
            for file in self.configs_dir.glob("*.json"):
                configs.append(file.stem)
        except:
            pass
        return configs
    
    def load_config(self, config_name):
        """Load configuration"""
        config = self.get_default_config()
        config_file = self.get_config_file_path(config_name)
        
        if config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    loaded = json.load(f)
                    config.update(loaded)
                logging.info(f"Configuration for '{config_name}' successfully loaded from {config_file}")
            except Exception as e:
                logging.warning(f"Failed to load configuration from {config_file}: {e}")
        
        return config
    
    def save_config(self, config_name, config):
        """Save configuration"""
        try:
            config_file = self.get_config_file_path(config_name)
            config_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=2)
            
            logging.info(f"Configuration for '{config_name}' successfully saved to {config_file}")
            return True
        except Exception as e:
            logging.error(f"Failed to save configuration to {config_file}: {e}")
            return False
    
    def delete_config(self, config_name):
        """Delete configuration"""
        if config_name == "Default":
            return False
        
        try:
            config_file = self.get_config_file_path(config_name)
            if config_file.exists():
                config_file.unlink()
            logging.info(f"Configuration '{config_name}' deleted from {config_file}")
            return True
        except Exception as e:
            logging.error(f"Failed to delete configuration '{config_name}' at {config_file}: {e}")
            return False