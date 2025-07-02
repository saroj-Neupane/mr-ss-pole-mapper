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
        """Get default configuration - hardcoded CONSUMERS.json values"""
        return {
            "power_company": "CONSUMERS ENERGY",
            "telecom_providers": [
                "Proposed MetroNet",
                "Lightower",
                "Comcast",
                "Verizon",
                "AT&T",
                "Zayo",
                "Jackson ISD"
            ],
            "power_keywords": [
                "Primary",
                "Secondary",
                "Neutral",
                "Secondary Drip Loop",
                "Riser",
                "Transformer"               
            ],
            "comm_keywords": [
                "catv com",
                "telco com", 
                "fiber optic com",
                "insulator",
                "power guy",
                "catv",
                "telco",
                "fiber",
                "communication",
                "comm"
            ],
            "ignore_scid_keywords": [
                "AT&T",
                "Unknown",
                "POLE",
                "FOREIGN"
            ],
            "telecom_keywords": {
                "Proposed MetroNet": [
                    "MetroNet",
                    "MNT",
                    "Proposed MNT"
                ],
                "AT&T": [
                    "AT&T",
                    "ATT"
                ],
                "Verizon": [
                    "verizon",
                    "Verizon"
                ],
                "Comcast": [
                    "comcast",
                    "Comcast"
                ],
                "Lightower": [
                    "lightower",
                    "Lightower"
                ],
                "Zayo": [
                    "zayo",
                    "Zayo"
                ],
                "Jackson ISD": [
                    "JACKSON ISD"
                ]
            },
            "output_settings": {
                "header_row": 3,
                "data_start_row": 4,
                "worksheet_name": "Consumers pg1"
            },
            "processing_options": {
                "use_geocoding": True,
                "open_output": False,
                "debug_mode": False,
                "use_decimal_format": False,
                "use_qc_routing": False,
                "span_length_tolerance": 3.0
            },
            "tension_calculator": {
                "file_path": "Test Files/Metronet tension calculator.xlsm",
                "worksheet_name": "Calculations"
            },
            "manual_routes_options": {
                "use_manual_routes": False,
                "clear_existing_routes": False
            },
            "column_mappings": [
                [
                    "System",
                    "Line Number",
                    "Line No."
                ],
                [
                    "Pole",
                    "Number",
                    "Pole"
                ],
                [
                    "Pole",
                    "Address",
                    "Pole Address (if available)"
                ],
                [
                    "Pole",
                    "Height & Class",
                    "Pole Height & Class"
                ],
                [
                    "Power",
                    "Lowest Height",
                    "Secondary or Neutral Power Height (Height of Lowest Power Conductor or Equipment, excluding streetlights)"
                ],
                [
                    "Street Light",
                    "Lowest Height",
                    "Streetlight"
                ],
                [
                    "Pole",
                    "To Pole",
                    "To Pole"
                ],
                [
                    "Span",
                    "Length",
                    "Pole to Pole Span Length (from starting point)"
                ],
                [
                    "New Guy",
                    "Lead",
                    "Guy Lead"
                ],
                [
                    "New Guy",
                    "Direction",
                    "Guy Direction"
                ],
                [
                    "New Guy",
                    "Size",
                    "Guy Size"
                ],
                [
                    "Pole",
                    "MR Notes",
                    "Notes (Items that need to be performed by Consumers Energy or other Companies)"
                ],
                [
                    "Proposed MetroNet",
                    "Attachment Ht",
                    "Proposed height of new attachment point"
                ],
                [
                    "Proposed MetroNet",
                    "Midspan Ht",
                    "Final Mid Span Ground Clearance of Proposed Attachment"
                ],
                [
                    "Verizon",
                    "Attachment Ht",
                    "Verizon"
                ],
                [
                    "Zayo",
                    "Midspan Ht",
                    "Zayo"
                ],
                [
                    "AT&T",
                    "Attachment Ht",
                    "AT&T"
                ],
                [
                    "Jackson ISD",
                    "Attachment Ht",
                    "Jackson ISD"
                ],
                [
                    "Comcast",
                    "Attachment Ht",
                    "Comcast"
                ],
                [
                    "Proposed MetroNet",
                    "Tension",
                    "Heavy Loaded Tension (NESC Rule 251)"
                ]
            ],
            "decimal_measurements": False
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