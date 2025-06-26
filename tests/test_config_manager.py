import unittest
from src.core.config_manager import ConfigManager
from pathlib import Path

class TestConfigManager(unittest.TestCase):

    def setUp(self):
        base_dir = Path(__file__).resolve().parent.parent / 'src'
        self.config_manager = ConfigManager(base_dir)

    def test_get_default_config(self):
        default_config = self.config_manager.get_default_config()
        self.assertIn("power_company", default_config)
        self.assertIn("telecom_providers", default_config)
        self.assertIn("power_keywords", default_config)
        self.assertIn("telecom_keywords", default_config)
        self.assertIn("output_settings", default_config)
        self.assertIn("column_mappings", default_config)

    def test_get_available_configs(self):
        available_configs = self.config_manager.get_available_configs()
        self.assertIn("Default", available_configs)

    def test_load_config(self):
        config = self.config_manager.load_config("Default")
        self.assertEqual(config["power_company"], "Xcel")

    def test_save_config(self):
        config_name = "test_config"
        config_data = self.config_manager.get_default_config()
        success = self.config_manager.save_config(config_name, config_data)
        self.assertTrue(success)

    def test_delete_config(self):
        config_name = "test_config"
        self.config_manager.save_config(config_name, self.config_manager.get_default_config())
        success = self.config_manager.delete_config(config_name)
        self.assertTrue(success)

if __name__ == '__main__':
    unittest.main()