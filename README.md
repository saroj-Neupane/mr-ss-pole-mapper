# MR SS Pole Mapper

## Overview

The MR SS Pole Mapper is a Python application designed to assist in managing and processing pole data, including attachments and connections. It provides a graphical user interface (GUI) for user interactions and utilizes various core functionalities to handle data efficiently.

## Features

- **Configuration Management**: Multiple configurations with customizable settings
- **Excel Data Processing**: Read and process pole data from Excel files
- **Attachment Data Handling**: Process SCID-based attachment information
- **Geocoding Support**: Convert coordinates to addresses automatically
- **Manual Route Definition**: Override automatic connection detection
- **Flexible Output Mapping**: Map processed data to custom Excel templates
- **Real-time Processing Log**: Monitor processing progress and debug issues

## Quick Start

1. Install Python 3.7+ and required dependencies: `pip install -r requirements.txt`
2. Run the application: `python src/main.py`
3. Configure your settings in the Configuration tab
4. Select input files and process data in the Processing tab

## Project Structure

```
mr-ss-pole-mapper/
├── src/                     # Source code for the application
│   ├── main.py              # Entry point for the application
│   ├── gui/                 # GUI components
│   │   └── main_window.py   # Main window implementation with full PoleMapperApp
│   ├── core/                # Core functionalities
│   │   ├── utils.py         # Utility functions
│   │   ├── config_manager.py # Configuration management
│   │   ├── route_parser.py   # Route parsing and validation
│   │   ├── attachment_data_reader.py # Reading attachment data
│   │   ├── geocoder.py      # Geocoding functionality
│   │   └── pole_data_processor.py # Main data processing logic
│   │   └── pole_data_processor.py # Processing pole data
│   └── models/              # Data models
│       ├── __init__.py      # Marks the models directory as a Python package
│       └── data_models.py    # Data models definitions
├── configs/                 # Configuration files
│   └── .gitkeep             # Keeps the configs directory in version control
├── cache/                   # Cache files
│   └── .gitkeep             # Keeps the cache directory in version control
├── logs/                    # Log files
│   └── .gitkeep             # Keeps the logs directory in version control
├── tests/                   # Unit tests
│   ├── __init__.py          # Marks the tests directory as a Python package
│   ├── test_utils.py        # Unit tests for utility functions
│   ├── test_config_manager.py # Unit tests for ConfigManager
│   └── test_pole_processor.py # Unit tests for PoleDataProcessor
├── requirements.txt         # Project dependencies
├── setup.py                 # Packaging information
├── .gitignore               # Files to ignore in version control
├── .vscode/                 # Visual Studio Code settings
│   ├── settings.json        # VS Code specific settings
│   ├── launch.json          # Debugging configurations
│   └── tasks.json           # Task configurations
└── README.md                # Project documentation
```

## Installation

To set up the project, clone the repository and install the required dependencies:

```bash
git clone <repository-url>
cd mr-ss-pole-mapper
pip install -r requirements.txt
```

## Usage

Run the application using the following command:

```bash
python src/main.py
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
