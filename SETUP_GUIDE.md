# MR SS Pole Mapper - Setup Guide

## Quick Start After GitHub Download

If you downloaded this project from GitHub and are getting empty output files, follow these steps:

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Template Files
The application requires template files to generate output. Template files are now included in the `templates/` directory.

**✅ Template files are already included!** The `Consumer SS Template.xltm` file is in the `templates/` folder and the application will automatically find it.

### 3. Run the Application
```bash
python src/main.py
```

### 4. Configuration
- Select your input files in the GUI
- The application will automatically use the template from the `templates/` directory
- Process your data normally

## Troubleshooting

### Empty Output Issue
If you're still getting empty output files:

1. **Check the logs** - The application shows detailed logging in the GUI
2. **Verify template location** - Make sure `templates/Consumer SS Template.xltm` exists
3. **Check file permissions** - Ensure the application can read/write files in your chosen directories

### Template File Locations
The application looks for templates in this order:
1. `templates/Consumer SS Template.xltm` (recommended - included in repo)
2. `Consumer SS Template.xltm` (in current directory)
3. `template.xlsx`
4. `template.xltm`

## What Was Fixed

Previously, the application had a hardcoded path that only worked on the developer's machine:
- ❌ `'C:/Users/nsaro/Desktop/Test/Consumer SS Template.xltm'`
- ✅ `'templates/Consumer SS Template.xltm'`

Now the template file is included in the repository and the path is relative, so it works for everyone who downloads from GitHub.

## Need Help?

If you still encounter issues:
1. Check that all required files are in the `templates/` directory
2. Verify your Python environment has all dependencies installed
3. Run the application and check the log output for specific error messages 