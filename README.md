# Bing New Functions Error Corrector

This project is an **interactive Windows conflict scanner and fixer**, 
originally intended for detecting potential conflicts with the **new Microsoft Edge search menu shortcut** (`Win + Shift + F`), 
but extendable to other future features.

## Features
- **Multi-language support** (10 languages)
- **Automatic dependency installation** (`psutil`, `reportlab` for PDF)
- **Conflict scanning**:
  - Running processes
  - Startup programs
  - HKCU registry startup values
  - Windows services
- **Interactive remediation** (terminate processes, remove registry entries, disable services)
- **Report generation** in 10 formats: `.txt`, `.json`, `.csv`, `.xml`, `.html`, `.md`, `.log`, `.yml`, `.ini`, `.pdf`

## Requirements
- Windows with PowerShell available
- Python 3.8+
- Internet connection for auto-installing missing dependencies

## Usage
```bash
python ErrorBroker.py
```
Follow the prompts to select a language, scan the system, and optionally perform interactive fixes.

## Repository Structure
```
Bing-new-functions-error-corrector/
│   ErrorBroker.py   # Main script
│   README.md        # This documentation
```

---
**Author**: ToraScriptCopy
**License**: MIT
