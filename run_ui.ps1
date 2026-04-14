# Simple desktop window for the shopping-route tool (plain language for new staff).
# Usage: .\run_ui.ps1
# Install deps once: pip install -r requirements.txt
#   (Charms tab drag-and-drop on Windows uses the optional "windnd" package.)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
python src/simple_ui.py
