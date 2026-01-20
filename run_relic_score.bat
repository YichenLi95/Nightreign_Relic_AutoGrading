@echo off
cd /d "%~dp0"
python export_relic_scores.py RelicScores.xlsx --all
pause