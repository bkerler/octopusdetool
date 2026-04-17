@echo off
setlocal

rd /S /Q build
rd /S /Q dist
python -m pip install briefcase
python -m briefcase create windows
python -m briefcase build windows
python -m briefcase package windows -i 8b5f17c056bf5eb66151e95b86831d32cd037dfe
