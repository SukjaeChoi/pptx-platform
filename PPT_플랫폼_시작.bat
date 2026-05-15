@echo off
cd /d "%~dp0"
start "PPTX" node server.js
start "" http://localhost:3100