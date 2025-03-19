@echo off
REM Mudar para o diretório do projeto
cd Moraca

REM Instalar todas as dependências necessárias
pip install -r ..\requirements.txt

REM Criar o executável para Windows
pyinstaller --name=Moraca --onefile --windowed --icon=assets\logo.ico --add-data "assets;assets" main.py

echo Executável criado com sucesso em dist\Moraca.exe 