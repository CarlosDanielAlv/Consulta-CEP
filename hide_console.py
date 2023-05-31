import subprocess

# Nome do arquivo principal do seu código Python
main_file = 'main.py'

# Comando para criar o executável usando o PyInstaller
subprocess.call(['pyinstaller', '--noconsole', main_file])