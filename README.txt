--------------------------------------------------------------------------------
Comando generazione file exe:
pyinstaller --onefile main.py --name "Generazione report excel" --noconsole
--------------------------------------------------------------------------------

--------------------------------------------------------------------------------
Modificando il file .spec il comando è:
pyinstaller "Generazione report excel.spec"
--------------------------------------------------------------------------------

--------------------------------------------------------------------------------
Creazione file dmg da windows
1. Installare WSL
2. Installare distro linux ( wsl --install Ubuntu-18.04 )
3. Entrare in ubuntu ed eseguire i seguenti comandi:
    1. sudo apt update
    2. sudo apt install python3.8-pip
    3. python3.8 -m pip install dmgbuild