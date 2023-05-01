import subprocess
import win32com.shell.shell as shell

# Funcion que permite al usuario elegir entre utilizar su propio servidor o uno publico
def ip_choice():
        print('''
        Quieres utilizar tu propio servidor KMS para activar el producto o utilizar el mio?

        1) Para utilizar tu propio servidor
        2) Para utilizar mi servidor
        
        ''')

        while True:
            user_input = input("Elige una opción: ")
            if user_input == "1":
                ip = input("Introduce tu ip: ")
                return ip
            elif user_input == "2":
                ip = "dhanjal.ddns.net"
                return ip
            else:
                print("Opción no valida! Por favor elige una opción entre el 1 o el 2")
                continue
ip = ip_choice()

'''---------------------------------------------------------------------------------------------------------------------------------'''
def activate_windows():
    # Diccionario de las claves
    keys = {"Windows 11 Pro" : "W269N-WFGWX-YVC9B-4J6C9-T83GX", "Windows 11 Home": "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", "Windows 10 Pro" : "W269N-WFGWX-YVC9B-4J6C9-T83GX", "Windows 10 Home": "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"}


    # Obtener la edición de Windows
    def obtener_edicion():
        print("Obteniendo la edición de tu Windows automaticamente............")
        output = subprocess.check_output('wmic os get caption', shell=True)
        output = output.decode('utf-8').strip()
        edition = output.split('\n')[1]
        # Eliminar el texto "Microsoft" de la edición
        edition = edition.replace('Microsoft', '').strip()
        print(f"La edición de tu Windows es: {edition}")
        return edition
        
    def ejecutar_comandos():
        print("Activando Windows............")
        commands = f'cscript //nologo slmgr.vbs /ipk {edition_key} & cscript //nologo slmgr.vbs /skms {ip} & cscript //nologo slmgr.vbs /ato'
        shell.ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+commands)
        print("Windows Activado!")

    
    edition_key = keys[obtener_edicion()]
    ejecutar_comandos()
'''---------------------------------------------------------------------------------------------------------------------------------'''
def activate_office():
    print("Activando Office 2019+............")
    commands = f'cd "C:\Program Files\Microsoft Office\Office16" & cscript //nologo slmgr.vbs /skms {ip} & cscript //nologo slmgr.vbs /ato'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+commands)
    print("Windows Activado!")



'''---------------------------------------------------------------------------------------------------------------------------------'''
def elegir_producto():
    print('''
    Que producto quieres activar?

    1. Activar Windows
    2. Activar Office 2019+
    ''')
    while True:
            user_input = input("Elige una opción: ")
            if user_input == "1":
                 print("Has elegido activar Windows")
                 activate_windows()
            if user_input == "2":
                 print("Has elegido activar Office")
                 activate_office()
            break
            

    
elegir_producto()

