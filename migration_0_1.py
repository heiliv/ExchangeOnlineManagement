import os
import csv
from office365.graph_client import GraphClient
from office365.directory.directory_object_collection import DirectoryObjectCollection
from office365.directory.user import User

# Instalar la biblioteca office365-python-client
def install_module(module_name):
    try:
        importlib.import_module(module_name)
        print(f"{module_name} ya está instalado")
    except ImportError:
        print(f"Instalando {module_name}...")
        os.system(f"python -m pip install {module_name}")

# Conectar a Exchange Online
client = GraphClient("https://graph.microsoft.com/v1.0", None)
client.connect()

# Autenticarse
username = input("Introduce tu nombre de usuario: ")
password = input("Introduce tu contraseña: ")
client.authenticate(username, password)

# Dar opcion de elegir qué accion realizar
opcion = input("¿Qué acción desea realizar? (1: Migrar usuarios, 2: Crear alias, 3: Eliminar usuarios, 4: Asignar licencias)")

# Realizar accion seleccionada
if opcion == "1":
    # Migrar usuarios
    with open("C:/csvUsuarios/usuarios.csv") as file:
        reader = csv.DictReader(file)
        for row in reader:
            user = User(client.context)
            user.display_name = row["DisplayName"]
            user.user_principal_name = row["UserPrincipalName"]
            user.set_password(row["Password"])
            user.account_enabled = True
            user.strong_password_required = True
            user.force_change_password_next_sign_in = True
            user.mail_nickname = row["UserPrincipalName"].split("@")[0]
            user.mailbox_settings = {
                "language": "es-ES",
                "time_zone": "Romance Standard Time",
                "automatic_replies_setting": "disabled",
                "storage_quota": 2147483648,
                "issue_warning_quota": 2097152000
            }
            user.address = {
                "street": "",
                "city": "",
                "state": "",
                "countryOrRegion": "",
                "postalCode": ""
            }
            user.save()

elif opcion == "2":
    # Crear alias
    with open("C:/csvUsuarios/alias.csv") as file:
        reader = csv.DictReader(file)
        for row in reader:
            user = DirectoryObjectCollection(client.context).get_by_id(row["User"])
            user.alias = [row["Alias"]]
            user.update()

elif opcion == "3":
    # Eliminar usuarios
    with open("C:/csvUsuarios/usuarios.csv") as file:
        reader = csv.DictReader(file)
        for row in reader:
            user = User(client.context).get_by_email(row["UserPrincipalName"])
            user.delete_object()

elif opcion == "4":
    # Asignar licencias
    with open("C:/csvUsuarios/usuarios.csv") as file:
        reader = csv.DictReader(file)
        for row in reader:
            user = User(client.context).get_by_email(row["UserPrincipalName"])
            user.assign_license(["reseller-account:O365_BUSINESS_PREMIUM"])

else:
    print("Opción inválida, saliendo del script.")
