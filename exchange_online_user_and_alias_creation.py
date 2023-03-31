import importlib
import subprocess
import sys
import csv
import msal
from exchangelib import Credentials, Account, DELEGATE

# Función para comprobar e instalar módulos necesarios
def install_module(module_name):
    try:
        importlib.import_module(module_name)
    except ImportError:
        print(f"'{module_name}' not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])

# Comprobar e instalar módulos necesarios
install_module("exchangelib")
install_module("msal")
install_module("csv")
install_module("subprocess")

# Comprobar si se han proporcionado los argumentos necesarios
if len(sys.argv) < 3:
    print("Uso: python script.py <admin_email> <password> <user_csv_file> <alias_csv_file>")
    sys.exit(1)

# Recoger argumentos
admin_email = sys.argv[1]
password = sys.argv[2]
user_csv_file = sys.argv[3]
alias_csv_file = sys.argv[4]

# Autenticación con Microsoft Graph API
app = msal.PublicClientApplication("your_client_id")
result = None
accounts = app.get_accounts(username=admin_email)
if accounts:
    result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=accounts[0])
if not result:
    result = app.acquire_token_by_username_password(admin_email, password, scopes=["https://graph.microsoft.com/.default"])
if "access_token" in result:
    access_token = result["access_token"]
else:
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))

# Conexión con Exchange Online
credentials = Credentials(access_token)
account = Account(primary_smtp_address=admin_email, credentials=credentials, autodiscover=True, access_type=DELEGATE)

# Leer usuarios a crear del archivo CSV
with open("users.csv", "r") as f:
    reader = csv.DictReader(f)
    for row in reader:
        # Crear usuario
        new_user = account.protocol.create_user(
            row["username"],
            password=password,
            first_name=row["first_name"],
            last_name=row["last_name"],
            display_name=row["display_name"],
            department=row["department"],
            job_title=row["job_title"],
            company=row["company"],
            street=row["street"],
            city=row["city"],
            state=row["state"],
            country=row["country"],
            postal_code=row["postal_code"],
            phone=row["phone"],
            fax=row["fax"],
            mobile=row["mobile"],
            home_phone=row["home_phone"],
            other_phone=row["other_phone"],
            initials=row["initials"],
            assistant_name=row["assistant_name"],
            manager=row["manager"],
            mailbox_size=row["mailbox_size"],
            user_principal_name=row["user_principal_name"]
        )
        print(f"Usuarios creados correctamente!")

# Leer alias a crear del archivo CSV
with open("alias_csv_file", "r") as f:
    reader = csv.DictReader(f)
    for row in reader:
        # Crear alias
        new_alias = account.protocol.create_alias(new_user, row["alias"])

        print(f"Alias creados correctamente!")

# Cerrar la sesión de Exchange Online
account.protocol.close()
