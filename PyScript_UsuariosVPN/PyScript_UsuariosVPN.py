import argparse
import re
import openpyxl

# Función para procesar el archivo de usuarios locales
def process_local_users(local_users_file):
    users = []
    user_data = {}
    
    with open(local_users_file, 'r') as file:
        for line in file:
            line = line.strip()
            
            # Detecta el nombre del usuario
            name_match = re.match(r'edit\s+"(.+)"', line)
            if name_match:
                if user_data:
                    users.append(user_data)
                user_data = {
                    "NAME": name_match.group(1),
                    "TYPE": "LOCAL",
                    "MFA": "NO",
                    "MFA TYPE": "",
                    "GROUP": "",
                    "STATUS": "enable"  # Por defecto, el status es "enable"
                }
            
            # Detecta el tipo de usuario
            if 'set type ldap' in line:
                user_data["TYPE"] = "LDAP"
            elif 'set type password' in line:
                user_data["TYPE"] = "LOCAL"

            # Detecta si MFA está habilitado
            if 'set two-factor' in line:
                user_data["MFA"] = "SI"
                if 'email' in line:
                    email_match = re.search(r'set email-to\s+(.+)', file.readline().strip())
                    if email_match:
                        user_data["MFA TYPE"] = f"EMAIL: {email_match.group(1)}"
                elif 'sms' in line:
                    sms_match = re.search(r'set sms-phone\s+(.+)', file.readline().strip())
                    if sms_match:
                        user_data["MFA TYPE"] = f"SMS: {sms_match.group(1)}"
                elif 'fortitoken' in line:
                    user_data["MFA TYPE"] = "FORTITOKEN"
                else:
                    # Otros métodos MFA
                    method_match = re.search(r'set two-factor\s+(.+)', line)
                    if method_match:
                        user_data["MFA TYPE"] = method_match.group(1).upper()

            # Detecta si el estado del usuario es "disable"
            if 'set status disable' in line:
                user_data["STATUS"] = "disable"

        # Agregar el último usuario
        if user_data:
            users.append(user_data)
    
    return users

# Función para procesar el archivo de grupos de usuarios
def process_user_groups(user_groups_file):
    groups = {}
    group_name = None

    with open(user_groups_file, 'r') as file:
        for line in file:
            line = line.strip()

            # Detecta el nombre del grupo
            group_match = re.match(r'edit\s+"(.+)"', line)
            if group_match:
                group_name = group_match.group(1)
                groups[group_name] = []

            # Detecta los miembros del grupo
            member_match = re.search(r'set member\s+(.+)', line)
            if member_match:
                members = re.findall(r'"([^"]+)"', member_match.group(1))
                groups[group_name].extend(members)

    return groups

# Función para asignar los grupos a los usuarios
def assign_groups_to_users(users, groups):
    for user in users:
        for group, members in groups.items():
            if user["NAME"] in members:
                user["GROUP"] = group
                break

# Función para generar el archivo Excel
def generate_excel_report(users, output_file='vpn_users_report.xlsx'):
    # Ordenar los usuarios por el nombre (NAME) en orden alfabético
    users.sort(key=lambda x: x["NAME"].lower())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VPN Users"

    # Encabezados
    headers = ["ID", "NAME", "TYPE", "MFA", "MFA TYPE", "GROUP", "STATUS"]  # Se agrega "STATUS"
    ws.append(headers)

    # Datos de los usuarios
    for idx, user in enumerate(users, start=1):
        ws.append([idx, user["NAME"], user["TYPE"], user["MFA"], user["MFA TYPE"], user["GROUP"], user["STATUS"]])

    # Guardar el archivo Excel
    wb.save(output_file)
    print(f"Reporte generado: {output_file}")

# Función principal
def main():
    parser = argparse.ArgumentParser(description="Generar reporte de usuarios VPN")
    parser.add_argument('-i', '--input_users', required=True, help='Archivo de usuarios locales')
    parser.add_argument('-g', '--input_groups', required=True, help='Archivo de grupos de usuarios')
    args = parser.parse_args()

    # Procesar los inputs
    users = process_local_users(args.input_users)
    groups = process_user_groups(args.input_groups)
    assign_groups_to_users(users, groups)

    # Generar el reporte en Excel
    generate_excel_report(users)

if __name__ == "__main__":
    main()
