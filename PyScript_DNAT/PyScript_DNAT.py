import argparse
import pandas as pd

def parse_vip_config(input_file):
    with open(input_file, 'r') as file:
        lines = file.readlines()

    vip_data = []
    current_vip = {}
    id_counter = 1

    for line in lines:
        line = line.strip()
        
        if line.startswith("edit"):
            if current_vip:  # Si current_vip no está vacío, guárdalo
                # Si no se encontró un protocolo, asignar TCP por defecto
                if 'Protocolo' not in current_vip:
                    current_vip['Protocolo'] = 'TCP'
                vip_data.append(current_vip)
                current_vip = {}
            current_vip['ID'] = id_counter
            id_counter += 1
            current_vip['Nombre'] = line.split('"')[1]
        
        elif line.startswith("set extip"):
            current_vip['IP Externa'] = line.split(" ")[-1]
        
        elif line.startswith("set mappedip"):
            current_vip['IP Interna'] = line.split(" ")[-1].replace('"', '')  # Eliminar comillas
        
        elif line.startswith("set portforward"):
            current_vip['Puerto Externo'] = ""
            current_vip['Puerto Interno'] = ""
            current_vip['Portforward'] = "disable"  # Por defecto en deshabilitar
            if "enable" in line:
                current_vip['Portforward'] = "enable"
        
        elif line.startswith("set extport"):
            if current_vip.get('Portforward') == "enable":
                current_vip['Puerto Externo'] = line.split(" ")[-1]
        
        elif line.startswith("set mappedport"):
            if current_vip.get('Portforward') == "enable":
                current_vip['Puerto Interno'] = line.split(" ")[-1]
        
        elif line.startswith("set protocol"):  # Capturando el protocolo
            current_vip['Protocolo'] = line.split(" ")[-1].upper()  # Convertir a mayúsculas

    if current_vip:  # Agregar la última entrada si existe
        # Si no se encontró un protocolo, asignar TCP por defecto
        if 'Protocolo' not in current_vip:
            current_vip['Protocolo'] = 'TCP'
        vip_data.append(current_vip)

    # Eliminar la clave 'Portforward' de cada diccionario
    for vip in vip_data:
        vip.pop('Portforward', None)

    return vip_data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate an Excel file from Fortigate VIP config.')
    parser.add_argument('-i', '--input', required=True, help='Input file containing the Fortigate VIP configuration.')
    parser.add_argument('-o', '--output', default='dnat_report.xlsx', help='Output Excel file name.')  # Cambiado a dnat_report.xlsx

    args = parser.parse_args()

    vip_info = parse_vip_config(args.input)
    save_to_excel(vip_info, args.output)

    print(f'Data has been written to {args.output}')
