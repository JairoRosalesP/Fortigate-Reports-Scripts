import argparse
import openpyxl
import re

# Función para analizar el archivo input.txt (categorías)
def parse_categories_file(categories_file):
    categories = {}
    current_group = None

    with open(categories_file, 'r') as file:
        for line in file:
            line = line.strip()
            
            # Si la línea comienza con 'g', es un grupo
            group_match = re.match(r'g(\d+)\s+(.*)', line)
            if group_match:
                group_id = group_match.group(1)
                group_name = group_match.group(2)
                current_group = group_name
                categories[group_id] = ('Group', current_group)
            
            # Si no comienza con 'g', es una categoría
            elif current_group is not None:
                category_match = re.match(r'(\d+)\s+(.*)', line)
                if category_match:
                    category_id = category_match.group(1)
                    category_name = category_match.group(2)
                    categories[category_id] = ('Category', category_name)

    return categories

# Función para analizar el archivo input2.txt (perfiles webfilter)
def parse_webfilter_profiles(profiles_file):
    profiles = {}
    current_profile = None

    with open(profiles_file, 'r') as file:
        for line in file:
            line = line.strip()

            # Buscar el nombre del perfil con 'edit "name"'
            profile_match = re.match(r'edit\s+"(.+)"', line)
            if profile_match:
                current_profile = profile_match.group(1)
                profiles[current_profile] = []

            # Buscar las acciones 'set category' y 'set action'
            elif current_profile is not None:
                category_match = re.match(r'set category\s+(\d+)', line)
                action_match = re.match(r'set action\s+(block|permit|warning)', line)

                if category_match:
                    category_id = category_match.group(1)
                    profiles[current_profile].append((category_id, 'permit'))  # Acción por defecto

                if action_match:
                    action = action_match.group(1)
                    # Reemplazar la acción en la última categoría si está presente
                    if profiles[current_profile]:
                        category_id, _ = profiles[current_profile][-1]
                        profiles[current_profile][-1] = (category_id, action)

    return profiles

# Función para generar el archivo de salida web_filter_report.xlsx
def generate_webfilter_report(categories_file, profiles_file):
    categories = parse_categories_file(categories_file)
    profiles = parse_webfilter_profiles(profiles_file)

    # Crear el archivo Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Web Filter Report'

    # Escribir encabezados de columna
    ws.append(['TYPE', 'NAME'] + list(profiles.keys()))

    # Verificar si la categoría "Unrated" (ID 0) existe, si no, añadirla
    unrated_category = ('0', ('Category', 'Unrated'))
    if '0' not in categories:
        categories['0'] = unrated_category[1]

    # Escribir las categorías y grupos en las filas
    for category_id, (category_type, category_name) in categories.items():
        row = [category_type, category_name]

        # Añadir columnas para cada perfil
        for profile_name in profiles:
            if category_type == 'Group':
                row.append('')  # Dejar en blanco para los grupos
            else:
                # Buscar la categoría en el perfil y definir la acción
                action = 'warning' if category_id == '0' else 'permit'  # Acción por defecto, warning para Unrated
                for cat_id, cat_action in profiles[profile_name]:
                    if cat_id == category_id:
                        action = cat_action
                        break
                row.append(action)

        # Añadir la fila de categoría o grupo al archivo Excel
        ws.append(row)

    # Guardar el archivo Excel con el nombre web_filter_report.xlsx
    wb.save('web_filter_report.xlsx')

# Función principal
def main():
    parser = argparse.ArgumentParser(description='Generate Web Filter report from input files.')
    parser.add_argument('-i', '--input1', required=True, help='Input file with categories (get webfilter categories)')
    parser.add_argument('-g', '--input2', required=True, help='Input file with profiles (show webfilter profile)')
    args = parser.parse_args()

    generate_webfilter_report(args.input1, args.input2)

if __name__ == '__main__':
    main()
