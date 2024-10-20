#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

from os import path 
import io
import sys
import re
import pandas as pd  # Importamos pandas para exportar a Excel
import os

# Definición de opciones
from optparse import OptionParser
from optparse import OptionGroup

# Configuración del parser para opciones de línea de comandos
parser = OptionParser(usage="%prog [options]")

main_grp = OptionGroup(parser, 'Parámetros principales')
main_grp.add_option('-i', '--input-file', help='Archivo de configuración de Fortigate. Ej: fgfw.cfg')
main_grp.add_option('-o', '--output-file', help='Archivo de salida de Excel (default ./firewall_policy_report.xlsx)', 
                    default=path.abspath(path.join(os.getcwd(), './firewall_policy_report.xlsx')))
main_grp.add_option('-s', '--skip-header', help='No imprimir el encabezado de Excel', action='store_true', default=False)
main_grp.add_option('-n', '--newline', help='Insertar un salto de línea entre cada política para mejor legibilidad', action='store_true', default=False)
main_grp.add_option('-e', '--input-encoding', help='Codificación del archivo de entrada (default "utf-8")', default='utf-8')
parser.option_groups.extend([main_grp])

# Compatibilidad con Python 2 y 3
if (sys.version_info < (3, 0)):
    fd_read_options = 'r'
else:
    fd_read_options = 'r'

# Patrones de expresiones regulares para identificar bloques de políticas
p_entering_policy_block = re.compile(r'^\s*config firewall policy$', re.IGNORECASE)
p_exiting_policy_block = re.compile(r'^end$', re.IGNORECASE)
p_policy_next = re.compile(r'^next$', re.IGNORECASE)
p_policy_number = re.compile(r'^\s*edit\s+(?P<policy_number>\d+)', re.IGNORECASE)
p_policy_set = re.compile(r'^\s*set\s+(?P<policy_key>\S+)\s+(?P<policy_value>.*)$', re.IGNORECASE)

# Función para analizar el archivo de configuración
def parse(options):
    """
        Analiza los datos según varias expresiones regulares
        
        @param options:  opciones
        @rtype: devuelve una lista de políticas y la lista de claves únicas vistas
    """
    global p_entering_policy_block, p_exiting_policy_block, p_policy_next, p_policy_number, p_policy_set
    
    in_policy_block = False
    skip_ssl_vpn_policy_block = False
    policy_list = []
    policy_elem = {}
    
    order_keys = []
    
    with io.open(options.input_file, mode=fd_read_options, encoding=options.input_encoding) as fd_input:
        for line in fd_input:
            line = line.strip()
            
            # Entrando en un bloque de política
            if p_entering_policy_block.search(line):
                in_policy_block = True
            
            # En un bloque de política
            if in_policy_block:
                if p_policy_number.search(line) and not(skip_ssl_vpn_policy_block):
                    policy_number = p_policy_number.search(line).group('policy_number')
                    policy_elem[u'id'] = policy_number
                    if not('id' in order_keys):
                        order_keys.append(u'id')
                
                # Coincidiendo con una configuración
                if p_policy_set.search(line) and not(skip_ssl_vpn_policy_block):
                    policy_key = p_policy_set.search(line).group('policy_key')
                    if not(policy_key in order_keys):
                        order_keys.append(policy_key)
                    
                    policy_value = p_policy_set.search(line).group('policy_value').strip()
                    policy_value = re.sub('["]', '', policy_value)
                    
                    policy_elem[policy_key] = policy_value
                
                # Finalizando el bloque de política actual
                if p_policy_next.search(line) and not(skip_ssl_vpn_policy_block):
                    policy_list.append(policy_elem)
                    policy_elem = {}
                    
            # Saliendo del bloque de política
            if p_exiting_policy_block.search(line):
                in_policy_block = False
    
    return (policy_list, order_keys)

# Función para generar el archivo de Excel
def generate_excel(results, keys, options):
    """
        Genera un archivo de Excel
    """
    if results and keys:
        # Convertimos la lista de políticas a un DataFrame de pandas
        df = pd.DataFrame(results, columns=keys)
        
        # Guardamos el DataFrame en un archivo de Excel
        df.to_excel(options.output_file, index=False, header=not options.skip_header)
    
    return None

def main():
    """
        Función principal
    """
    global parser
    
    options, arguments = parser.parse_args()
    
    if (options.input_file is None):
        parser.error('Por favor, especifique un archivo de entrada válido')
    
    results, keys = parse(options)
    generate_excel(results, keys, options)  # Generamos el archivo de Excel
    
    return None

if __name__ == "__main__":
    main()
