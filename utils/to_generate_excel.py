# utils/to_generate_excel.py
"""
Módulo para la generación y guardado de un archivo Excel a partir de un DataFrame.
"""

import os
import pandas as pd
from typing import Optional

def generar_excel(
    data: pd.DataFrame,
    output: Optional[str] = None
) -> None:
    """
    Genera un archivo Excel a partir de un DataFrame con los datos provistos.

    :param data: DataFrame con la información a guardar.
    :param output: Ruta completa del archivo de salida. Si no se especifica, se 
                   guarda un archivo 'data_download.xlsx' en el directorio actual.
    :return: None
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío. No se generará el archivo Excel.")
        return

    # Si el índice tiene información de zona horaria (UTC), la eliminamos.
    if data.index.dtype == "datetime64[ns, UTC]":
        data.index = data.index.tz_convert(None)

    # Determinar el nombre de archivo de salida
    file_name = output if output else "data_download.xlsx"

    # Escribir a archivo Excel
    try:
        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            data.to_excel(writer, sheet_name="data")
    except Exception as e:
        print(f"Ocurrió un error al guardar el archivo Excel: {e}")
        return

    # Obtener la ruta absoluta del archivo creado
    file_path = os.path.abspath(file_name)
    print(f"El archivo ha sido guardado como '{file_name}' en '{file_path}'")
