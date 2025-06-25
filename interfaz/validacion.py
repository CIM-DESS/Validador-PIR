import pandas as pd
import os
import re

from django.core.files.uploadedfile import UploadedFile
from typing import List, Dict
from dataclasses import dataclass
from datetime import datetime

@dataclass
class Novedad:
    archivo: str
    hoja: str
    fila: int
    descripcion: str

def procesar_archivos_excel(archivos: List[UploadedFile]) -> List[Dict[str, str]]:
    novedades = []
    codigos_vistos = {}
    trafos_expansion = {}

    trafos_excluidos = cargar_datos_pivote("trafos_excluidos", 0)   # ✅ Cargar transformadores excluidos desde el archivo pivote (hoja, columna)
    uc_validas = cargar_datos_pivote("uc", 3)   # ✅ lee la columna "Código UC" de la hoja 'uc' (hoja, columna)
    lineas_validas = cargar_datos_pivote("lineas", 0)  # 🟩 Columna 0 = 'Código línea'
    municipios_validos = cargar_datos_pivote("municipio", 1)  # Columna 1 = segunda columna
    regionales_validos = cargar_datos_pivote("regionales", 1)  # 🟡 Columna 'Cod Dane'
    cod_dane_validos = cargar_datos_pivote("municipio", 2)  # 🟡 Columna 'CODIGO'
    ius_validos = cargar_datos_pivote("subestaciones", 2)  # 🟡 Columna 'IUS'

    # # ✅ Diccionario desde 'equipo_proteccion' (usamos varias columnas)
    df_equipo_proteccion = pd.read_excel(RUTA_PIVOTE, sheet_name="equipo_proteccion", dtype=str)

    # Creamos un diccionario por FID con las columnas necesarias
    dicc_equipo_proteccion = {}
    for _, fila in df_equipo_proteccion.iterrows():
        fid = str(fila.iloc[1]).strip().upper()  # FID en columna 2 (índice 1)
        if fid != "" and fid != "NAN":
            dicc_equipo_proteccion[fid] = {
                "codigo_uc": str(fila.iloc[6]).strip().upper(),  # Columna 7
                "anio_operacion": str(fila.iloc[8]).strip()       # Columna 9
            }


    for archivo in archivos:
        nombre_archivo = archivo.name

        try:
            xls = pd.ExcelFile(archivo)

            for hoja in xls.sheet_names:
                if hoja.lower() in ["listas", "plan presentado 2025"]:
                    continue


                df = xls.parse(hoja, dtype=str)
                print("➡️ Analizando hoja:", hoja)

                for index, fila in df.iterrows():
                    num_fila = index + 2  # Ajuste por encabezado de Excel

                    # === 🧾 EXTRAER VALORES COMUNES PARA TODAS LAS VALIDACIONES ===
                    codigo_fid = str(fila.get("Código FID", "")).strip()
                    nombre_proyecto = fila.get("Nombre del proyecto", "")
                    tipo_inversion = str(fila.get("Tipo inversión", "")).strip().upper()
                    uc = str(fila.get("Unidad Constructiva", "")).strip().upper()
                    uc_rep = str(fila.get("Codigo UC_rep", "")).strip().upper()
                    rpp = fila.get("RPP", None)
                    transformador = str(fila.iloc[6]).strip().upper()  # Si está en la quinta columna (índice 4)
                    descripcion_rep = fila.get("DESCRIPCION_rep", None)
                    cantidad_rep = fila.get("Cantidad_rep", None)
                    codigo_fid_rep = fila.get("Código FID_rep", None)
                    numero_conductores_rep = fila.get("Número de conductores_rep", None)
                    año_entrada_rep = fila.get("Año entrada operación_rep", None)
                    año_actual = datetime.now().year
                    rpp_rep = fila.get("Rpp_rep", None)
                    anio_entrada_op = fila.get("Año entrada operación", "")
                    anio_salida_op = fila.get("Año salida operación", "")
                    anio_entrada_rep = fila.get("Año entrada operación_rep", None)
                    codigo_linea = str(fila.get("Código línea", "")).strip().upper()
                    municipio = str(fila.get("Municipio", "")).strip().upper()
                    cantidad = fila.get("Cantidad", None)
                    nivel = fila.get("Nivel", "")
                    nombre = str(fila.get("Nombre", "")).strip()
                    contrato = str(fila.get("Contrato/Soporte", "")).strip()
                    nombre_archivo_sin_extension = os.path.splitext(nombre_archivo)[0]
                    nombre_plantilla = str(fila.get("Nombre de la Plantilla", "")).strip()
                    sobrepuesto = str(fila.get("Sobrepuesto", "")).strip().upper()
                    numero_conductores = fila.get("Número de conductores", None)
                    tension_operacion = fila.get("Tensión de operación", None)
                    descripcion = fila.get("DESCRIPCION", None)
                    codigo_proyecto = fila.get("Código proyecto", None)
                    cod_regional = str(fila.get("Cod Regional", "")).strip().upper()
                    cod_dane_municipio = str(fila.get("Cod DANE municipio", "")).strip().upper()
                    ius = fila.get("IUS", None)
                    fraccion_costo = str(fila.get("Fracción costo", "")).strip()
                    porcentaje_uso = fila.get("Porcentaje uso", None)
                    iul = str(fila.get("IUL", "")).strip()
                    ius_final = fila.get("IUS Final", None)
                    ius_inicial = fila.get("IUS inicial", None)
                    piec = str(fila.get("PIEC", "")).strip().upper()
                    str_construccion = str(fila.get("STR construcción", "")).strip().upper()
                    tipo_proyecto = fila.get("Tipo de Proyecto", None)
                    activo_construido = fila.get("Activo Construido y en Operación", None)

                    # === 🟡 VALIDACIÓN 1: Código FID duplicado ===
                    if pd.notna(codigo_fid) and str(codigo_fid).strip().upper() not in ["", "NAN"]:
                        clave_fid = str(codigo_fid).strip().upper()
                        if clave_fid in codigos_vistos:
                            info = codigos_vistos[clave_fid]
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"El código FID '{codigo_fid}' se repite en el ARCHIVO: {info['archivo']} - HOJA: {info['hoja']} - FILA: {info['fila']}")
                            ))
                        else:
                            codigos_vistos[clave_fid] = {
                                "archivo": nombre_archivo,
                                "hoja": hoja,
                                "fila": num_fila
                            }

                    # === 🟡 VALIDACIÓN 2: Duplicidad de Transformadores de Expansión ===
                    if hoja == "Transformador_Distribucion" and tipo_inversion in {"II", "IV"}:
                        if transformador in trafos_expansion:
                            info = trafos_expansion[transformador]
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"El Transformador '{transformador}' con tipo de inversión '{tipo_inversion}' "
                                            f"ya fue reportado como expansión en el ARCHIVO: {info['archivo']} - HOJA: {info['hoja']} - FILA: {info['fila']}")
                            ))
                        else:
                            # Registrar el transformador en el diccionario
                            trafos_expansion[transformador] = {
                                "archivo": nombre_archivo,
                                "hoja": hoja,
                                "fila": num_fila
                            }

                    # === 🟡 VALIDACIÓN 3: Campo obligatorio 'Nombre del proyecto' ===
                    if nombre_proyecto is None or pd.isna(nombre_proyecto) or str(nombre_proyecto).strip() == "":
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="La columna 'Nombre del proyecto' está vacía."
                        ))

                    # === 🟡 VALIDACIÓN 4: Tipo de inversión válido ===
                    if tipo_inversion not in {"I", "II", "III", "IV"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Tipo inversión' no es válido. Debe ser 'I', 'II', 'III' o 'IV'. Valor recibido: {}".format(tipo_inversion)
                        ))
                        continue  # No tiene sentido seguir validando si esto está mal

                    # === 🔵 VALIDACIÓN 5: VALIDACIONES ESPECÍFICAS PARA I y III (Reposición) ===
                    if tipo_inversion in {"I", "III"}:

                        # 5.1 Campo 'Codigo UC_rep' obligatorio
                        if uc_rep in ["", "nan", "NAN"]:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Codigo UC_rep' no debe estar vacío cuando el tipo de inversión es '{tipo_inversion}'."
                            ))

                        # 5.2 Coincidencia entre tipo de UC y UC_rep (primeros 3 caracteres)
                        if uc not in {"DESMANTELADO", "MONTAJE INTEGRAL"}:
                            if uc[:3] != uc_rep[:3]:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"El Código UC_rep '{uc_rep}' no es del mismo tipo que la UC de ingreso '{uc}'."
                                ))

                        # 5.3 Verificar que 'Codigo UC_rep' esté listado en la hoja 'uc' del archivo pivote
                        if uc_rep and uc_rep not in uc_validas:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"El Código UC_rep '{uc_rep}' no se encuentra en el archivo pivote en la columna 'Código UC' de la hoja 'uc'."
                            ))

                        # 5.4 Validar que 'DESCRIPCION_rep' no esté vacío
                        if descripcion_rep is None or pd.isna(descripcion_rep) or str(descripcion_rep).strip() == "" or str(descripcion_rep).strip() == "nan" or str(descripcion_rep).strip() == "NAN" or str(descripcion_rep).strip() == "Identifique UC_rep":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="El campo 'DESCRIPCION_rep' no puede estar vacío o categorizado como 'Identifique UC_rep'."
                            ))
                        
                        # 5.5 Validar que 'Cantidad_rep' sea mayor a cero
                        try:
                            cantidad_valor = float(str(cantidad_rep).strip())

                            if cantidad_valor <= 0:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"El campo 'Cantidad_rep' debe ser mayor a cero. El valor recibido es: {cantidad_rep}"
                                ))

                            # Validar valor exacto = 1 solo si es hoja Transformador_Distribucion
                            if hoja.strip().lower() == "transformador_distribucion" and cantidad_valor != 1:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"Error en 'Cantidad_rep' ya que se esperaba un '1' y el valor recibido es: {cantidad_rep}"
                                ))

                        except (ValueError, TypeError):
                            # Si no es numérico o está vacío
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"El campo 'Cantidad_rep' debe ser numérico y mayor a cero. Valor recibido: {cantidad_rep}"
                            ))
                        
                        # 5.6 Validar que 'Código FID_rep' no esté vacío
                        if codigo_fid_rep is None or str(codigo_fid_rep).strip() == "" or str(codigo_fid_rep).strip() == "nan" or str(codigo_fid_rep).strip() == "NAN":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="El campo 'Código FID_rep' no puede estar vacío."
                            ))

                         # 5.7 Validar que 'Número de conductores_rep' sea numérico y válido según la hoja
                        if numero_conductores_rep is None or not str(numero_conductores_rep).strip().replace('.', '', 1).isdigit():
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="El campo 'Número de conductores_rep' no debe estar vacío y debe ser numérico."
                            ))
                        else:
                            try:
                                numero = int(float(numero_conductores_rep))
                                
                                # Solo aplica esta validación adicional en hojas específicas
                                if hoja in ["Conductor_N1", "Conductor_N2-N3"]:
                                    if numero not in [1, 2, 3]:
                                        novedades.append(Novedad(
                                            archivo=nombre_archivo,
                                            hoja=hoja,
                                            fila=num_fila,
                                            descripcion=(
                                                f"El campo 'Número de conductores_rep' debe tener un valor entre 1 y 3. "
                                                f"Valor recibido: {numero_conductores_rep}"
                                            )
                                        ))
                            except ValueError:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(
                                        f"No se pudo convertir 'Número de conductores_rep' a número. Valor recibido: {numero_conductores_rep}"
                                    )
                                ))
   
                        # 5.8 Validar que 'Año entrada operación_rep' esté entre 1900 y el año actual
                        try:
                            año = int(float(str(año_entrada_rep).strip()))

                            if not (1900 <= año <= año_actual):
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(
                                        f"El campo 'Año entrada operación_rep' debe estar entre 1900 y {año_actual}. "
                                        f"Valor recibido: {año_entrada_rep}"
                                    )
                                ))
                        except (ValueError, TypeError):
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(
                                    f"El campo 'Año entrada operación_rep' debe ser numérico y estar entre 1900 y {año_actual}. "
                                    f"Valor recibido: {año_entrada_rep}"
                                )
                            ))

                        # 5.9 Validar que 'Rpp_rep' sea 0 o 1
                        try:
                            valor_rpp_rep = int(float(str(rpp_rep).strip()))
                            if valor_rpp_rep not in [0, 1]:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"'Rpp_rep' debe ser 0 o 1. El valor recibido es: {rpp_rep}"
                                ))
                        except (ValueError, TypeError):
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Rpp_rep' debe ser un número entero (0 o 1). Valor recibido: {rpp_rep}"
                            ))

                        # 5.10 Validar que "Año entrada operación" no esté vacío y que "Año salida operación" sea igual
                        # Validar si está vacío o nulo el año de entrada
                        if pd.isna(anio_entrada_op) or str(anio_entrada_op).strip() == "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="'Año entrada operación' no puede estar vacío."
                            ))
                        else:
                            try:
                                entrada = int(float(str(anio_entrada_op).strip()))
                                salida = int(float(str(anio_salida_op).strip()))

                                if entrada != salida:
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion=(f"'Año salida operación' debe ser igual a 'Año entrada operación' ({entrada}). "
                                                    f"Valor recibido: {anio_salida_op}")
                                    ))
                            except (ValueError, TypeError):
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"'Año salida operación' debe ser numérico y coincidir con 'Año entrada operación'. "
                                                f"Valor recibido: {anio_salida_op}")
                                ))

                    # === 🔵 VALIDACIÓN 6: VALIDACIONES ESPECÍFICAS PARA II y IV (Expansión) ===
                    elif tipo_inversion in {"II", "IV"}:
                        
                        # 6.1 Validación: 'Código FID_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(codigo_fid_rep) and str(codigo_fid_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Código FID_rep' no debe tener valor. El valor recibido es: {codigo_fid_rep}"
                            ))
                            
                        # 6.2 Validación: 'Número de conductores_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(numero_conductores_rep) and str(numero_conductores_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Número de conductores_rep' no debe tener valor. El valor recibido es: {numero_conductores_rep}"
                            ))

                        # 6.3 Validación: 'Cantidad_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(cantidad_rep) and str(cantidad_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Cantidad_rep' no debe tener valor. El valor recibido es: {cantidad_rep}"
                            ))

                        # 6.4 Validación: 'Rpp_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(rpp_rep) and str(rpp_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Rpp_rep' no debe tener valor. El valor recibido es: {rpp_rep}"
                            ))

                        # 6.5 Validación: 'Año entrada operación_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(anio_entrada_rep) and str(anio_entrada_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'Año entrada operación_rep' no debe tener valor. El valor recibido es: {anio_entrada_rep}"
                            ))

                        # 6.6 Validación: 'Codigo UC_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if uc_rep not in [None, "", "nan", "NAN"] and not pd.isna(uc_rep):
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(
                                    f"Si 'Tipo inversión' es '{tipo_inversion}' el campo 'Codigo UC_rep' debe estar vacío. "
                                    f"El valor de 'Codigo UC_rep' es: {uc_rep}"
                                )
                            ))

                        # 6.7 Validación: 'DESCRIPCION_rep' no debe tener valor cuando es tipo de inversión II o IV
                        if not pd.isna(descripcion_rep) and str(descripcion_rep).strip() != "":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"'DESCRIPCION_rep' no debe tener valor. El valor recibido es: {descripcion_rep}"
                            ))

                    # === 🟡 VALIDACIÓN 7: Código Línea debe estar en el archivo pivote ===
                    if codigo_linea and codigo_linea not in lineas_validas:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=(f"El Código Línea '{codigo_linea}' no se encuentra en el archivo pivote, "
                                        "columna 'Código línea' de la hoja 'lineas'.")
                        ))

                    # === 🟡 VALIDACIÓN 8: Municipio debe estar en el archivo pivote ===
                    if municipio and municipio not in municipios_validos:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=(f"El municipio '{municipio}' no se encuentra en el archivo pivote, "
                                        "columna 'MUNICIPIOS' de la hoja 'municipio'.")
                        ))
                    
                    # === 🟡 VALIDACIÓN 9: Reglas entre 'Código FID' y 'Cantidad' ===
                    if codigo_fid in ["", "nan", "NAN"]:
                        # Si 'Código FID' es nulo o vacío, entonces 'Cantidad' también debe estar vacía
                        if cantidad is not None and str(cantidad).strip() not in ["", "nan", "NAN"]:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"Error en 'Cantidad' debido a que 'Código FID' es nulo. "
                                            f"El valor en 'Cantidad' es: {cantidad}")
                            ))
                    else:
                        # Si 'Código FID' tiene valor, la 'Cantidad' debe ser mayor a 0
                        try:
                            cantidad_valor = float(str(cantidad).strip())
                            if cantidad_valor <= 0:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"La 'Cantidad' debe ser mayor a cero. "
                                                f"El valor recibido es: {cantidad}")
                                ))
                            # Además, si es hoja 'Transformador_Distribucion', la cantidad debe ser 1
                            if hoja.strip().lower() == "transformador_distribucion" and cantidad_valor != 1:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"En la hoja 'Transformador_Distribucion' se esperaba una 'Cantidad' de 1. "
                                                f"El valor recibido es: {cantidad}")
                                ))
                        except:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"La 'Cantidad' debe ser un valor numérico. "
                                            f"Valor recibido: {cantidad}")
                            ))

                    # === 🟡 VALIDACIÓN 10: Reglas según el campo 'nivel' y 'transformador' ===
                    if codigo_fid and not pd.isna(codigo_fid) and str(codigo_fid).strip().upper() not in ["", "NAN"]:    # Solo aplica si el FID no es nulo
                        try:
                            nivel = int(float(str(nivel).strip()))  # Obtener y convertir el nivel a entero
                        except:
                            nivel = None  # Si no es numérico, asignamos -1 como valor inválido
                        if nivel == 1:
                            if transformador and len(transformador) == 7 and transformador[:2] in {"1T", "2T", "3T", "4T", "5T"}:
                                pass  # Cumple con el formato esperado
                            else:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"Para 'nivel' igual a 1, 'transformador' debe ser una cadena de 7 caracteres que empiece con '1T', '2T', '3T', '4T' o '5T'. Valor recibido: '{transformador}'"
                                ))

                        elif nivel in [2, 3]:
                            if str(transformador).strip() != "*":
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"Para 'nivel' igual a 2 o 3, 'transformador' debe ser '*'. Valor recibido: '{transformador}'"
                                ))

                        elif nivel == 4:
                            try:
                                tension_valor = float(str(tension_operacion).strip())
                                if tension_valor > 34.5 and str(transformador).strip() == "*":
                                    pass  # Correcto
                                else:
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion="Para 'nivel' igual a 4, 'Tensión de operación' debe ser > 34.5 y 'transformador' debe ser '*'. Valores recibidos: tensión = {}, transformador = '{}'".format(tension_operacion, transformador)
                                    ))
                            except:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion="Error al interpretar 'Tensión de operación' como numérico para 'nivel' 4. Valor recibido: '{}'".format(tension_operacion)
                                ))

                        elif nivel == 0 and uc in {"N4L93", "N4L94"}:
                            pass  # Excepción para fibra óptica

                        else:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"El valor en la columna 'nivel' debe ser 0 con UC especial o 1, 2 o 3. Valor recibido: {nivel}"
                            ))

                    # === 🟡 VALIDACIÓN 11: ' Verificar si el valor de "Nombre" es nulo ===
                    if not nombre or nombre.upper() in {"", "NAN"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error en 'Nombre' debido a que es nulo o vacío."
                        ))

                    # === 🟡 VALIDACIÓN 12: ' Verificar si el valor de "Contrato/Soporte" es nulo ===
                    if not contrato or contrato.upper() in {"", "NAN"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error en 'Contrato/Soporte' debido a que es nulo o vacío."
                        ))
                    
                    # === 🟡 VALIDACIÓN 13: Verificar si el valor de "Nombre de la Plantilla" es igual al nombre del archivo sin extensión o si la celda está vacía ===
                    if not nombre_plantilla or nombre_plantilla.upper() in {"", "NAN"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="La columna 'Nombre de la Plantilla' está vacía."
                        ))
                    elif nombre_plantilla != nombre_archivo_sin_extension:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Nombre de la Plantilla' no coincide con el nombre del archivo."
                        ))

                    # === 🟡 VALIDACIÓN 14: Reglas para RPP según si 'Código FID' está vacío o no ===
                    if not codigo_fid or str(codigo_fid).strip().upper() in {"", "NAN"}:
                        # Si 'Código FID' está vacío, 'RPP' también debe estar vacío
                        if rpp is not None and str(rpp).strip().upper() not in {"", "NAN"}:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="'RPP' debe estar vacío cuando 'Código FID' está vacío. Valor recibido: {}".format(rpp)
                            ))
                    else:
                        # Si 'Código FID' tiene valor, 'RPP' debe ser 0 o 1
                        try:
                            valor_rpp = int(float(str(rpp).strip()))
                            if valor_rpp not in [0, 1]:
                                raise ValueError
                        except:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="'RPP' debe ser 0 o 1 cuando 'Código FID' tiene valor. Valor recibido: {}".format(rpp)
                            ))


                    # === 🟡 VALIDACIÓN 15: Validar campo 'Sobrepuesto' sea 'S' o 'N' ===
                    if sobrepuesto not in {"S", "N"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="'Sobrepuesto' debe ser 'S' o 'N'."
                        ))

                    # === 🟡 VALIDACIÓN 16: Reglas para 'Número de conductores' según 'Código FID' ===
                    if not codigo_fid or str(codigo_fid).strip().upper() in ["", "NAN"]:
                        # Si 'Código FID' es nulo o vacío, 'Número de conductores' también debe ser nulo o vacío
                        if numero_conductores is not None and str(numero_conductores).strip().upper() not in ["", "NAN"]:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="'Número de conductores' debe ser nulo cuando 'Código FID' es nulo."
                            ))
                    else:
                        # Si 'Código FID' tiene valor, validar 'Número de conductores'
                        try:
                            numero = int(float(str(numero_conductores).strip()))
                            if hoja in {"Conductor_N1", "Conductor_N2-N3"} and numero not in {1, 2, 3}:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"'Número de conductores' debe ser 1, 2 o 3 cuando 'Código FID' tiene un valor. "
                                                f"Valor recibido: {numero_conductores}")
                                ))
                        except:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="'Número de conductores' no debe ser nulo y debe ser un valor numérico cuando 'Código FID' tiene un valor."
                            ))
                    
                    # === 🟡 VALIDACIÓN 17: Validar Unidad Constructiva contra el pivote y su tensión de operación ===
                    if uc and uc not in uc_validas:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=(f"La Unidad Constructiva '{uc}' no se encuentra en el archivo pivote "
                                        "en la columna 'Código UC' de la hoja 'uc'.")
                        ))
                    elif uc:  # Si tiene UC válida
                        if uc in {"N4L93", "N4L94", "N4L52"}:
                            try:
                                if float(tension_operacion) != 0:
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion=("'Tensión de operación' debe ser 0 para "
                                                    "la UC N4L93, N4L94 o N4L52.")
                                    ))
                            except:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion="'Tensión de operación' debe ser un número válido para UC N4L93/N4L94/N4L52."
                                ))
                        elif nivel is not None:
                            try:
                                nivel_int = int(float(nivel))
                                tension = float(tension_operacion)
                                if nivel_int == 2 and not (1 <= tension < 20):
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion="Para 'Nivel' 2, la 'Tensión de operación' debe estar entre 1 y 20."
                                    ))
                                elif nivel_int == 3 and not (30 <= tension < 57.5):
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion="Para 'Nivel' 3, la 'Tensión de operación' debe estar entre 30 y 57.5."
                                    ))
                                elif nivel_int == 4 and not (57.5 <= tension < 220):
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion="Para 'Nivel' 4, la 'Tensión de operación' debe estar entre 57.5 y 220."
                                    ))
                            except:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion="'Nivel' y 'Tensión de operación' deben ser numéricos."
                                ))
                    else:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="La Unidad Constructiva es nula o vacía."
                        ))

                    # === 🟡 VALIDACIÓN 18: Validar que 'DESCRIPCION' no sea nulo ni de tipo incorrecto ===
                    if descripcion is None or str(descripcion).strip() in {"", "NAN"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error en 'DESCRIPCION': Campo vacío o no es una cadena de caracteres."
                        ))
                    elif not isinstance(descripcion, str):
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error en 'DESCRIPCION': No es una cadena de caracteres."
                        ))
                    
                    # === 🟡 VALIDACIÓN 19: Validar 'Código proyecto' con formato válido ===     
                    valor_codigo_proyecto = str(codigo_proyecto).strip().upper()

                    if not valor_codigo_proyecto or valor_codigo_proyecto == "NAN":
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Código proyecto' está vacío."
                        ))
                    elif not re.match(r"^[A-Z0-9\-_]+$", valor_codigo_proyecto):
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Código proyecto' no cumple con los requisitos."
                        ))
                    
                    # === 🟡 VALIDACIÓN 20: Validar 'Nivel' como número entero entre 0 y 4 ===
                    valor_nivel = str(nivel).strip().upper()

                    if not valor_nivel or valor_nivel == "NAN":
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Nivel' es nulo."
                        ))
                    else:
                        try:
                            nivel_int = int(float(valor_nivel))
                            if nivel_int not in [0, 1, 2, 3, 4]:
                                raise ValueError
                        except:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion="El valor en la columna 'Nivel' no es válido. Debe ser un número entero entre 0 y 4."
                            ))

                    # === 🟡 VALIDACIÓN 21: 'Cod Regional' debe estar en el archivo pivote hoja 'regionales' ===
                    if cod_regional and cod_regional not in regionales_validos:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=f"El Cod Regional '{cod_regional}' no se encuentra en el archivo pivote, columna 'Cod Dane' de la hoja 'regionales'."
                        ))

                    # === 🟡 VALIDACIÓN 22: 'Cod DANE municipio' debe estar en el archivo pivote hoja 'municipio' ===
                    if cod_dane_municipio and cod_dane_municipio not in cod_dane_validos:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=f"El Cod DANE municipio '{cod_dane_municipio}' no se encuentra en el archivo pivote, columna 'CODIGO' de la hoja 'municipio'."
                        ))
                    
                    # === 🟡 VALIDACIÓN 23: Año entrada operación debe ser 2025 ===
                    if anio_entrada_op is None or str(anio_entrada_op).strip() != "2025":
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="El valor en la columna 'Año entrada operacion' no es igual a 2025."
                        ))

                    # # === 🟡 VALIDACIÓN 24: Verificar si el valor de "IUS" no está en el archivo pivote, excluyendo 0 ===
                    if ius is not None:
                        ius_str = str(ius).strip()  # No lo convertimos a número, solo texto limpio

                        if ius_str == "*" and codigo_linea == "SANALBERTO":
                            pass  # Excepción válida
                        else:
                            # Verifica que tenga exactamente 4 dígitos y todos sean numéricos
                            if not (len(ius_str) == 4 and ius_str.isdigit()):
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"El IUS '{ius_str}' tiene un formato inválido. Debe tener exactamente 4 dígitos con ceros a la izquierda."
                                ))
                            elif ius_str != "0000" and ius_str not in ius_validos:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=f"El IUS '{ius_str}' no se encuentra en el archivo pivote en la columna 'IUS' de la hoja 'subestaciones'."
                                ))


                    # === 🟡 VALIDACIÓN 25: Validación de 'Fracción costo' y 'Porcentaje uso' según 'Código FID' ===
                    # Caso 1: Si 'Código FID' es nulo, entonces 'Fracción costo' también debe ser nulo
                    if (codigo_fid is None or codigo_fid.strip() == "" or codigo_fid.strip().lower() == "nan"):
                        if fraccion_costo != "" and fraccion_costo.lower() != "nan":
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=f"Error en 'Fracción costo' debido a que 'Código FID' es nulo. Valor recibido en 'Fracción costo': '{fraccion_costo}'."
                            ))
                    else:
                        # Caso 2: Si 'Código FID' existe, verificar que 'Fracción costo' esté entre 0 y 100
                        try:
                            fraccion_valor = float(fraccion_costo)
                            if not (0 <= fraccion_valor <= 100):
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion = f"Error en 'Fracción costo': Debe ser un valor entre 0 y 100. Valor recibido: '{fraccion_costo}'."
                                ))
                        except (ValueError, TypeError):
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion = f"Error en 'Fracción costo': Valor no numérico o inválido. Valor recibido: '{fraccion_costo}'."
                            ))

                        # Caso 3: Validar 'Porcentaje uso' > 0 y <= 100
                        try:
                            porcentaje_valor = float(porcentaje_uso)
                            if not (0 < porcentaje_valor <= 100):
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion = f"Error: 'Porcentaje uso' debe ser mayor a 0 y menor o igual a 100. Valor recibido: '{porcentaje_uso}'."
                                ))
                        except (ValueError, TypeError):
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion = f"Error: 'Porcentaje uso' debe ser un número válido mayor a 0 y menor o igual a 100. Valor recibido: '{porcentaje_uso}'."
                            ))

                    # === 🟡 VALIDACIÓN 26: Validación de 'IUL' debe ser una cadena de 4 caracteres ===
                    if not iul or len(iul) != 4:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=f"El valor en la columna 'IUL' no es válido. Debe ser una cadena de 4 caracteres. Valor recibido: '{iul}'."
                        ))
                    
                    # === 🟡 VALIDACIÓN 27: Validar que 'IUS Final' esté en el archivo pivote si es distinto de 0 ===
                    if ius_final is not None:
                        try:
                            ius_final_int = int(float(str(ius_final).strip()))

                            if ius_final_int != 0 and str(ius_final_int) not in ius_validos:
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"El IUS final '{ius_final_int}' no se encuentra en el archivo pivote, "
                                                "columna 'IUS' de la hoja 'subestaciones'.")
                                ))
                        except Exception as ex:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"El IUS final '{ius_final}' tiene un formato inválido: {str(ex)}")
                            ))
                    
                    # === 🟡 VALIDACIÓN 28: Validar que 'IUS inicial' esté en el archivo pivote, con excepción para SANALBERTO ===
                    if ius_inicial is not None:
                        valor_ius = str(ius_inicial).strip().upper()

                        if valor_ius == "*" and codigo_linea == "SANALBERTO":
                            pass  # No validar contra el pivote si es caso especial
                        elif valor_ius not in ius_validos:
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"El IUS inicial '{ius_inicial}' no se encuentra en el archivo pivote, "
                                            "columna 'IUS' de la hoja 'subestaciones'.")
                            ))

                    # === 🟡 VALIDACIÓN 29: Validar campo 'PIEC' ===
                    if not piec or piec not in {"S", "N"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error: 'PIEC' debe ser 'S' o 'N'."
                        ))

                    # === 🟡 VALIDACIÓN 30: Validar campo 'STR construcción' ===
                    if not str_construccion or str_construccion not in {"S", "N"}:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error: 'STR construcción' debe ser 'S' o 'N'."
                        ))

                    # === 🟡 VALIDACIÓN 31: Validar campo 'Tipo de Proyecto' entre 1 y 5 ===
                    try:
                        tipo_proyecto_int = int(float(tipo_proyecto))
                        if tipo_proyecto_int not in {1, 2, 3, 4, 5}:
                            raise ValueError
                    except:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error: 'Tipo de Proyecto' debe ser uno de los valores entre 1 y 5."
                        ))

                    # === 🟡 VALIDACIÓN 32: Validar campo 'Activo Construido y en Operación' sea 0 o 1 ===
                    try:
                        activo_construido_int = int(float(activo_construido))
                        if activo_construido_int not in {0, 1}:
                            raise ValueError
                    except:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion="Error: 'Activo Construido y en Operación' debe ser 0 o 1."
                        ))
                    
                   # === 🟡 VALIDACIÓN 33: Transformador excluido (solo hoja específica) ===
                    if hoja.strip().lower() == "transformador_distribucion" and transformador in trafos_excluidos:
                        novedades.append(Novedad(
                            archivo=nombre_archivo,
                            hoja=hoja,
                            fila=num_fila,
                            descripcion=f"El Transformador '{transformador}' está en la lista de excluidos del archivo pivote y no debe ser reportado."
                        )) 
                    
                    # === 🟡 VALIDACIÓN 34: Cruce con archivo pivote 'equipo_proteccion' para hoja Equipos Proteccion ===
                    if hoja.strip().lower() == "equipos proteccion":
                        print(f"\n🛠 Fila {num_fila} — VALIDACIÓN 34")
                        print(f"➡️ Código FID_rep: '{codigo_fid_rep}'")
                        print(f"➡️ Código UC_rep: '{uc_rep}'")
                        print(f"➡️ Año entrada operación_rep: '{anio_entrada_rep}'")
                        print(f"➡️ Código FID: '{codigo_fid}'")

                        # Si 'Código FID_rep' está lleno, debe existir en el pivote
                        if codigo_fid_rep and str(codigo_fid_rep).strip().upper() not in ["", "NAN", "NONE"]:
                            if codigo_fid_rep in dicc_equipo_proteccion:
                                fila_pivote = dicc_equipo_proteccion[codigo_fid_rep]
                                if uc_rep != fila_pivote["codigo_uc"] or anio_entrada_rep != fila_pivote["anio_operacion"]:
                                    print("⚠️ Coincidencia encontrada, pero los datos no coinciden con el pivote.")
                                    novedades.append(Novedad(
                                        archivo=nombre_archivo,
                                        hoja=hoja,
                                        fila=num_fila,
                                        descripcion=(f"Datos no coinciden con el archivo pivote para el Código FID_rep '{codigo_fid_rep}'. "
                                                    f"UC esperada: {fila_pivote['codigo_uc']}, Año operación esperado: {fila_pivote['anio_operacion']}")
                                    ))
                            else:
                                print("❌ Código FID_rep no está en el pivote y no es vacío.")
                                novedades.append(Novedad(
                                    archivo=nombre_archivo,
                                    hoja=hoja,
                                    fila=num_fila,
                                    descripcion=(f"El Código FID_rep '{codigo_fid_rep}' no se encuentra en el archivo pivote 'equipo_proteccion'.")
                                ))
                        else:
                            print("✅ Código FID_rep está vacío. No se valida contra el pivote.")

                        # Verificar que 'Código FID' no esté en el pivote
                        if codigo_fid and codigo_fid in dicc_equipo_proteccion:
                            print("❌ Código FID del Excel ya existe en el pivote (lo cual no debe ocurrir).")
                            novedades.append(Novedad(
                                archivo=nombre_archivo,
                                hoja=hoja,
                                fila=num_fila,
                                descripcion=(f"El Código FID ingresado '{codigo_fid}' ya existe en el archivo pivote y no debe estar allí.")
                            ))

                        
        except Exception as e:
            novedades.append(Novedad(
                archivo=nombre_archivo,
                hoja="-",
                fila=0,
                descripcion=f"Error al procesar archivo: {str(e)}"
            ))

    return [n.__dict__ for n in novedades]



# 📂 Ruta del archivo pivote local o en red
RUTA_PIVOTE = "D:/OneDrive - Grupo EPM/Descargas/Pivote.xlsx"  # ajustar si el archivo pivote esta en otra carpeta


def cargar_datos_pivote(hoja: str, columna: int = 0) -> set:
    """
    Carga los datos únicos (no vacíos) de una hoja y columna del archivo pivote.
    """
    try:
        df = pd.read_excel(RUTA_PIVOTE, sheet_name=hoja, dtype=str)
        valores = df.iloc[:, columna].dropna().astype(str).str.strip().str.upper()
        return set(valores)
    except Exception as e:
        print(f"⚠️ Error cargando pivote: hoja={hoja}, columna={columna}, error={e}")
        return set()

