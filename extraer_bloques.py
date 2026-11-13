# -*- coding: utf-8 -*-
"""
Script de extracción de bloques dinámicos por estados de visibilidad.
Conecta con AutoCAD usando pyautocad (create_if_not_exists=True),
recorre los bloques dinámicos definidos en LIBRERIA_BACKUP_09.10.2026.dxf,
y genera un nuevo archivo LIBRERIA INDIVIDUAL.dxf con un bloque individual
por estado de visibilidad, conservando punto de inserción original,
atributos y capas.
"""

import os
import sys
import time
from pyautocad import Autocad, APoint


# Bloques dinámicos identificados en LIBRERIA_BACKUP_09.10.2026.dxf
BLOQUES_DINAMICOS = [
    "ANCLAS",
    "CORTA_CIRCUITO",
    "TIERRAS",
    "PARARRAYOS",
    "TRAFOS",
    "ESTRUCTURAS",
    "P.C",
    "AMORTIGUADORES",
    "CIMENTACION",
    "P.M.",
    "P.F",
    "P.METAL",
    "VIVIENDAS",
    "p.existentes",
]

DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_ORIGINAL = os.path.join(DIRECTORIO_BASE, "Bloques", "LIBRERIA_BACKUP_09.10.2026.dxf")
ARCHIVO_DESTINO = os.path.join(DIRECTORIO_BASE, "Bloques", "LIBRERIA INDIVIDUAL.dxf")


def conectar_autocad():
    """Establece conexión con AutoCAD usando pyautocad."""
    try:
        acad = Autocad(create_if_not_exists=True)
        print(f"Conectado a: {acad.doc.Name}")
        return acad
    except Exception as e:
        print(f"[ERROR] No se pudo conectar con AutoCAD: {e}")
        print("Asegúrate de que AutoCAD esté abierto con el archivo cargado.")
        sys.exit(1)


def hacer_reintento(funcion, *args, reintentos=3, espera=1.0):
    """Ejecuta una función COM con reintentos en caso de E_CALL_REJECTED."""
    for intento in range(reintentos):
        try:
            return funcion(*args)
        except Exception as e:
            if "E_CALL_REJECTED" in str(e) or "-2147418111" in str(e):
                if intento < reintentos - 1:
                    print(f"    [REINTENTO {intento + 1}/{reintentos}] AutoCAD ocupado, esperando {espera}s...")
                    time.sleep(espera)
                    continue
            raise


def insertar_referencia(acad, nombre_bloque, x, y):
    """Inserta una referencia del bloque dinámico en el punto indicado."""
    comando = f"_.-INSERT {nombre_bloque} {x},{y} 1 1 0 "
    hacer_reintento(acad.doc.SendCommand, comando)


def obtener_estados_visibilidad(referencia):
    """Obtiene los nombres de los estados de visibilidad de un bloque dinámico."""
    try:
        props = hacer_reintento(referencia.GetDynamicBlockProperties)
        for i in range(props.Count):
            prop = props.Item(i)
            if prop.PropertyName == "Visibilidad":
                allowed = prop.AllowedValues
                estados = []
                for j in range(allowed.Count):
                    estados.append(allowed.Item(j))
                return estados
    except Exception as e:
        print(f"  [AVISO] No se pudieron obtener estados de visibilidad: {e}")
    return []


def cambiar_estado_visibilidad(referencia, estado):
    """Cambia el estado de visibilidad de una referencia de bloque dinámico."""
    try:
        props = hacer_reintento(referencia.GetDynamicBlockProperties)
        for i in range(props.Count):
            prop = props.Item(i)
            if prop.PropertyName == "Visibilidad":
                prop.Value = estado
                return True
    except Exception as e:
        print(f"    [ERROR] No se pudo cambiar el estado: {e}")
    return False


def crear_bloque_desde_entidades(acad, nombre_bloque_nuevo, punto_insercion, entidades):
    """
    Crea un bloque individual en el diccionario de bloques usando las
    entidades proporcionadas, conservando el punto de inserción original.
    """
    try:
        # Verificar si ya existe un bloque con ese nombre
        for i in range(acad.doc.Blocks.Count):
            if acad.doc.Blocks.Item(i).Name == nombre_bloque_nuevo:
                print(f"    [AVISO] El bloque '{nombre_bloque_nuevo}' ya existe. Eliminando...")
                acad.doc.Blocks.Item(i).Delete()
                break

        # Crear el nuevo bloque en el diccionario
        # Blocks.Add(InsertionPoint, Name) - InsertionPoint es un array de 3 doubles
        bloque_nuevo = acad.doc.Blocks.Add(punto_insercion, nombre_bloque_nuevo)

        # Reunir entidades válidas
        objetos_a_copiar = []
        for ent in entidades:
            try:
                objetos_a_copiar.append(ent)
            except:
                pass

        if not objetos_a_copiar:
            print(f"    [ERROR] No hay entidades válidas para copiar.")
            return False

        # CopyObjects copia las entidades desde modelspace al bloque
        # Sintaxis: Block.CopyObjects(Objects, Destination)
        # Objects es un array de variant con las entidades
        bloque_nuevo.CopyObjects(objetos_a_copiar, acad.model)

        return True
    except Exception as e:
        print(f"    [ERROR] No se pudo crear el bloque: {e}")
        return False


def eliminar_entidades(entidades):
    """Elimina las entidades indicadas del modelspace."""
    for ent in entidades:
        try:
            ent.Delete()
        except:
            pass


def limpiar_modelspace(acad):
    """Limpia TODO el modelspace eliminando todas las entidades."""
    try:
        hacer_reintento(acad.doc.SendCommand, "_.ERASE _ALL ")
    except:
        pass


def guardar_como(acad, ruta_destino):
    """Guarda el documento con el nombre indicado."""
    try:
        hacer_reintento(acad.doc.SaveAs, ruta_destino)
        print(f"Archivo guardado: {ruta_destino}")
        return True
    except Exception as e:
        print(f"[ERROR] No se pudo guardar el archivo: {e}")
        return False


def verificar_bloque_existe(acad, nombre_bloque):
    """Verifica si un bloque existe en el diccionario de bloques."""
    try:
        for i in range(acad.doc.Blocks.Count):
            if acad.doc.Blocks.Item(i).Name == nombre_bloque:
                return True
    except:
        pass
    return False


def obtener_punto_origen(acad, nombre_bloque):
    """Obtiene el punto de inserción original de un bloque."""
    try:
        for i in range(acad.doc.Blocks.Count):
            bloque = acad.doc.Blocks.Item(i)
            if bloque.Name == nombre_bloque:
                origen = bloque.Origin
                return [origen[0], origen[1], origen[2]]
    except Exception as e:
        print(f"  [AVISO] No se pudo obtener el punto base: {e}")
    return [0.0, 0.0, 0.0]


def main():
    print("=" * 60)
    print("EXTRACTOR DE BLOQUES DINÁMICOS POR ESTADOS DE VISIBILIDAD")
    print("=" * 60)

    # Verificar que el archivo original existe
    if not os.path.exists(ARCHIVO_ORIGINAL):
        print(f"[ERROR] No se encuentra el archivo: {ARCHIVO_ORIGINAL}")
        return False

    # Conectar con AutoCAD
    print("\nConectando con AutoCAD...")
    acad = conectar_autocad()

    # Verificar que el archivo cargado es el correcto
    nombre_doc = acad.doc.Name
    print(f"Documento activo: {nombre_doc}")

    if "LIBRERIA" not in nombre_doc.upper():
        respuesta = input(
            f"\n[AVISO] El documento activo es '{nombre_doc}', no parece ser "
            f"LIBRERIA_BACKUP_09.10.2026.dxf.\n"
            f"¿Deseas continuar? (s/n): "
        )
        if respuesta.lower() != 's':
            print("Operación cancelada.")
            return False

    # Limpiar modelspace al inicio
    print("\nLimpiando modelspace...")
    limpiar_modelspace(acad)

    # Registrar resultados
    resultados = []
    total_bloques_creados = 0

    for nombre_bloque in BLOQUES_DINAMICOS:
        print(f"\n{'─' * 50}")
        print(f"Procesando bloque dinámico: {nombre_bloque}")
        print(f"{'─' * 50}")

        # Verificar que el bloque existe en el documento
        if not verificar_bloque_existe(acad, nombre_bloque):
            print(f"  [AVISO] El bloque '{nombre_bloque}' no existe en el documento. Saltando.")
            resultados.append((nombre_bloque, "SALTADO", []))
            continue

        # Obtener el punto de inserción original del bloque
        punto_base = obtener_punto_origen(acad, nombre_bloque)
        print(f"  Punto de inserción original: ({punto_base[0]:.4f}, {punto_base[1]:.4f})")

        # Insertar referencia del bloque en (0,0)
        print(f"  Insertando referencia en (0, 0)...")
        indice_antes = acad.model.Count
        insertar_referencia(acad, nombre_bloque, 0, 0)

        # Obtener la referencia insertada
        entidades_nuevas = []
        for i in range(indice_antes, acad.model.Count):
            entidades_nuevas.append(acad.model.Item(i))

        if not entidades_nuevas:
            print(f"  [ERROR] No se insertó la referencia. Saltando.")
            resultados.append((nombre_bloque, "ERROR_INSERCION", []))
            continue

        referencia = entidades_nuevas[0]

        # Obtener estados de visibilidad
        print(f"  Obteniendo estados de visibilidad...")
        estados = obtener_estados_visibilidad(referencia)

        if not estados:
            print(f"  [AVISO] No se encontraron estados de visibilidad. Saltando.")
            eliminar_entidades(entidades_nuevas)
            resultados.append((nombre_bloque, "SIN_ESTADOS", []))
            continue

        print(f"  Estados encontrados: {len(estados)}")
        for i, estado in enumerate(estados, 1):
            print(f"    {i}. {estado}")

        # Procesar cada estado de visibilidad
        bloques_creados = []

        for estado in estados:
            print(f"\n  → Procesando estado: '{estado}'")

            # Insertar nueva referencia para este estado
            indice_antes_estado = acad.model.Count
            insertar_referencia(acad, nombre_bloque, 0, 0)

            entidades_ref = []
            for i in range(indice_antes_estado, acad.model.Count):
                entidades_ref.append(acad.model.Item(i))

            if not entidades_ref:
                print(f"    [ERROR] No se insertó la referencia.")
                continue

            ref_actual = entidades_ref[0]

            # Cambiar el estado de visibilidad
            if not cambiar_estado_visibilidad(ref_actual, estado):
                eliminar_entidades(entidades_ref)
                continue
            print(f"    Estado cambiado a: '{estado}'")

            # Explotar la referencia
            print(f"    Explotando referencia...")
            hacer_reintento(acad.doc.SendCommand, "_.EXPLODE ")
            hacer_reintento(acad.doc.SendCommand, "_.SELECT _L ")

            # Obtener entidades resultantes (las nuevas desde la inserción de la referencia)
            entidades_para_bloque = []
            for i in range(indice_antes_estado, acad.model.Count):
                entidades_para_bloque.append(acad.model.Item(i))

            if not entidades_para_bloque:
                print(f"    [ERROR] No se generaron entidades al explotar.")
                limpiar_modelspace(acad)
                continue

            # Crear bloque individual con las entidades explotadas
            print(f"    Creando bloque '{estado}'...")

            if crear_bloque_desde_entidades(acad, estado, punto_base, entidades_para_bloque):
                bloques_creados.append(estado)
                total_bloques_creados += 1
                print(f"    ✓ Bloque '{estado}' creado correctamente.")
            else:
                print(f"    [ERROR] No se pudo crear el bloque '{estado}'.")

            # Limpiar las entidades temporales de este estado
            limpiar_modelspace(acad)

        # Limpiar la referencia original del bloque dinámico
        eliminar_entidades(entidades_nuevas)

        resultados.append((nombre_bloque, "OK", bloques_creados))
        print(f"\n  Total de bloques creados para '{nombre_bloque}': {len(bloques_creados)}")

    # Guardar como LIBRERIA INDIVIDUAL.dxf
    print(f"\n{'=' * 60}")
    print(f"Total de bloques individuales creados: {total_bloques_creados}")
    print("Guardando archivo final...")
    if guardar_como(acad, ARCHIVO_DESTINO):
        print(f"{'=' * 60}")
        print("PROCESO COMPLETADO")
        print(f"{'=' * 60}")
        print(f"\nResumen de bloques creados:")
        for nombre_bloque, estado, bloques in resultados:
            if estado == "OK":
                print(f"  ✓ {nombre_bloque}: {len(bloques)} bloques")
                for b in bloques:
                    print(f"      - {b}")
            else:
                print(f"  ✗ {nombre_bloque}: {estado}")
        return True
    else:
        print("[ERROR] No se pudo guardar el archivo final.")
        return False


if __name__ == "__main__":
    exito = main()
    sys.exit(0 if exito else 1)
