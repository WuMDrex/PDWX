# -*- coding: utf-8 -*-
"""Diagnóstico rápido de conexión AutoCAD y bloques dinámicos."""
import sys
from pyautocad import Autocad, APoint

try:
    acad = Autocad(create_if_not_exists=False)
    print(f"Conectado a: {acad.doc.Name}")
    print(f"Total bloques: {acad.doc.Blocks.Count}")
    
    bloques_dinamicos = [
        "ANCLAS", "CORTA_CIRCUITO", "TIERRAS", "PARARRAYOS",
        "TRAFOS", "ESTRUCTURAS", "P.C", "AMORTIGUADORES",
        "CIMENTACION", "P.M.", "P.F", "P.METAL",
        "VIVIENDAS", "p.existentes",
    ]
    
    for nombre in bloques_dinamicos:
        encontrado = False
        for i in range(acad.doc.Blocks.Count):
            if acad.doc.Blocks.Item(i).Name == nombre:
                encontrado = True
                bloque = acad.doc.Blocks.Item(i)
                print(f"  ✓ {nombre} - Origen: ({bloque.Origin[0]:.2f}, {bloque.Origin[1]:.2f})")
                break
        if not encontrado:
            print(f"  ✗ {nombre} - NO ENCONTRADO")
    
    # Probar inserción y lectura de estados
    print("\n--- Prueba de inserción y estados ---")
    acad.doc.SendCommand("_.ERASE _ALL ")
    acad.doc.SendCommand("_.-INSERT ANCLAS 0,0 1 1 0 ")
    
    if acad.model.Count > 0:
        ref = acad.model.Item(acad.model.Count - 1)
        print(f"Referencia insertada: {ref.ObjectName}")
        try:
            props = ref.GetDynamicBlockProperties()
            print(f"Propiedades dinámicas: {props.Count}")
            for i in range(props.Count):
                p = props.Item(i)
                print(f"  - {p.PropertyName}")
                if p.PropertyName == "Visibilidad":
                    allowed = p.AllowedValues
                    print(f"    AllowedValues.Count: {allowed.Count}")
                    for j in range(allowed.Count):
                        print(f"      Estado {j}: {allowed.Item(j)}")
        except Exception as e:
            print(f"Error obteniendo propiedades: {e}")
    else:
        print("No se insertó ninguna referencia")
    
    # Limpiar
    acad.doc.SendCommand("_.ERASE _ALL ")
    print("\nDiagnóstico completado.")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
