# -*- coding: utf-8 -*-
"""Diagnóstico simple de conexión COM con AutoCAD usando win32com."""
import sys
import time

print("Intentando importar win32com.client...", flush=True)
import win32com.client
print("win32com.client importado OK", flush=True)

print("Intentando GetActiveObject...", flush=True)
try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print(f"GetActiveObject OK: {acad}", flush=True)
except Exception as e:
    print(f"GetActiveObject FALLÓ: {e}", flush=True)
    sys.exit(1)

print("Intentando acceder a ActiveDocument...", flush=True)
try:
    doc = acad.ActiveDocument
    print(f"ActiveDocument OK: {doc.Name}", flush=True)
except Exception as e:
    print(f"ActiveDocument FALLÓ: {e}", flush=True)
    sys.exit(1)

print("Intentando acceder a ModelSpace.Count...", flush=True)
try:
    count = doc.ModelSpace.Count
    print(f"ModelSpace.Count = {count}", flush=True)
except Exception as e:
    print(f"ModelSpace.Count FALLÓ: {e}", flush=True)
    sys.exit(1)

print("Intentando acceder a Blocks.Count...", flush=True)
try:
    count = doc.Blocks.Count
    print(f"Blocks.Count = {count}", flush=True)
except Exception as e:
    print(f"Blocks.Count FALLÓ: {e}", flush=True)
    sys.exit(1)

print("\n--- Diagnóstico completado exitosamente ---", flush=True)
