import tkinter as tk
from tkinter import filedialog
from typing import List

def seleccionar_archivos() -> List[str]:
    """
    Permite al usuario seleccionar múltiples archivos Excel mediante diálogo.
    
    Returns:
        List[str]: Lista de rutas de archivos seleccionados
    """
    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal
    
    archivos = filedialog.askopenfilenames(
        title="Selecciona los archivos Excel de asistencia",
        filetypes=[
            ("Archivos Excel", "*.xlsx *.xls"),
            ("Todos los archivos", "*.*")
        ]
    )
    
    root.destroy()
    return list(archivos)