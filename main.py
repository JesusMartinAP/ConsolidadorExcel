import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

# Función para seleccionar archivos Excel
def seleccionar_archivos():
    archivos = filedialog.askopenfilenames(
        title="Seleccionar archivos Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    return archivos

# Función para consolidar archivos Excel
def consolidar_archivos():
    archivos = seleccionar_archivos()
    if not archivos:
        messagebox.showwarning("Advertencia", "No se seleccionaron archivos.")
        return

    # Crear un dataframe vacío para el consolidado
    df_consolidado = pd.DataFrame()

    # Leer cada archivo y añadir al consolidado
    for archivo in archivos:
        df = pd.read_excel(archivo)
        df_consolidado = pd.concat([df_consolidado, df], ignore_index=True)

    # Guardar el archivo consolidado con la fecha actual
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nombre_archivo = f"Consolidado_{fecha_actual}.xlsx"
    df_consolidado.to_excel(nombre_archivo, index=False)

    messagebox.showinfo("Éxito", f"Consolidado guardado como {nombre_archivo}.")
    
    return nombre_archivo

# Función para abrir el archivo consolidado
def abrir_consolidado():
    archivo = consolidar_archivos()
    if archivo:
        os.startfile(archivo)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Consolidador de Archivos Excel")

# Botón para consolidar archivos
btn_consolidar = tk.Button(ventana, text="Consolidar Archivos", command=abrir_consolidado)
btn_consolidar.pack(pady=20)

# Iniciar el bucle de la interfaz
ventana.mainloop()
