import tkinter as tk
from tkinter import filedialog
import win32com.client

def abrir_archivo():
    ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo de ensamblaje", filetypes=[("Archivos de ensamblaje", "*.iam")])
    if ruta_archivo:
        # Crear una instancia de la aplicación de Autodesk Inventor
        invApp = win32com.client.Dispatch("Inventor.Application")
        print(invApp)

        # Abrir el ensamblaje
        doc = invApp.Documents.Open(ruta_archivo)

        # Acciones adicionales según lo que necesites hacer con el ensamblaje
        # Por ejemplo:
        # componente = doc.ComponentDefinition
        # Realizar otras operaciones con el componente...

# Configuración de la ventana principal
root = tk.Tk()
root.title("Aplicación para abrir archivos de ensamblaje")

# Botón para abrir archivo
btn_abrir_archivo = tk.Button(root, text="Abrir archivo de ensamblaje", command=abrir_archivo)
btn_abrir_archivo.pack(pady=20)

# Ejecutar la aplicación
root.mainloop()
