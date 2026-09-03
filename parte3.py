import tkinter as tk
from tkinter import filedialog
import win32com.client

class AplicacionMedicion:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplicación para medir componentes")

        # Botón para abrir archivo de ensamblaje
        self.btn_abrir_archivo = tk.Button(root, text="Abrir archivo de ensamblaje", command=self.abrir_archivo)
        self.btn_abrir_archivo.pack(pady=20)

        # Variable para almacenar el archivo de ensamblaje abierto
        self.archivo_ensamblaje = None

    def abrir_archivo(self):
        ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo de ensamblaje", filetypes=[("Archivos de ensamblaje", "*.iam")])
        if ruta_archivo:
            self.archivo_ensamblaje = ruta_archivo
            print("Archivo de ensamblaje seleccionado:", self.archivo_ensamblaje)

            # Crear una instancia de la aplicación de Autodesk Inventor
            self.invApp = win32com.client.Dispatch("Inventor.Application")
            print(self.invApp)

            # Abrir el ensamblaje
            self.doc = self.invApp.Documents.Open(self.archivo_ensamblaje)

            # Esperar a que el usuario seleccione dos componentes
            self.root.bind("<Button-1>", self.seleccionar_componentes)

    def seleccionar_componentes(self, event):
        # Asegurarse de que la interfaz de usuario esté en modo de selección de componentes
        self.invApp.UserInterfaceManager.InteractionEvents.SelectSet.SelectByRay()
        
        # Esperar a que el usuario seleccione dos componentes
        while self.invApp.ActiveDocument.SelectSet.Count < 2:
            pass

        # Obtener los componentes seleccionados
        componente1 = self.invApp.ActiveDocument.SelectSet.Item(1)
        componente2 = self.invApp.ActiveDocument.SelectSet.Item(2)

        # Medir la distancia entre los componentes y comparar con valores dados
        distancia_caseta_arnes1 = self.medir_distancia(componente1, componente2)
        print("Distancia entre la caseta y el arnés 1:", distancia_caseta_arnes1, "mm")
        print("Valor dado: 12651.939 mm")

        # Realizar más mediciones y comparaciones según sea necesario
        
        # Desvincular el evento de selección
        self.root.unbind("<Button-1>")

    def medir_distancia(self, componente1, componente2):
        # Calcular la distancia entre los centros de masa de los componentes
        distancia = componente1.RangeBox.DistanceTo(componente2.RangeBox)
        return distancia

# Configurar y ejecutar la aplicación
root = tk.Tk()
app = AplicacionMedicion(root)
root.mainloop()
