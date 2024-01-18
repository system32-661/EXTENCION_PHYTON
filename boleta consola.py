import tkinter as tk
import pyautogui
import time
import pandas as pd
import json

from tkinter import ttk
from tkinter import filedialog

def ejecutar_codigo(datos_para_insertar):
    # Coordenadas de la consola
    x_clic_adicional_2 = 1115
    y_clic_adicional_2 = 702

    for datos in datos_para_insertar:
        precio_unitario_ajustado = round(datos["precio"] / 1.18, 10)

        javascript_code = f'''
        document.getElementById('item.subTipoTI02').click();
        document.getElementById('item.cantidad').value = '{datos["unidad"]}';
        document.getElementById('item.codigoItem').value = '1';
        document.getElementById('item.descripcion').value = '{datos["nombre"]}';
        document.getElementById('item.precioUnitario').value = '{precio_unitario_ajustado}';
        '''

        if datos["IGV"] == 'no':
            javascript_code += '''
                document.getElementById('item.subTipoTB01').click();
            '''

        javascript_code += '''
            document.getElementById('item.botonAceptar').click();
        '''

        pyautogui.click(x_clic_adicional_2, y_clic_adicional_2, duration=1)
        
        time.sleep(1)

        pyautogui.write(javascript_code)
        pyautogui.press('enter')

        time.sleep(2)

def leer_excel(nombre_archivo):
    try:
        # Lee el archivo Excel
        datos = pd.read_excel(nombre_archivo)
        return datos
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return None

def actualizar_tabla(tabla, datos):
    # Limpia la tabla existente
    for row in tabla.get_children():
        tabla.delete(row)

    # Llena la tabla con los nuevos datos del DataFrame
    for fila in datos.itertuples(index=False):
        tabla.insert("", "end", values=tuple(fila))

def iniciar_boleteo(tabla):
    # Obtén los datos actuales de la tabla
    datos_actuales = []
    for item_id in tabla.get_children():
        datos = [tabla.item(item_id, 'values')[i] for i in range(4)]
        datos_actuales.append({'unidad': datos[0], 'nombre': datos[1], 'precio': float(datos[2]), 'IGV': datos[3]})

    # Imprime el JSON con los datos actuales
    json_datos = json.dumps(datos_actuales, indent=2)
    print("Datos actuales en formato JSON:")
    print(json_datos)
    # Llama a la función ejecutar_codigo con los datos actuales
    ejecutar_codigo(datos_actuales)

def seleccionar_archivo_excel(tabla):
    def seleccionar_archivo():
        # Abre el cuadro de diálogo para seleccionar el archivo
        archivo_excel = filedialog.askopenfilename(
            title="Seleccionar Archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo_excel)

        # Lee los datos del archivo Excel
        datos = leer_excel(archivo_excel)

        # Si se pudieron leer los datos, actualiza la tabla
        if datos is not None:
            actualizar_tabla(tabla, datos)

    # Crea una ventana tkinter para la selección de archivo
    ventana_seleccion = tk.Tk()
    ventana_seleccion.title("Selección de Archivo Excel")

    # Etiqueta y entrada para mostrar el nombre del archivo
    label_archivo = tk.Label(ventana_seleccion, text="Archivo Excel:")
    label_archivo.grid(row=0, column=0, padx=10, pady=10)

    entry_archivo = tk.Entry(ventana_seleccion, width=50)
    entry_archivo.grid(row=0, column=1, padx=10, pady=10)

    # Botón para seleccionar archivo
    boton_seleccionar = tk.Button(ventana_seleccion, text="Seleccionar Archivo", command=seleccionar_archivo)
    boton_seleccionar.grid(row=0, column=2, padx=10, pady=10)

    # Crear una tabla para mostrar los datos
    tabla = ttk.Treeview(ventana_seleccion)
    tabla["columns"] = ("Unidad", "Nombre", "Precio", "IGV")
    tabla["show"] = "headings"

    # Configurar las columnas
    for columna in tabla["columns"]:
        tabla.heading(columna, text=columna)
        tabla.column(columna, anchor="center")

    # Mostrar la tabla
    tabla.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    # Botón para iniciar el boleteo
    boton_iniciar_boleteo = tk.Button(ventana_seleccion, text="Iniciar Boleteo", command=lambda: iniciar_boleteo(tabla))
    boton_iniciar_boleteo.grid(row=2, column=1, pady=10)

    # Ejecuta el bucle de la interfaz gráfica
    ventana_seleccion.mainloop()

if __name__ == "__main__":
    # Crear una tabla para pasar como argumento a la función
    tabla_inicial = tk.Tk()
    tabla_inicial.withdraw()
    seleccionar_archivo_excel(tabla_inicial)