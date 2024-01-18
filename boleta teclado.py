import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import json
import pyautogui
import time

def ejecutar_codigo(datos_para_insertar):
    # Coordenadas de la consola
    x_clic_adicional_1 = 919
    y_clic_adicional_1 = 263

    
    pyautogui.click(x_clic_adicional_1, y_clic_adicional_1, duration=1)

    pyautogui.press('tab')      
    pyautogui.press('tab')
    for datos in datos_para_insertar:

        if datos["IGV"] == 'no':
                precio_unitario_ajustado = (datos["precio"])
        else:
                precio_unitario_ajustado = round(datos["precio"] / 1.18, 10)
        
        # Presiona la tecla Tab
        pyautogui.press('enter')
        time.sleep(2)

        pyautogui.press('right')
        # pyautogui.click(x_clic_adicional_2, y_clic_adicional_2, duration=1)
        pyautogui.press('tab') 
        pyautogui.press('1')

        pyautogui.press('tab') 
        pyautogui.write(datos["nombre"])

        pyautogui.press('tab') 
        pyautogui.press('tab') 
        pyautogui.press('backspace')
        pyautogui.press('backspace')
        pyautogui.press('backspace')
        pyautogui.press('backspace')
        pyautogui.write(str(precio_unitario_ajustado))

        pyautogui.press('tab')
        pyautogui.press('tab')

        if datos["IGV"] == 'no':
                pyautogui.press('right')
        
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')


        pyautogui.write(datos["unidad"])

        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
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

def iniciar_boleteo(tabla, ventana_resultado):
    # Obtén los datos actuales de la tabla
    datos_actuales = []
    for item_id in tabla.get_children():
        datos = [tabla.item(item_id, 'values')[i] for i in range(4)]
        datos_actuales.append({'unidad': datos[0], 'nombre': datos[1], 'precio': float(datos[2]), 'IGV': datos[3]})

    # Llama a la función ejecutar_codigo con los datos actuales
    ejecutar_codigo(datos_actuales)

    # Muestra la ventana de resultado
    ventana_resultado.deiconify()

def cerrar_ventana(ventana_resultado):
    # Cierra la ventana de resultado
    ventana_resultado.destroy()

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

    def crear_ventana_resultado():
        # Crea una nueva ventana para mostrar el resultado
        ventana_resultado = tk.Toplevel(ventana_seleccion)
        ventana_resultado.title("Resultado")

        # Etiqueta y botón de "Listo"
        label_listo = tk.Label(ventana_resultado, text="Listo")
        label_listo.pack(pady=10)

        boton_cerrar = tk.Button(ventana_resultado, text="Cerrar", command=lambda: cerrar_ventana(ventana_resultado))
        boton_cerrar.pack(pady=10)

        # Oculta la ventana principal durante la ejecución
        ventana_seleccion.withdraw()

        # Inicia el boleteo con la nueva ventana de resultado
        iniciar_boleteo(tabla, ventana_resultado)

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
    tabla["columns"] = ("Unidades", "Nombres", "Precios", "IGV")
    tabla["show"] = "headings"

    # Configurar las columnas
    for columna in tabla["columns"]:
        tabla.heading(columna, text=columna)
        tabla.column(columna, anchor="center")

    # Mostrar la tabla
    tabla.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    # Botón para iniciar el boleteo
    boton_iniciar_boleteo = tk.Button(ventana_seleccion, text="Iniciar Boleteo", command=crear_ventana_resultado)
    boton_iniciar_boleteo.grid(row=2, column=1, pady=10)


    # Ejecuta el bucle de la interfaz gráfica
    ventana_seleccion.mainloop()

if __name__ == "__main__":
    # Crear una tabla para pasar como argumento a la función
    tabla_inicial = tk.Tk()
    tabla_inicial.withdraw()
    seleccionar_archivo_excel(tabla_inicial)
