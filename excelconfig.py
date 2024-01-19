import pandas as pd
from tkinter import Tk, filedialog, Button, Label

def procesar_excel():
    # Abrir el cuadro de diálogo para seleccionar el archivo Excel
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])

    # Verificar si se seleccionó un archivo
    if not file_path:
        return

    # Leer el archivo Excel
    df = pd.read_excel(file_path)

    # Crear un nuevo DataFrame con las columnas requeridas
    new_df = df[['NOTA PEDIDO', 'NOMBRE DEL CLIENTE', 'CANT', 'PRODUCTOS']]

    # Buscar el precio unitario en ambas columnas (E e I)
    precio_unitario = df['PRECIO UNIT'] if 'PRECIO UNIT' in df.columns else df['PRECIO UNIT ']

    # Añadir la columna de Precio Unit y Total
    new_df['PRECIO UNIT'] = precio_unitario.fillna(0.00)
    new_df['TOTAL'] = new_df['CANT'] * new_df['PRECIO UNIT']

    # Guardar el nuevo DataFrame en un nuevo archivo Excel
    new_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if new_file_path:
        new_df.to_excel(new_file_path, index=False)
        Label(root, text="Proceso completado. Nuevo archivo guardado.").pack()

# Crear la interfaz gráfica
root = Tk()
root.title("Procesar Excel")
root.geometry("300x100")

# Crear un botón para iniciar el proceso
Button(root, text="Iniciar", command=procesar_excel).pack(pady=20)

root.mainloop()
