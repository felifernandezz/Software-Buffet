
import pandas as pd
import os
import pywhatkit as kit
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import messagebox, simpledialog

# Configurar el idioma en español
locale.setlocale(locale.LC_TIME, "Spanish_Spain.1252")

# Crear archivo Excel si no existe
def crear_archivo_excel():
    archivo = "registro_bufet.xlsx"
    if not os.path.exists(archivo):
        df = pd.DataFrame(columns=["Fecha", "Alumno", "Precio"])
        df.to_excel(archivo, index=False)

# Registrar un consumo
def registrar_consumo():
    fecha = simpledialog.askstring("Fecha", "Ingresa la fecha (DD-MM):")
    alumno = simpledialog.askstring("Alumno", "Ingresa el nombre del alumno:")
    precio = simpledialog.askfloat("Precio", "Ingresa el precio:")

    if fecha and alumno and precio:
        archivo = "registro_bufet.xlsx"
        df = pd.read_excel(archivo)
        nueva_fila = pd.DataFrame({"Fecha": [fecha], "Alumno": [alumno], "Precio": [precio]})
        df = pd.concat([df, nueva_fila], ignore_index=True)
        df.to_excel(archivo, index=False)
        messagebox.showinfo("Registro", "✅ Consumo registrado exitosamente.")

# Calcular liquidación semanal
def calcular_liquidacion():
    archivo = "registro_bufet.xlsx"
    df = pd.read_excel(archivo)

    # Agrupar por alumno y sumar lo que debe
    liquidacion = df.groupby("Alumno")["Precio"].sum().reset_index()
    liquidacion.columns = ["Alumno", "Total a Pagar"]

    # Guardar la liquidación
    liquidacion.to_excel("liquidacion_semanal.xlsx", index=False)
    messagebox.showinfo("Liquidación", "✅ Liquidación semanal generada correctamente.")

# Enviar liquidación por WhatsApp
def enviar_liquidacion_whatsapp():
    archivo = "liquidacion_semanal.xlsx"
    if not os.path.exists(archivo):
        messagebox.showwarning("Error", "⚠ Primero debes calcular la liquidación semanal.")
        return

    df = pd.read_excel(archivo)
    for index, row in df.iterrows():
        alumno = row["Alumno"]
        total_pagar = row["Total a Pagar"]
        telefono = simpledialog.askstring("Número", f"Ingrese el número de WhatsApp del padre de {alumno}:")

        if telefono:
            mensaje = f"Hola, soy Silvana del bufet escolar. La liquidación de {alumno} es ${total_pagar:.2f}. Gracias."
            try:
                kit.sendwhatmsg_instantly(telefono, mensaje, wait_time=10)
                messagebox.showinfo("Enviado", f"✅ Liquidación enviada a {telefono}")
            except Exception as e:
                messagebox.showerror("Error", f"❌ No se pudo enviar a {telefono}: {e}")

# Reiniciar liquidación y eliminar registros de consumo
def reiniciar_liquidacion():
    archivo_liquidacion = "liquidacion_semanal.xlsx"
    archivo_registro = "registro_bufet.xlsx"

    if not os.path.exists(archivo_liquidacion):
        messagebox.showwarning("Error", "⚠ No hay liquidaciones registradas.")
        return

    alumno = simpledialog.askstring("Reiniciar Liquidación", "Ingresa el nombre del alumno que pagó:")

    if alumno:
        df_liquidacion = pd.read_excel(archivo_liquidacion)
        df_registro = pd.read_excel(archivo_registro)

        if alumno in df_liquidacion["Alumno"].values:
            # Eliminar los registros de consumo del alumno
            df_registro = df_registro[df_registro["Alumno"] != alumno]
            df_registro.to_excel(archivo_registro, index=False)

            # Poner su deuda en $0 en liquidación
            df_liquidacion.loc[df_liquidacion["Alumno"] == alumno, "Total a Pagar"] = 0
            df_liquidacion.to_excel(archivo_liquidacion, index=False)

            messagebox.showinfo("Reiniciado", f"✅ La deuda de {alumno} ha sido eliminada y sus registros borrados.")
        else:
            messagebox.showwarning("No encontrado", "⚠ Alumno no encontrado en la liquidación.")

# Salir del programa
def salir():
    ventana.destroy()

# Crear ventana
ventana = tk.Tk()
ventana.title("Menú del Sistema")
ventana.geometry("400x400")

# Etiqueta de título
titulo = tk.Label(ventana, text="📌 MENÚ DEL SISTEMA", font=("Arial", 14, "bold"))
titulo.pack(pady=10)

# Botones del menú
botones = [
    ("Registrar Consumo", registrar_consumo),
    ("Calcular Liquidación", calcular_liquidacion),
    ("Enviar Liquidaciones por WhatsApp", enviar_liquidacion_whatsapp),
    ("Reiniciar Deuda de un Alumno", reiniciar_liquidacion),
    ("Salir", salir),
]

for texto, comando in botones:
    btn = tk.Button(ventana, text=texto, font=("Arial", 12), width=30, height=2, command=comando)
    btn.pack(pady=5)

# Ejecutar la ventana
ventana.mainloop()

# Ejecutar el programa
if __name__ == "__main__":
    crear_archivo_excel()
