import re
import webbrowser
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import tkinter as tk

from openpyxl import load_workbook
from PIL import Image, ImageTk
from tkinter import ttk, messagebox

EXCEL_PATH = Path(__file__).resolve().parent / "diafragmas.xlsx"
IMAGES_PATH = Path(__file__).resolve().parent / "imagenes"
COLUMNAS_REQUERIDAS = ["Marca", "Modelo", "Carburador", "Diafragma"]


def normalizar_texto(texto: str) -> str:
    if texto is None:
        return ""
    texto = str(texto).upper()
    return re.sub(r"[^A-Z0-9]", "", texto)


def extraer_numero_base(texto: str) -> str:
    texto_norm = normalizar_texto(texto)
    match = re.search(r"\d+", texto_norm)
    return match.group(0) if match else ""


class DiafragmasApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Consulta de Diafragmas")
        self.root.geometry("780x670")
        self.root.minsize(740, 620)

        self.df = pd.DataFrame(columns=COLUMNAS_REQUERIDAS)
        self.imagen_actual = None

        self.var_marca = tk.StringVar()
        self.var_modelo = tk.StringVar()
        self.var_busqueda_manual = tk.StringVar()

        self.var_resultado_marca = tk.StringVar(value="-")
        self.var_resultado_modelo = tk.StringVar(value="-")
        self.var_resultado_carburador = tk.StringVar(value="-")
        self.var_resultado_diafragma = tk.StringVar(value="-")

        self.var_nueva_marca = tk.StringVar()
        self.var_nuevo_modelo = tk.StringVar()
        self.var_nuevo_carburador = tk.StringVar()
        self.var_nuevo_diafragma = tk.StringVar()

        self._construir_interfaz()
        self.carregar_excel()

    def _construir_interfaz(self) -> None:
        contenedor = ttk.Frame(self.root, padding=14)
        contenedor.pack(fill="both", expand=True)

        panel_combos = ttk.LabelFrame(contenedor, text="Búsqueda por marca y modelo", padding=12)
        panel_combos.pack(fill="x", pady=(0, 10))

        ttk.Label(panel_combos, text="Marca:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.combo_marca = ttk.Combobox(panel_combos, textvariable=self.var_marca, state="readonly")
        self.combo_marca.grid(row=0, column=1, sticky="ew", pady=4)
        self.combo_marca.bind("<<ComboboxSelected>>", self._al_cambiar_marca)

        ttk.Label(panel_combos, text="Modelo:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.combo_modelo = ttk.Combobox(panel_combos, textvariable=self.var_modelo, state="readonly")
        self.combo_modelo.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Button(panel_combos, text="Buscar", command=self._buscar_por_combos).grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        panel_combos.columnconfigure(1, weight=1)

        panel_manual = ttk.LabelFrame(contenedor, text="Búsqueda manual", padding=12)
        panel_manual.pack(fill="x", pady=(0, 10))

        entry_manual = ttk.Entry(panel_manual, textvariable=self.var_busqueda_manual)
        entry_manual.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        entry_manual.bind("<Return>", lambda _e: self.buscar_modelo())

        ttk.Button(panel_manual, text="Buscar", command=self.buscar_modelo).grid(row=0, column=1)
        panel_manual.columnconfigure(0, weight=1)

        panel_res = ttk.LabelFrame(contenedor, text="Resultados", padding=12)
        panel_res.pack(fill="x", pady=(0, 10))

        ttk.Label(panel_res, text="Marca:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(panel_res, textvariable=self.var_resultado_marca).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(panel_res, text="Modelo:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(panel_res, textvariable=self.var_resultado_modelo).grid(row=1, column=1, sticky="w", pady=2)
        ttk.Label(panel_res, text="Carburador:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(panel_res, textvariable=self.var_resultado_carburador).grid(row=2, column=1, sticky="w", pady=2)
        ttk.Label(panel_res, text="Diafragma:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Label(panel_res, textvariable=self.var_resultado_diafragma).grid(row=3, column=1, sticky="w", pady=2)

        self.label_imagen = ttk.Label(panel_res, text="Imagen no disponible")
        self.label_imagen.grid(row=4, column=0, columnspan=2, pady=(15, 5))

        self.panel_sin_resultados = ttk.LabelFrame(contenedor, text="Sin resultados", padding=12)

        ttk.Label(self.panel_sin_resultados, text="No se encontró el modelo").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )
        ttk.Button(self.panel_sin_resultados, text="Buscar en ChatGPT", command=self.abrir_chatgpt).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(self.panel_sin_resultados, text="Marca:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nueva_marca).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(self.panel_sin_resultados, text="Modelo:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_modelo).grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Label(self.panel_sin_resultados, text="Carburador:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_carburador).grid(row=4, column=1, sticky="ew", pady=2)
        ttk.Label(self.panel_sin_resultados, text="Diafragma:").grid(row=5, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_diafragma).grid(row=5, column=1, sticky="ew", pady=2)

        ttk.Button(self.panel_sin_resultados, text="Guardar en Excel", command=self.guardar_nuevo_registro).grid(
            row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0)
        )
        self.panel_sin_resultados.columnconfigure(1, weight=1)
        self.ocultar_panel_sin_resultados()

    def carregar_excel(self) -> None:
        if not EXCEL_PATH.exists():
            messagebox.showerror("Error", f"No existe el archivo Excel: {EXCEL_PATH.name}")
            return
        try:
            df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
        except PermissionError:
            messagebox.showerror("Error", "No se puede abrir el Excel. Verifique que no esté abierto.")
            return
        except Exception as exc:
            messagebox.showerror("Error", f"Error al leer el Excel: {exc}")
            return

        faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
        if faltantes:
            messagebox.showerror("Error", f"Faltan columnas requeridas: {', '.join(faltantes)}")
            return

        self.df = df[COLUMNAS_REQUERIDAS].fillna("").copy()
        self._actualizar_combos()

    def _actualizar_combos(self) -> None:
        marcas = sorted({str(m).strip() for m in self.df["Marca"] if str(m).strip()})
        self.combo_marca["values"] = marcas
        self.combo_modelo["values"] = []
        self.var_marca.set("")
        self.var_modelo.set("")

    def _al_cambiar_marca(self, _event=None) -> None:
        marca = self.var_marca.get().strip()
        if not marca:
            self.combo_modelo["values"] = []
            self.var_modelo.set("")
            return
        subset = self.df[self.df["Marca"].astype(str).str.strip() == marca]
        modelos = sorted({str(m).strip() for m in subset["Modelo"] if str(m).strip()})
        self.combo_modelo["values"] = modelos
        self.var_modelo.set("")

    def _buscar_por_combos(self) -> None:
        marca = self.var_marca.get().strip()
        modelo = self.var_modelo.get().strip()
        if not marca or not modelo:
            messagebox.showwarning("Atención", "Seleccione marca y modelo para buscar.")
            return

        marca_norm = normalizar_texto(marca)
        modelo_norm = normalizar_texto(modelo)
        resultados = []
        for _, fila in self.df.iterrows():
            fila_dict = {col: str(fila[col]).strip() for col in COLUMNAS_REQUERIDAS}
            if (
                normalizar_texto(fila_dict["Marca"]) == marca_norm
                and normalizar_texto(fila_dict["Modelo"]) == modelo_norm
            ):
                resultados.append(fila_dict)

        if resultados:
            self.mostrar_resultados(resultados)
        else:
            self.mostrar_sin_resultados(marca, modelo)

    def buscar_modelo(self) -> None:
        texto = self.var_busqueda_manual.get().strip()
        if not texto:
            messagebox.showwarning("Atención", "Ingrese un modelo para buscar.")
            return

        texto_norm = normalizar_texto(texto)
        numero_busqueda = extraer_numero_base(texto)
        es_numerica = texto_norm.isdigit()

        resultados = []
        for _, fila in self.df.iterrows():
            fila_dict = {col: str(fila[col]).strip() for col in COLUMNAS_REQUERIDAS}
            modelo = fila_dict["Modelo"]
            modelo_norm = normalizar_texto(modelo)
            numero_modelo = extraer_numero_base(modelo)

            if es_numerica:
                if numero_busqueda and numero_busqueda == numero_modelo:
                    resultados.append(fila_dict)
            else:
                if texto_norm and texto_norm == modelo_norm:
                    resultados.append(fila_dict)

        if resultados:
            self.mostrar_resultados(resultados)
        else:
            self.mostrar_sin_resultados("", texto)

    def mostrar_resultados(self, resultados) -> None:
        primero = resultados[0]
        self.var_resultado_marca.set(primero["Marca"])
        self.var_resultado_modelo.set(primero["Modelo"])
        self.var_resultado_carburador.set(primero["Carburador"])
        self.var_resultado_diafragma.set(primero["Diafragma"])
        self.mostrar_imagen_diafragma(primero["Diafragma"])
        self.ocultar_panel_sin_resultados()

    def mostrar_sin_resultados(self, marca: str = "", modelo: str = "") -> None:
        self.limpiar_imagen()
        self.var_resultado_marca.set("-")
        self.var_resultado_modelo.set("-")
        self.var_resultado_carburador.set("-")
        self.var_resultado_diafragma.set("-")
        if marca and not self.var_nueva_marca.get().strip():
            self.var_nueva_marca.set(marca)
        if modelo and not self.var_nuevo_modelo.get().strip():
            self.var_nuevo_modelo.set(modelo)
        self.panel_sin_resultados.pack(fill="x", pady=(0, 10))

    def ocultar_panel_sin_resultados(self) -> None:
        self.panel_sin_resultados.pack_forget()

    def mostrar_imagen_diafragma(self, codigo: str) -> None:
        print("Código recibido:", codigo)
        codigo = str(codigo).strip()
        if not codigo:
            print("No existe imagen")
            self.limpiar_imagen()
            return

        for ext in (".jpg", ".jpeg", ".png"):
            ruta = IMAGES_PATH / f"{codigo}{ext}"
            print("Ruta probada:", ruta)
            print("Existe:", ruta.exists())
            if not ruta.exists():
                continue
            try:
                img = Image.open(ruta)
                img.thumbnail((250, 180))
                foto = ImageTk.PhotoImage(img)
                self.label_imagen.configure(image=foto, text="")
                self.label_imagen.image = foto
                self.imagen_actual = foto
                print("Imagen cargada OK")
                return
            except Exception as exc:
                print("Error cargando:", exc)
                continue

        print("No existe imagen")
        self.limpiar_imagen()

    def limpiar_imagen(self) -> None:
        self.label_imagen.configure(image="", text="Imagen no disponible")
        self.label_imagen.image = None
        self.imagen_actual = None

    def abrir_chatgpt(self) -> None:
        marca = self.var_nueva_marca.get().strip() or self.var_marca.get().strip() or "(sin marca)"
        modelo = self.var_nuevo_modelo.get().strip() or self.var_modelo.get().strip() or self.var_busqueda_manual.get().strip() or "(sin modelo)"
        consulta = f"Buscar diafragma para:\n\nMarca: {marca}\nModelo: {modelo}"
        url = f"https://chatgpt.com/?q={quote(consulta)}"
        webbrowser.open_new_tab(url)

    def guardar_nuevo_registro(self) -> None:
        print("Guardando nuevo registro")

        marca = self.var_nueva_marca.get().strip()
        modelo = self.var_nuevo_modelo.get().strip()
        carburador = self.var_nuevo_carburador.get().strip()
        diafragma = self.var_nuevo_diafragma.get().strip()

        if not all([marca, modelo, carburador, diafragma]):
            messagebox.showwarning("Atención", "Complete todos los campos para guardar el registro.")
            return

        clave_nueva = f"{normalizar_texto(marca)}|{normalizar_texto(modelo)}"
        for _, fila in self.df.iterrows():
            clave_actual = f"{normalizar_texto(fila['Marca'])}|{normalizar_texto(fila['Modelo'])}"
            if clave_actual == clave_nueva:
                messagebox.showwarning("Duplicado", "Ese modelo ya existe en la base.")
                return

        if not EXCEL_PATH.exists():
            messagebox.showerror("Error", f"No existe el archivo Excel: {EXCEL_PATH.name}")
            return

        try:
            workbook = load_workbook(EXCEL_PATH)
            worksheet = workbook.active
            headers = [cell.value for cell in worksheet[1]]
            faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in headers]
            if faltantes:
                messagebox.showerror("Error", f"Faltan columnas requeridas: {', '.join(faltantes)}")
                workbook.close()
                return

            worksheet.append([marca, modelo, carburador, diafragma])
            workbook.save(EXCEL_PATH)
            workbook.close()
        except PermissionError:
            messagebox.showerror("Error", "No se puede guardar. Cierre el archivo Excel si está abierto.")
            return
        except Exception as exc:
            messagebox.showerror("Error", f"Error al guardar en Excel: {exc}")
            return

        self.carregar_excel()
        self.var_nueva_marca.set("")
        self.var_nuevo_modelo.set("")
        self.var_nuevo_carburador.set("")
        self.var_nuevo_diafragma.set("")
        self.ocultar_panel_sin_resultados()
        messagebox.showinfo("OK", "Registro agregado correctamente")


def main() -> None:
    root = tk.Tk()
    DiafragmasApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
