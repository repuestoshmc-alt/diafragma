import re
import webbrowser
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import tkinter as tk
from openpyxl import load_workbook
from tkinter import messagebox, ttk

EXCEL_PATH = Path(__file__).resolve().parent / "diafragmas.xlsx"
COLUMNAS_REQUERIDAS = ["Marca", "Modelo", "Carburador", "Diafragma"]


def normalizar_texto(texto: str) -> str:
    if texto is None:
        return ""
    texto = str(texto).upper()
    return re.sub(r"[^A-Z0-9]", "", texto)


def extraer_numero_base(texto: str) -> str:
    normalizado = normalizar_texto(texto)
    match = re.search(r"\d+", normalizado)
    if not match:
        return ""
    numero = match.group(0)
    return str(int(numero)) if numero.isdigit() else numero


class DiafragmasApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Consulta de Diafragmas")
        self.root.geometry("760x620")
        self.root.minsize(720, 580)

        self.df = pd.DataFrame(columns=COLUMNAS_REQUERIDAS)

        self.var_marca = tk.StringVar()
        self.var_modelo = tk.StringVar()
        self.var_manual = tk.StringVar()

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
        contenedor = ttk.Frame(self.root, padding=16)
        contenedor.pack(fill="both", expand=True)

        panel_comb = ttk.LabelFrame(contenedor, text="Búsqueda por marca y modelo", padding=12)
        panel_comb.pack(fill="x", pady=(0, 10))

        ttk.Label(panel_comb, text="Marca:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.combo_marca = ttk.Combobox(panel_comb, textvariable=self.var_marca, state="readonly")
        self.combo_marca.grid(row=0, column=1, sticky="ew", pady=4)
        self.combo_marca.bind("<<ComboboxSelected>>", self._al_cambiar_marca)

        ttk.Label(panel_comb, text="Modelo:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.combo_modelo = ttk.Combobox(panel_comb, textvariable=self.var_modelo, state="readonly")
        self.combo_modelo.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Button(panel_comb, text="Buscar", command=self._buscar_desde_combos).grid(
            row=2, column=0, columnspan=2, pady=(8, 0), sticky="ew"
        )
        panel_comb.columnconfigure(1, weight=1)

        panel_manual = ttk.LabelFrame(contenedor, text="Búsqueda manual", padding=12)
        panel_manual.pack(fill="x", pady=(0, 10))

        self.entry_manual = ttk.Entry(panel_manual, textvariable=self.var_manual)
        self.entry_manual.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.entry_manual.bind("<Return>", lambda _e: self.buscar_modelo())

        ttk.Button(panel_manual, text="Buscar", command=lambda: self.buscar_modelo(entrada_manual=self.var_manual.get())).grid(row=0, column=1)
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

        self.panel_sin_resultados = ttk.LabelFrame(contenedor, text="Sin resultados", padding=12)

        ttk.Label(
            self.panel_sin_resultados,
            text="No se encontró el modelo en la planilla de consulta.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        ttk.Button(
            self.panel_sin_resultados,
            text="Buscar en ChatGPT",
            command=self.abrir_chatgpt,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ttk.Label(self.panel_sin_resultados, text="Agregar nuevo registro:").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        ttk.Label(self.panel_sin_resultados, text="Marca:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nueva_marca).grid(row=3, column=1, sticky="ew", pady=2)

        ttk.Label(self.panel_sin_resultados, text="Modelo:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_modelo).grid(row=4, column=1, sticky="ew", pady=2)

        ttk.Label(self.panel_sin_resultados, text="Carburador:").grid(row=5, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_carburador).grid(row=5, column=1, sticky="ew", pady=2)

        ttk.Label(self.panel_sin_resultados, text="Diafragma:").grid(row=6, column=0, sticky="w", pady=2)
        ttk.Entry(self.panel_sin_resultados, textvariable=self.var_nuevo_diafragma).grid(row=6, column=1, sticky="ew", pady=2)

        ttk.Button(
            self.panel_sin_resultados,
            text="Guardar en Excel",
            command=lambda: (print("Botón guardar presionado"), self.guardar_nuevo_registro()),
        ).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))
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

    def _buscar_desde_combos(self) -> None:
        marca = self.var_marca.get().strip()
        modelo = self.var_modelo.get().strip()

        if not marca or not modelo:
            messagebox.showwarning("Atención", "Seleccione marca y modelo para buscar.")
            return

        resultados = self.buscar_modelo(marca=marca, modelo=modelo, entrada_manual="")
        if resultados:
            self.mostrar_resultados(resultados)
        else:
            self.mostrar_sin_resultados(marca, modelo)

    def buscar_modelo(self, marca: str = "", modelo: str = "", entrada_manual: str | None = None):
        if self.df.empty:
            return []

        busqueda_manual = entrada_manual if entrada_manual is not None else self.var_manual.get()

        marca_input = marca.strip()
        modelo_input = modelo.strip()
        manual_input = busqueda_manual.strip()

        resultados = []

        if manual_input:
            texto = manual_input
            texto_normalizado = normalizar_texto(texto)
            numero_base = extraer_numero_base(texto)

            print("Texto buscado:", texto)
            print("Texto normalizado:", texto_normalizado)
            print("Número base:", numero_base)

            es_busqueda_numerica = texto_normalizado.isdigit()

            for _, fila in self.df.iterrows():
                fila_dict = {col: str(fila[col]).strip() for col in COLUMNAS_REQUERIDAS}
                modelo_fila = fila_dict["Modelo"]
                modelo_normalizado = normalizar_texto(modelo_fila)
                numero_base_modelo = extraer_numero_base(modelo_fila)

                if es_busqueda_numerica:
                    if numero_base and numero_base == numero_base_modelo:
                        resultados.append(fila_dict)
                else:
                    if texto_normalizado and texto_normalizado == modelo_normalizado:
                        resultados.append(fila_dict)

            print("Cantidad de resultados:", len(resultados))
        else:
            marca_norm = normalizar_texto(marca_input)
            modelo_norm = normalizar_texto(modelo_input)

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
            self.mostrar_sin_resultados(marca_input, modelo_input or manual_input)

        return resultados

    def mostrar_resultados(self, resultados) -> None:
        if not resultados:
            self._limpiar_resultados()
            return

        primero = resultados[0]
        self.var_resultado_marca.set(primero["Marca"])
        self.var_resultado_modelo.set(primero["Modelo"])
        self.var_resultado_carburador.set(primero["Carburador"])
        self.var_resultado_diafragma.set(primero["Diafragma"])

        self.ocultar_panel_sin_resultados()

    def mostrar_sin_resultados(self, marca: str = "", modelo: str = "") -> None:
        self._limpiar_resultados()
        if marca and not self.var_nueva_marca.get().strip():
            self.var_nueva_marca.set(marca)
        if modelo and not self.var_nuevo_modelo.get().strip():
            self.var_nuevo_modelo.set(modelo)
        print("Mostrando panel sin resultados")
        self.panel_sin_resultados.pack(fill="x", pady=(0, 10))

    def ocultar_panel_sin_resultados(self) -> None:
        self.panel_sin_resultados.pack_forget()

    def abrir_chatgpt(self) -> None:
        marca = self.var_nueva_marca.get().strip() or self.var_marca.get().strip() or "(sin marca)"
        modelo = self.var_nuevo_modelo.get().strip() or self.var_modelo.get().strip() or self.var_manual.get().strip() or "(sin modelo)"

        consulta = f"Buscar diafragma para:\n\nMarca: {marca}\nModelo: {modelo}"
        encoded = quote(consulta)
        url = f"https://chatgpt.com/?q={encoded}"
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
        print("Registro agregado correctamente")
        messagebox.showinfo("OK", "Registro agregado correctamente")

    def _limpiar_resultados(self) -> None:
        self.var_resultado_marca.set("-")
        self.var_resultado_modelo.set("-")
        self.var_resultado_carburador.set("-")
        self.var_resultado_diafragma.set("-")


def main() -> None:
    root = tk.Tk()
    app = DiafragmasApp(root)
    _ = app
    root.mainloop()


if __name__ == "__main__":
    main()
