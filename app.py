import os
import re
import unicodedata
import urllib.parse
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd
from openpyxl import load_workbook
from rapidfuzz import fuzz

EXCEL_FILE = "diafragmas.xlsx"
COLUMNAS_REQUERIDAS = ["Marca", "Modelo", "Carburador", "Diafragma"]


def normalizar_texto(valor: str) -> str:
    """Normaliza texto ignorando acentos, espacios, guiones y mayúsculas/minúsculas."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"[\s\-]+", "", texto)


class DiafragmaApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Consulta Inteligente de Diafragmas")
        self.root.geometry("900x620")
        self.root.minsize(820, 560)

        self.df = pd.DataFrame(columns=COLUMNAS_REQUERIDAS)
        self.resultados_actuales: list[dict] = []

        self.configurar_estilo()
        self.construir_ui()
        self.cargar_excel()

    def configurar_estilo(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.root.configure(bg="#eef2f7")

        style.configure("TFrame", background="#eef2f7")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), background="#eef2f7")
        style.configure("Hint.TLabel", font=("Segoe UI", 10), background="#eef2f7", foreground="#596579")
        style.configure("TLabel", font=("Segoe UI", 10), background="#eef2f7")
        style.configure("TLabelframe", background="#eef2f7")
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), background="#eef2f7")
        style.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.configure("Info.TLabel", font=("Segoe UI", 11, "bold"), background="#ffffff")

    def construir_ui(self) -> None:
        contenedor = ttk.Frame(self.root, padding=16)
        contenedor.pack(fill="both", expand=True)

        ttk.Label(contenedor, text="Consulta Inteligente de Diafragmas", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            contenedor,
            text="Busca por marca/modelo o por texto libre. Coincidencias exactas y similares con RapidFuzz.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 14))

        panel_sel = ttk.LabelFrame(contenedor, text="Búsqueda por desplegables")
        panel_sel.pack(fill="x", pady=(0, 10))

        ttk.Label(panel_sel, text="Marca:").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        self.var_marca = tk.StringVar()
        self.cmb_marca = ttk.Combobox(panel_sel, textvariable=self.var_marca, state="readonly", width=35)
        self.cmb_marca.grid(row=0, column=1, sticky="ew", padx=10, pady=8)
        self.cmb_marca.bind("<<ComboboxSelected>>", self.al_elegir_marca)

        ttk.Label(panel_sel, text="Modelo:").grid(row=1, column=0, sticky="w", padx=10, pady=8)
        self.var_modelo = tk.StringVar()
        self.cmb_modelo = ttk.Combobox(panel_sel, textvariable=self.var_modelo, state="readonly", width=35)
        self.cmb_modelo.grid(row=1, column=1, sticky="ew", padx=10, pady=8)
        self.cmb_modelo.bind("<<ComboboxSelected>>", self.al_elegir_modelo)
        panel_sel.columnconfigure(1, weight=1)

        panel_busq = ttk.LabelFrame(contenedor, text="Búsqueda libre inteligente")
        panel_busq.pack(fill="both", expand=True, pady=(0, 10))

        self.var_busqueda = tk.StringVar()
        self.entry_busqueda = ttk.Entry(panel_busq, textvariable=self.var_busqueda)
        self.entry_busqueda.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        self.entry_busqueda.bind("<Return>", lambda _e: self.buscar_modelo())

        ttk.Button(panel_busq, text="Buscar", style="Big.TButton", command=self.buscar_modelo).grid(
            row=0, column=1, padx=(0, 10), pady=8
        )

        self.lista_resultados = tk.Listbox(panel_busq, height=9, font=("Segoe UI", 10), activestyle="none")
        self.lista_resultados.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=(10, 0), pady=(0, 10))
        self.lista_resultados.bind("<<ListboxSelect>>", self.al_seleccionar_resultado)

        scrollbar = ttk.Scrollbar(panel_busq, orient="vertical", command=self.lista_resultados.yview)
        scrollbar.grid(row=1, column=2, sticky="ns", pady=(0, 10), padx=(0, 10))
        self.lista_resultados.config(yscrollcommand=scrollbar.set)

        panel_busq.columnconfigure(0, weight=1)
        panel_busq.rowconfigure(1, weight=1)

        self.var_estado = tk.StringVar(value="")
        self.lbl_estado = ttk.Label(contenedor, textvariable=self.var_estado, foreground="#b42318")
        self.lbl_estado.pack(anchor="w", pady=(0, 6))

        self.btn_chatgpt = ttk.Button(
            contenedor, text="Buscar en ChatGPT", style="Big.TButton", command=self.abrir_chatgpt
        )

        panel_detalle = ttk.Frame(contenedor, style="Card.TFrame", padding=14)
        panel_detalle.pack(fill="x", pady=(0, 10))

        ttk.Label(panel_detalle, text="Carburador:").grid(row=0, column=0, sticky="w", pady=4)
        self.lbl_carburador = ttk.Label(panel_detalle, text="-", style="Info.TLabel")
        self.lbl_carburador.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(panel_detalle, text="Diafragma:").grid(row=1, column=0, sticky="w", pady=4)
        self.lbl_diafragma = ttk.Label(panel_detalle, text="-", style="Info.TLabel")
        self.lbl_diafragma.grid(row=1, column=1, sticky="w", pady=4)

        self.frm_nuevo = ttk.LabelFrame(contenedor, text="Agregar nuevo registro")
        self.var_nueva_marca = tk.StringVar()
        self.var_nuevo_modelo = tk.StringVar()
        self.var_nuevo_carburador = tk.StringVar()
        self.var_nuevo_diafragma = tk.StringVar()

        campos = [
            ("Marca", self.var_nueva_marca),
            ("Modelo", self.var_nuevo_modelo),
            ("Carburador", self.var_nuevo_carburador),
            ("Diafragma", self.var_nuevo_diafragma),
        ]
        for idx, (texto, var) in enumerate(campos):
            ttk.Label(self.frm_nuevo, text=f"{texto}:").grid(row=idx // 2, column=(idx % 2) * 2, sticky="w", padx=10, pady=8)
            ttk.Entry(self.frm_nuevo, textvariable=var, width=30).grid(
                row=idx // 2, column=(idx % 2) * 2 + 1, sticky="ew", padx=10, pady=8
            )
        ttk.Button(self.frm_nuevo, text="Guardar en Excel", style="Big.TButton", command=self.guardar_nuevo_registro).grid(
            row=2, column=0, columnspan=4, pady=(4, 10)
        )
        self.frm_nuevo.columnconfigure(1, weight=1)
        self.frm_nuevo.columnconfigure(3, weight=1)

    def cargar_excel(self) -> None:
        if not os.path.exists(EXCEL_FILE):
            messagebox.showerror("Error", f"No existe el archivo: {EXCEL_FILE}")
            return
        try:
            df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        except PermissionError:
            messagebox.showerror("Error", "No se puede leer el Excel. ¿Está abierto en otra aplicación?")
            return
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el Excel:\n{exc}")
            return

        faltantes = [col for col in COLUMNAS_REQUERIDAS if col not in df.columns]
        if faltantes:
            messagebox.showerror("Error", f"Faltan columnas requeridas: {', '.join(faltantes)}")
            return

        self.df = df[COLUMNAS_REQUERIDAS].fillna("").copy()
        self.df["norm_marca"] = self.df["Marca"].map(normalizar_texto)
        self.df["norm_modelo"] = self.df["Modelo"].map(normalizar_texto)
        self.df["norm_diafragma"] = self.df["Diafragma"].map(normalizar_texto)

        marcas = sorted({m for m in self.df["Marca"].astype(str) if str(m).strip()})
        self.cmb_marca["values"] = marcas
        self.cmb_modelo.set("")
        self.cmb_modelo["values"] = []

    def al_elegir_marca(self, _event=None) -> None:
        marca = self.var_marca.get()
        modelos = sorted({m for m in self.df[self.df["Marca"] == marca]["Modelo"].astype(str) if m.strip()})
        self.cmb_modelo["values"] = modelos
        self.var_modelo.set("")
        self.mostrar_resultados(None)

    def al_elegir_modelo(self, _event=None) -> None:
        marca = self.var_marca.get()
        modelo = self.var_modelo.get()
        fila = self.df[(self.df["Marca"] == marca) & (self.df["Modelo"] == modelo)]
        self.mostrar_resultados(fila.iloc[0].to_dict() if not fila.empty else None)

    def buscar_similares(self, consulta_norm: str) -> tuple[list[dict], list[dict]]:
        """Retorna dos listas: coincidencias exactas y similares."""
        exactos, similares = [], []
        for _, fila in self.df.iterrows():
            score_modelo = fuzz.WRatio(consulta_norm, fila["norm_modelo"])
            score_diafragma = fuzz.WRatio(consulta_norm, fila["norm_diafragma"])
            score = max(score_modelo, score_diafragma)
            es_exacto = consulta_norm in {fila["norm_modelo"], fila["norm_diafragma"]}

            item = {
                "Marca": fila["Marca"],
                "Modelo": fila["Modelo"],
                "Carburador": fila["Carburador"],
                "Diafragma": fila["Diafragma"],
                "score": int(score),
            }
            if es_exacto:
                exactos.append(item)
            elif score >= 60:
                similares.append(item)

        exactos.sort(key=lambda x: x["score"], reverse=True)
        similares.sort(key=lambda x: x["score"], reverse=True)
        return exactos, similares

    def buscar_modelo(self) -> None:
        consulta = self.var_busqueda.get().strip()
        self.lista_resultados.delete(0, tk.END)
        self.resultados_actuales = []

        if not consulta:
            self.var_estado.set("Ingrese un modelo o código para buscar.")
            self.ocultar_formulario_nuevo()
            return

        consulta_norm = normalizar_texto(consulta)
        exactos, similares = self.buscar_similares(consulta_norm)

        if exactos:
            self.lista_resultados.insert(tk.END, "=== Coincidencias exactas ===")
            for r in exactos:
                self.resultados_actuales.append(r)
                self.lista_resultados.insert(tk.END, f"{r['Marca']} | {r['Modelo']} | Diafragma: {r['Diafragma']} | {r['score']}%")

        if similares:
            self.lista_resultados.insert(tk.END, "=== Coincidencias similares ===")
            for r in similares[:20]:
                self.resultados_actuales.append(r)
                self.lista_resultados.insert(tk.END, f"{r['Marca']} | {r['Modelo']} | Diafragma: {r['Diafragma']} | {r['score']}%")

        if not exactos and not similares:
            self.mostrar_sin_resultados()
            self.var_nuevo_modelo.set(consulta)
            self.mostrar_resultados(None)
            return

        self.var_estado.set("")
        self.ocultar_formulario_nuevo()

    def mostrar_sin_resultados(self) -> None:
        """Muestra aviso de no encontrado y habilita acciones de ayuda/carga manual."""
        self.var_estado.set("No se encontró el modelo en la planilla de consulta.")
        self.btn_chatgpt.pack(anchor="w", pady=(0, 8))
        self.mostrar_formulario_nuevo()

    def ocultar_formulario_nuevo(self) -> None:
        """Oculta botón de ChatGPT y formulario de alta manual."""
        self.btn_chatgpt.pack_forget()
        self.frm_nuevo.pack_forget()

    def mostrar_formulario_nuevo(self) -> None:
        """Muestra formulario de alta manual y precarga marca/modelo conocidos."""
        self.frm_nuevo.pack(fill="x")
        self.var_nueva_marca.set(self.var_marca.get().strip())
        if not self.var_nuevo_modelo.get().strip():
            self.var_nuevo_modelo.set(self.var_busqueda.get().strip())

    def mostrar_resultados(self, fila: dict | None) -> None:
        if not fila:
            self.lbl_carburador.config(text="-")
            self.lbl_diafragma.config(text="-")
            return
        self.lbl_carburador.config(text=str(fila.get("Carburador", "-")))
        self.lbl_diafragma.config(text=str(fila.get("Diafragma", "-")))

    def al_seleccionar_resultado(self, _event=None) -> None:
        i = self.lista_resultados.curselection()
        if not i:
            return
        texto = self.lista_resultados.get(i[0])
        if texto.startswith("==="):
            return

        idx_datos = 0
        for pos in range(i[0] + 1):
            if not self.lista_resultados.get(pos).startswith("==="):
                idx_datos += 1
        registro = self.resultados_actuales[idx_datos - 1]
        self.mostrar_resultados(registro)
        self.var_marca.set(str(registro["Marca"]))
        self.al_elegir_marca()
        self.var_modelo.set(str(registro["Modelo"]))

    def guardar_nuevo_registro(self) -> None:
        """Agrega una nueva fila al final del Excel existente y recarga la base."""
        marca = self.var_nueva_marca.get().strip()
        modelo = self.var_nuevo_modelo.get().strip()
        carburador = self.var_nuevo_carburador.get().strip()
        diafragma = self.var_nuevo_diafragma.get().strip()

        if not marca or not modelo or not carburador or not diafragma:
            messagebox.showwarning("Validación", "Complete todos los campos antes de guardar.")
            return

        marca_n = normalizar_texto(marca)
        modelo_n = normalizar_texto(modelo)
        existe = self.df[(self.df["norm_marca"] == marca_n) & (self.df["norm_modelo"] == modelo_n)]
        if not existe.empty:
            messagebox.showinfo("Información", "Ese modelo ya existe en la base.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            ws.append([marca, modelo, carburador, diafragma])
            wb.save(EXCEL_FILE)
        except PermissionError:
            messagebox.showerror("Error", "No se puede guardar. Cierre el Excel si está abierto e intente de nuevo.")
            return
        except Exception as exc:
            messagebox.showerror("Error", f"Ocurrió un error al guardar:\n{exc}")
            return

        self.cargar_excel()
        self.var_nueva_marca.set("")
        self.var_nuevo_modelo.set("")
        self.var_nuevo_carburador.set("")
        self.var_nuevo_diafragma.set("")
        self.ocultar_formulario_nuevo()
        self.var_estado.set("")
        messagebox.showinfo("Éxito", "Registro agregado correctamente.")

    def abrir_chatgpt(self) -> None:
        marca = self.var_nueva_marca.get().strip() or self.var_marca.get().strip() or "(sin marca)"
        modelo = self.var_nuevo_modelo.get().strip() or self.var_busqueda.get().strip() or "(sin modelo)"
        prompt = f"Buscar diafragma para:\nMarca: {marca}\nModelo: {modelo}"
        url = "https://chatgpt.com/?q=" + urllib.parse.quote(prompt)
        webbrowser.open(url)


if __name__ == "__main__":
    root = tk.Tk()
    app = DiafragmaApp(root)
    root.mainloop()
