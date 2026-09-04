import io
import os
import re
import sys
import ctypes
import pandas as pd
from openpyxl import load_workbook
from PIL import Image
import customtkinter as ctk

# Configuración de tema visual
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

EXCEL_FILE = "Prueba1.xlsx"

def ocultar_archivo_windows(ruta):
    """ Oculta un archivo o carpeta en sistemas Windows mediante la API Win32 """
    if os.name == 'nt' and os.path.exists(ruta):
        try:
            ctypes.windll.kernel32.SetFileAttributesW(str(ruta), 2)
        except Exception as e:
            print(f"[AVISO] No se pudo ocultar {ruta}: {e}")

def obtener_ruta_recurso(ruta_relativa):
    """ Obtiene la ruta absoluta para recursos empaquetados por PyInstaller """
    try:
        ruta_base = sys._MEIPASS
    except Exception:
        ruta_base = os.path.abspath(".")
    return os.path.join(ruta_base, ruta_relativa)

class SOPApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Ocultar la carpeta _internal al compilar en modo --onedir
        ruta_internal = os.path.join(os.path.dirname(sys.executable), "_internal")
        if os.path.exists(ruta_internal):
            ocultar_archivo_windows(ruta_internal)

        self.title("Visor de SOPs - Control de Procesos")
        self.geometry("1280x850")
        self.minsize(1100, 750)
      
    # Contenedores de datos
    self.data_general = pd.DataFrame()
    self.data_estandar = pd.DataFrame()
    self.data_componentes = pd.DataFrame()
    self.data_alertas = pd.DataFrame()
    self.data_pasos = pd.DataFrame()
    self.data_puntos_seg = pd.DataFrame()
    self.data_epp_ob = pd.DataFrame()
    self.data_epp_req = pd.DataFrame()
    self.data_historial = pd.DataFrame()

    self.materiales_unicos = []
    self.material_actual = None
    self.paso_actual_index = 0
    self.pasos_filtrados = []

    # Construir UI base
    self._crear_interfaz()

    # Cargar Excel local
    self.cargar_datos_excel()

  def _crear_interfaz(self):
    # 1. Header Superior (Buscador y Selección)
    self.frame_header = ctk.CTkFrame(self, corner_radius=10)
    self.frame_header.pack(fill="x", padx=15, pady=10)

    self.lbl_titulo_app = ctk.CTkLabel(
        self.frame_header,
        text="📘 DIGITALIZACIÓN DE SOPs",
        font=("Helvetica", 16, "bold"),
        text_color="#1F4E79",
    )
    self.lbl_titulo_app.pack(side="left", padx=15, pady=10)

    # Buscador Desplegable
    self.frame_search_container = ctk.CTkFrame(
        self.frame_header, fg_color="transparent"
    )
    self.frame_search_container.pack(
        side="right", fill="x", expand=True, padx=15, pady=10
    )

    self.entry_busqueda = ctk.CTkEntry(
        self.frame_search_container,
        placeholder_text="🔍 Escriba para buscar Material o Descripción...",
        font=("Helvetica", 13),
        height=38,
    )
    self.entry_busqueda.pack(fill="x", expand=True)
    self.entry_busqueda.bind("<KeyRelease>", self.al_escribir_buscador)

    self.frame_sugerencias = ctk.CTkScrollableFrame(
        self.frame_search_container,
        height=140,
        corner_radius=6,
        border_width=1,
        border_color="#1F4E79",
    )

    # 2. Resumen Banner del Material Seleccionado
    self.frame_banner = ctk.CTkFrame(
        self, corner_radius=8, fg_color="#E6F0FA", height=50
    )
    self.frame_banner.pack(fill="x", padx=15, pady=(0, 10))

    self.lbl_banner_mat = ctk.CTkLabel(
        self.frame_banner,
        text="MATERIAL: ---",
        font=("Helvetica", 16, "bold"),
        text_color="#111111",
    )
    self.lbl_banner_mat.pack(side="left", padx=15, pady=8)

    self.lbl_banner_desc = ctk.CTkLabel(
        self.frame_banner,
        text="Seleccione un material...",
        font=("Helvetica", 14),
        text_color="#333333",
    )
    self.lbl_banner_desc.pack(side="left", padx=15, pady=8)

    # 3. Pestañas Principales (CTkTabview)
    self.tabview = ctk.CTkTabview(self, corner_radius=10)
    self.tabview.pack(fill="both", expand=True, padx=15, pady=(0, 10))

    self.tab_info = self.tabview.add("1. Info & Estándar")
    self.tab_comp = self.tabview.add("2. Componentes")
    self.tab_seg = self.tabview.add("3. Seguridad & EPP")
    self.tab_alertas = self.tabview.add("4. Alertas & Docs")
    self.tab_pasos = self.tabview.add("5. Pasos Operativos")
    self.tab_historial = self.tabview.add("6. Historial Cambios")

    # Inicializar las vistas dentro de cada pestaña
    self._setup_tab_info()
    self._setup_tab_pasos()

  # ==============================================================================
  # CARGA Y PROCESAMIENTO DEL EXCEL
  # ==============================================================================
  def cargar_datos_excel(self):
    if not os.path.exists(EXCEL_FILE):
      ctk.CTkLabel(
          self.tab_info,
          text=f"⚠️ No se encontró el archivo '{EXCEL_FILE}' en la carpeta actual.",
          font=("Helvetica", 16, "bold"),
          text_color="red",
      ).pack(pady=50)
      return

    try:
      # Leer hojas usando pandas (ultra rápido)
      xls = pd.ExcelFile(EXCEL_FILE)
      self.data_general = pd.read_excel(xls, sheet_name=0)
      self.data_estandar = pd.read_excel(xls, sheet_name=1)
      self.data_componentes = pd.read_excel(xls, sheet_name=2)
      self.data_alertas = pd.read_excel(xls, sheet_name=3)
      self.data_pasos = pd.read_excel(xls, sheet_name=4)
      self.data_puntos_seg = pd.read_excel(xls, sheet_name=5)
      self.data_epp_ob = pd.read_excel(xls, sheet_name=6)
      self.data_epp_req = pd.read_excel(xls, sheet_name=7)
      self.data_historial = pd.read_excel(xls, sheet_name=8)

      # Extraer lista de materiales únicos de la Hoja 1
      col_mat = self.data_general.columns[0]
      col_desc = (
          self.data_general.columns[6]
          if len(self.data_general.columns) > 6
          else col_mat
      )

      self.materiales_unicos.clear()
      for _, row in self.data_general.iterrows():
        m = limpiar_texto(row[col_mat])
        d = limpiar_texto(row[col_desc])
        if m:
          self.materiales_unicos.append((m, d))

      if self.materiales_unicos:
        primer_mat, _ = self.materiales_unicos[0]
        self.seleccionar_material(primer_mat)

    except Exception as e:
      print(f"Error al leer Excel: {e}")

  # ==============================================================================
  # BUSCADOR Y NAVEGACIÓN
  # ==============================================================================
  def al_escribir_buscador(self, event=None):
    query = normalizar_texto(self.entry_busqueda.get())
    if not query or not self.materiales_unicos:
      self.frame_sugerencias.pack_forget()
      return

    for child in self.frame_sugerencias.winfo_children():
      child.destroy()

    coincidencias = 0
    for mat, desc in self.materiales_unicos:
      eval_txt = f"{mat} {desc}"
      if query in normalizar_texto(eval_txt):
        btn = ctk.CTkButton(
            self.frame_sugerencias,
            text=f"{mat}  │  {desc}",
            anchor="w",
            fg_color="transparent",
            text_color="#111111",
            hover_color="#D0E0F0",
            height=28,
            command=lambda m=mat: self.seleccionar_material(m),
        )
        btn.pack(fill="x", pady=1)
        coincidencias += 1
        if coincidencias >= 15:
          break

    if coincidencias > 0:
      self.frame_sugerencias.pack(fill="x", pady=(4, 0))
    else:
      self.frame_sugerencias.pack_forget()

  def seleccionar_material(self, material):
    self.frame_sugerencias.pack_forget()
    self.entry_busqueda.delete(0, "end")
    self.focus_set()

    self.material_actual = material
    self.lbl_banner_mat.configure(text=f"MATERIAL: {material}")

    # Cargar y Renderizar datos correspondientes
    self.renderizar_sop(material)

  def renderizar_sop(self, material):
    # 1. Pestaña 1: Info & Estándar
    self.desplegar_info_y_estandar(material)

    # 2. Pestaña 5: Pasos Operativos
    self.desplegar_pasos_operativos(material)

  # ==============================================================================
  # PESTAÑA 1: INFO & ESTÁNDAR
  # ==============================================================================
  def _setup_tab_info(self):
    self.scroll_info = ctk.CTkScrollableFrame(self.tab_info)
    self.scroll_info.pack(fill="both", expand=True)

  def desplegar_info_y_estandar(self, material):
    for child in self.scroll_info.winfo_children():
      child.destroy()

    # Filtrar datos Hoja 1
    col_mat_g = self.data_general.columns[0]
    df_g = self.data_general[
        self.data_general[col_mat_g].apply(limpiar_texto) == material
    ]

    if not df_g.empty:
      row = df_g.iloc[0]
      desc = (
          limpiar_texto(row.iloc[6]) if len(row) > 6 else "Sin descripción"
      )
      self.lbl_banner_desc.configure(text=desc)

      # Marco Datos Generales
      frame_f = ctk.CTkFrame(self.scroll_info, corner_radius=8)
      frame_f.pack(fill="x", padx=10, pady=10)

      ctk.CTkLabel(
          frame_f,
          text="📋 DATOS GENERALES Y APROBACIONES",
          font=("Helvetica", 14, "bold"),
          text_color="#1F4E79",
      ).pack(anchor="w", padx=10, pady=5)

      # Muestra dinámica de campos
      cols = self.data_general.columns
      grid_f = ctk.CTkFrame(frame_f, fg_color="transparent")
      grid_f.pack(fill="x", padx=10, pady=5)

      for idx, col_name in enumerate(cols):
        val = limpiar_texto(row[col_name])
        r, c = divmod(idx, 3)
        lbl = ctk.CTkLabel(
            grid_f,
            text=f"• {col_name}: {val}",
            font=("Helvetica", 12),
            anchor="w",
        )
        lbl.grid(row=r, column=c, sticky="w", padx=10, pady=3)

    # Filtrar datos Hoja 2 (Estándar & Herramientas)
    col_mat_e = self.data_estandar.columns[0]
    df_e = self.data_estandar[
        self.data_estandar[col_mat_e].apply(limpiar_texto) == material
    ]

    if not df_e.empty:
      frame_e = ctk.CTkFrame(self.scroll_info, corner_radius=8)
      frame_e.pack(fill="x", padx=10, pady=10)

      ctk.CTkLabel(
          frame_e,
          text="⚙️ ESTÁNDAR DE PRODUCCIÓN Y HERRAMIENTAS",
          font=("Helvetica", 14, "bold"),
          text_color="#1F4E79",
      ).pack(anchor="w", padx=10, pady=5)

      # Tiempos (primera fila)
      row1 = df_e.iloc[0]
      info_tiempos = (
          f"Pzs/Hr: {limpiar_texto(row1.iloc[1])}  │  1er Turno:"
          f" {limpiar_texto(row1.iloc[2])}  │  2do Turno:"
          f" {limpiar_texto(row1.iloc[3])}  │  3er Turno:"
          f" {limpiar_texto(row1.iloc[4])}\n"
          f"Pzs/Ciclo: {limpiar_texto(row1.iloc[5])}  │  Tiempo Ciclo:"
          f" {limpiar_texto(row1.iloc[6])}  │  WIP MAX:"
          f" {limpiar_texto(row1.iloc[7])}"
      )

      ctk.CTkLabel(
          frame_e,
          text=info_tiempos,
          font=("Helvetica", 12, "bold"),
          justify="left",
      ).pack(anchor="w", padx=15, pady=5)

      # Herramientas asociadas (Lista dinámica)
      col_herram = df_e.columns[8] if len(df_e.columns) > 8 else None
      if col_herram:
        herramientas = [
            limpiar_texto(h)
            for h in df_e[col_herram].dropna()
            if limpiar_texto(h)
        ]
        if herramientas:
          ctk.CTkLabel(
              frame_e,
              text=f"🔧 Herramientas Requeridas: {', '.join(herramientas)}",
              font=("Helvetica", 12, "italic"),
              text_color="#333333",
          ).pack(anchor="w", padx=15, pady=(0, 10))

  # ==============================================================================
  # PESTAÑA 5: PASOS OPERATIVOS Y PUNTOS DE SEGURIDAD
  # ==============================================================================
  def _setup_tab_pasos(self):
    # Área Superior: Visor de Paso Activo
    self.frame_paso_container = ctk.CTkFrame(self.tab_pasos, corner_radius=10)
    self.frame_paso_container.pack(fill="both", expand=True, padx=5, pady=5)

    self.lbl_paso_num = ctk.CTkLabel(
        self.frame_paso_container,
        text="PASO 0 DE 0",
        font=("Helvetica", 18, "bold"),
        text_color="#1F4E79",
    )
    self.lbl_paso_num.pack(pady=(10, 5))

    self.box_paso_desc = ctk.CTkTextbox(
        self.frame_paso_container,
        font=("Helvetica", 16),
        height=140,
        corner_radius=6,
    )
    self.box_paso_desc.pack(fill="x", padx=20, pady=5)

    self.lbl_paso_nota = ctk.CTkLabel(
        self.frame_paso_container,
        text="",
        font=("Helvetica", 13, "italic"),
        text_color="darkred",
    )
    self.lbl_paso_nota.pack(pady=2)

    # Navegación entre pasos (Botones Anterior / Siguiente)
    self.frame_nav_botones = ctk.CTkFrame(
        self.frame_paso_container, fg_color="transparent"
    )
    self.frame_nav_botones.pack(pady=10)

    self.btn_anterior = ctk.CTkButton(
        self.frame_nav_botones,
        text="◄ Paso Anterior",
        font=("Helvetica", 13, "bold"),
        command=self.paso_anterior,
        width=150,
    )
    self.btn_anterior.pack(side="left", padx=10)

    self.btn_siguiente = ctk.CTkButton(
        self.frame_nav_botones,
        text="Paso Siguiente ►",
        font=("Helvetica", 13, "bold"),
        command=self.paso_siguiente,
        width=150,
    )
    self.btn_siguiente.pack(side="left", padx=10)

    # Area Inferior FIJA: Leyenda Puntos de Seguridad (6a Hoja)
    self.frame_puntos_seguridad = ctk.CTkFrame(
        self.tab_pasos, height=110, corner_radius=8, fg_color="#FFF3CD"
    )
    self.frame_puntos_seguridad.pack(fill="x", padx=5, pady=(5, 0))

    ctk.CTkLabel(
        self.frame_puntos_seguridad,
        text="🛡️ LEYENDA - PUNTOS DE SEGURIDAD E INDICADORES",
        font=("Helvetica", 11, "bold"),
        text_color="#856404",
    ).pack(anchor="w", padx=10, pady=(4, 0))

    self.scroll_puntos_seg = ctk.CTkScrollableFrame(
        self.frame_puntos_seguridad, orientation="horizontal", height=65
    )
    self.scroll_puntos_seg.pack(fill="x", padx=5, pady=4)

  def desplegar_pasos_operativos(self, material):
    col_mat_p = self.data_pasos.columns[0]
    self.pasos_filtrados = (
        self.data_pasos[
            self.data_pasos[col_mat_p].apply(limpiar_texto) == material
        ]
        .to_dict("records")
        if not self.data_pasos.empty
        else []
    )

    self.paso_actual_index = 0
    self.actualizar_vista_paso()

    # Cargar leyenda de seguridad fija (Hoja 6)
    for child in self.scroll_puntos_seg.winfo_children():
      child.destroy()

    if not self.data_puntos_seg.empty:
      for _, r in self.data_puntos_seg.iterrows():
        tit = limpiar_texto(r.iloc[0])
        desc = limpiar_texto(r.iloc[1])
        simb = limpiar_texto(r.iloc[2])

        card_seg = ctk.CTkFrame(
            self.scroll_puntos_seg, fg_color="#FFFFFF", corner_radius=4
        )
        card_seg.pack(side="left", padx=5, pady=2)

        txt_p_seg = f"[{simb}] {tit}: {desc}" if simb else f"{tit}: {desc}"
        ctk.CTkLabel(
            card_seg,
            text=txt_p_seg,
            font=("Helvetica", 10, "bold"),
            text_color="#333333",
        ).pack(padx=8, pady=4)

  def actualizar_vista_paso(self):
    total = len(self.pasos_filtrados)
    if total == 0:
      self.lbl_paso_num.configure(text="PASO 0 DE 0")
      self.box_paso_desc.delete("1.0", "end")
      self.box_paso_desc.insert(
          "1.0", "No hay pasos registrados para este material."
      )
      self.lbl_paso_nota.configure(text="")
      self.btn_anterior.configure(state="disabled")
      self.btn_siguiente.configure(state="disabled")
      return

    paso_data = self.pasos_filtrados[self.paso_actual_index]
    num_paso = (
        limpiar_texto(paso_data.get(self.data_pasos.columns[1]))
        or str(self.paso_actual_index + 1)
    )
    desc_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[2]))
    nota_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[3]))
    simb_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[4]))

    txt_num = f"PASO {self.paso_actual_index + 1} DE {total} (Paso {num_paso})"
    if simb_paso:
      txt_num += f"  - [ Símbolo: {simb_paso} ]"

    self.lbl_paso_num.configure(text=txt_num)

    self.box_paso_desc.delete("1.0", "end")
    self.box_paso_desc.insert("1.0", desc_paso)

    self.lbl_paso_nota.configure(
        text=f"⚠️ Nota: {nota_paso}" if nota_paso else ""
    )

    # Estado de botones
    self.btn_anterior.configure(
        state="normal" if self.paso_actual_index > 0 else "disabled"
    )
    self.btn_siguiente.configure(
        state="normal" if self.paso_actual_index < total - 1 else "disabled"
    )

  def paso_anterior(self):
    if self.paso_actual_index > 0:
      self.paso_actual_index -= 1
      self.actualizar_vista_paso()

  def paso_siguiente(self):
    if self.paso_actual_index < len(self.pasos_filtrados) - 1:
      self.paso_actual_index += 1
      self.actualizar_vista_paso()


if __name__ == "__main__":
  app = SOPApp()
  app.mainloop()
