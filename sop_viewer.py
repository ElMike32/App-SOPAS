import ctypes
import io
import os
import sys
import zipfile
import customtkinter as ctk
from openpyxl import load_workbook
import pandas as pd
from PIL import Image

# Configuración del tema visual y paleta moderna UI/UX
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

EXCEL_FILE = "Prueba1.xlsx"

# Paleta de colores profesionales
COLOR_BG_APP = "#F8FAFC"
COLOR_CARD_BG = "#FFFFFF"
COLOR_BORDER = "#E5E7EB"
COLOR_TEXT_PRIMARY = "#1E293B"
COLOR_TEXT_SECONDARY = "#64748B"
COLOR_ACCENT = "#1F4E79"
COLOR_BADGE_BG = "#F1F5F9"
FONT_FAMILY = "Roboto"


def ocultar_archivo_windows(ruta):
  if os.name == "nt" and os.path.exists(ruta):
    try:
      ctypes.windll.kernel32.SetFileAttributesW(str(ruta), 2)
    except Exception:
      pass


def limpiar_texto(val):
  if pd.isna(val):
    return ""
  s = str(val).strip()
  return "" if s.lower() in ["nan", "none"] else s


def limpiar_entero(val):
  s = limpiar_texto(val)
  if not s:
    return ""
  try:
    num = float(s)
    if num.is_integer():
      return str(int(num))
    return str(num)
  except ValueError:
    return s


def normalizar_texto(texto):
  texto = limpiar_texto(texto).lower()
  reemplazos = (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"))
  for a, b in reemplazos:
    texto = texto.replace(a, b)
  return texto


def normalizar_fondo_blanco(pil_img):
  try:
    if pil_img.mode in ("RGBA", "LA") or (
        pil_img.mode == "P" and "transparency" in pil_img.info
    ):
      alpha_img = pil_img.convert("RGBA")
      canvas = Image.new("RGBA", alpha_img.size, (255, 255, 255, 255))
      canvas.paste(alpha_img, mask=alpha_img.split()[3])
      return canvas.convert("RGB")
    return pil_img.convert("RGB")
  except Exception:
    return pil_img.convert("RGB")


# ==============================================================================
# MOTOR HÍBRIDO DE EXTRACCIÓN DE IMÁGENES A PRUEBA DE ERRORES
# ==============================================================================
def extraer_imagenes_hoja(ruta_excel, nombre_hoja, col_mat_idx=0):
  """Extrae imágenes combinando openpyxl y análisis ZIP directo para soportar cualquier tipo de anclaje de Excel"""
  mapa_imgs = {}
  if not os.path.exists(ruta_excel):
    return mapa_imgs

  try:
    wb = load_workbook(ruta_excel, data_only=True)
    if nombre_hoja not in wb.sheetnames:
      return mapa_imgs
    ws = wb[nombre_hoja]

    # Map de Fila -> Material en la hoja especificada
    filas_mat = {}
    for row in range(2, ws.max_row + 1):
      val_mat = limpiar_texto(ws.cell(row=row, column=col_mat_idx + 1).value)
      if val_mat:
        filas_mat[row] = val_mat

    if not filas_mat:
      return mapa_imgs

    # --- MÉTODO A: Extracción directa vía openpyxl ---
    for img in getattr(ws, "_images", []):
      try:
        row_target = None
        if hasattr(img, "anchor"):
          anc = img.anchor
          if hasattr(anc, "_from") and hasattr(anc._from, "row"):
            row_target = anc._from.row + 1
          elif hasattr(anc, "row"):
            row_target = anc.row + 1

        if row_target is None or row_target not in filas_mat:
          if row_target is not None:
            filas_validas = [r for r in filas_mat.keys() if r <= row_target]
            if filas_validas:
              row_target = max(filas_validas)

        if row_target in filas_mat:
          mat = filas_mat[row_target]
          pil_img = Image.open(io.BytesIO(img._data()))
          pil_img = normalizar_fondo_blanco(pil_img)

          if mat not in mapa_imgs:
            mapa_imgs[mat] = []
          mapa_imgs[mat].append(pil_img)
      except Exception:
        pass

    # --- MÉTODO B (Respaldos): Si openpyxl no detectó imágenes, extraer vía ZIP en orden estricto ---
    if not mapa_imgs:
      with zipfile.ZipFile(ruta_excel, "r") as z:
        media_files = [
            f for f in z.namelist() if f.startswith("xl/media/")
        ]
        if media_files:
          mats_unicos = list(dict.fromkeys(filas_mat.values()))
          for idx, media_path in enumerate(sorted(media_files)):
            if idx < len(mats_unicos):
              mat = mats_unicos[idx]
              img_data = z.read(media_path)
              pil_img = Image.open(io.BytesIO(img_data))
              pil_img = normalizar_fondo_blanco(pil_img)
              if mat not in mapa_imgs:
                mapa_imgs[mat] = []
              mapa_imgs[mat].append(pil_img)

  except Exception as e:
    print(f"[AVISO] Error al procesar imágenes de {nombre_hoja}: {e}")

  return mapa_imgs


# ==============================================================================
# APLICACIÓN PRINCIPAL
# ==============================================================================
class SOPApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    ruta_internal = os.path.join(os.path.dirname(sys.executable), "_internal")
    if os.path.exists(ruta_internal):
      ocultar_archivo_windows(ruta_internal)

    self.title("Visor de SOPs - Digitalización de Procesos")
    self.geometry("1400x890")
    self.minsize(1180, 780)
    self.configure(fg_color=COLOR_BG_APP)

    self.data_general = pd.DataFrame()
    self.data_estandar = pd.DataFrame()
    self.data_componentes = pd.DataFrame()
    self.data_alertas = pd.DataFrame()
    self.data_pasos = pd.DataFrame()
    self.data_puntos_seg = pd.DataFrame()
    self.data_epp_ob = pd.DataFrame()
    self.data_epp_req = pd.DataFrame()
    self.data_historial = pd.DataFrame()

    self.imgs_general = {}
    self.imgs_componentes = {}
    self.imgs_pasos = {}
    self.imgs_epp_ob = {}
    self.imgs_epp_req = {}

    self.materiales_unicos = []
    self.material_actual = None
    self.maquina_actual = None
    self.paso_actual_index = 0
    self.pasos_filtrados = []

    self._crear_interfaz()
    self.cargar_datos_excel()

  def _crear_interfaz(self):
    # 1. Cabecera Limpia (Top Bar)
    self.frame_header = ctk.CTkFrame(
        self,
        fg_color=COLOR_CARD_BG,
        corner_radius=12,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    self.frame_header.pack(fill="x", padx=20, pady=(12, 6))

    self.lbl_titulo_app = ctk.CTkLabel(
        self.frame_header,
        text="Digitalización de SOPs",
        font=(FONT_FAMILY, 18, "bold"),
        text_color=COLOR_ACCENT,
    )
    self.lbl_titulo_app.pack(side="left", padx=20, pady=10)

    self.frame_search_container = ctk.CTkFrame(
        self.frame_header, fg_color="transparent"
    )
    self.frame_search_container.pack(
        side="right", fill="x", expand=True, padx=20, pady=8
    )

    self.entry_busqueda = ctk.CTkEntry(
        self.frame_search_container,
        placeholder_text="🔍 Buscar por Material o Máquina...",
        font=(FONT_FAMILY, 13),
        height=36,
        corner_radius=8,
        border_color=COLOR_BORDER,
    )
    self.entry_busqueda.pack(fill="x", expand=True)
    self.entry_busqueda.bind("<KeyRelease>", self.al_escribir_buscador)

    self.frame_sugerencias = ctk.CTkScrollableFrame(
        self.frame_search_container,
        height=150,
        corner_radius=8,
        border_width=1,
        border_color=COLOR_ACCENT,
    )

    # 2. Banner Superior Horizontal Compacto
    self.frame_banner = ctk.CTkFrame(
        self,
        corner_radius=8,
        fg_color=COLOR_CARD_BG,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    self.frame_banner.pack(fill="x", padx=20, pady=(0, 8))

    self.kpi_line = ctk.CTkFrame(self.frame_banner, fg_color="transparent")
    self.kpi_line.pack(side="left", padx=12, pady=6)

    ctk.CTkLabel(
        self.kpi_line,
        text="MATERIAL:",
        font=(FONT_FAMILY, 11, "bold"),
        text_color=COLOR_TEXT_SECONDARY,
    ).pack(side="left", padx=(0, 4))

    self.lbl_banner_mat = ctk.CTkLabel(
        self.kpi_line,
        text="---",
        font=(FONT_FAMILY, 14, "bold"),
        text_color=COLOR_TEXT_PRIMARY,
    )
    self.lbl_banner_mat.pack(side="left", padx=(0, 16))

    ctk.CTkLabel(
        self.kpi_line,
        text="│",
        font=(FONT_FAMILY, 14),
        text_color=COLOR_BORDER,
    ).pack(side="left", padx=(0, 16))

    ctk.CTkLabel(
        self.kpi_line,
        text="MÁQUINA:",
        font=(FONT_FAMILY, 11, "bold"),
        text_color=COLOR_TEXT_SECONDARY,
    ).pack(side="left", padx=(0, 4))

    self.lbl_banner_maq = ctk.CTkLabel(
        self.kpi_line,
        text="---",
        font=(FONT_FAMILY, 14, "bold"),
        text_color=COLOR_ACCENT,
    )
    self.lbl_banner_maq.pack(side="left")

    self.lbl_banner_desc = ctk.CTkLabel(
        self.frame_banner,
        text="",
        font=(FONT_FAMILY, 13, "italic"),
        text_color=COLOR_TEXT_SECONDARY,
    )
    self.lbl_banner_desc.pack(side="right", padx=16, pady=6)

    # 3. Pestañas
    self.tabview = ctk.CTkTabview(
        self, corner_radius=12, fg_color=COLOR_CARD_BG
    )
    self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 12))

    self.tab_info = self.tabview.add("1. Info & Estándar")
    self.tab_comp = self.tabview.add("2. Componentes")
    self.tab_seg = self.tabview.add("3. Seguridad & EPP")
    self.tab_alertas = self.tabview.add("4. Alertas & Docs")
    self.tab_pasos = self.tabview.add("5. Pasos Operativos")
    self.tab_historial = self.tabview.add("6. Historial Cambios")

    self._setup_tab_pasos()

  def cargar_datos_excel(self):
    if not os.path.exists(EXCEL_FILE):
      return

    try:
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

      self.imgs_general = extraer_imagenes_hoja(EXCEL_FILE, xls.sheet_names[0])
      self.imgs_componentes = extraer_imagenes_hoja(
          EXCEL_FILE, xls.sheet_names[2]
      )
      self.imgs_pasos = extraer_imagenes_hoja(EXCEL_FILE, xls.sheet_names[4])
      self.imgs_epp_ob = extraer_imagenes_hoja(EXCEL_FILE, xls.sheet_names[6])
      self.imgs_epp_req = extraer_imagenes_hoja(EXCEL_FILE, xls.sheet_names[7])

      col_mat = self.data_general.columns[0]
      col_maq = (
          self.data_general.columns[2]
          if len(self.data_general.columns) > 2
          else col_mat
      )
      col_desc = (
          self.data_general.columns[6]
          if len(self.data_general.columns) > 6
          else col_mat
      )

      self.materiales_unicos.clear()
      for _, row in self.data_general.iterrows():
        m = limpiar_texto(row[col_mat])
        mq = limpiar_texto(row[col_maq])
        d = limpiar_texto(row[col_desc])
        if m:
          self.materiales_unicos.append((m, mq, d))

      if self.materiales_unicos:
        pm, pmq, _ = self.materiales_unicos[0]
        self.seleccionar_material(pm, pmq)

    except Exception as e:
      print(f"Error al procesar Excel: {e}")

  def al_escribir_buscador(self, event=None):
    query = normalizar_texto(self.entry_busqueda.get())
    if not query or not self.materiales_unicos:
      self.frame_sugerencias.pack_forget()
      return

    for child in self.frame_sugerencias.winfo_children():
      child.destroy()

    coincidencias = 0
    for mat, mq, desc in self.materiales_unicos:
      eval_txt = f"{mat} {mq} {desc}"
      if query in normalizar_texto(eval_txt):
        lbl_btn = f"{mat}  │  Máq: {mq}" + (f" ({desc})" if desc else "")
        btn = ctk.CTkButton(
            self.frame_sugerencias,
            text=lbl_btn,
            anchor="w",
            fg_color="transparent",
            text_color=COLOR_TEXT_PRIMARY,
            hover_color=COLOR_BADGE_BG,
            height=30,
            command=lambda m=mat, mq=mq: self.seleccionar_material(m, mq),
        )
        btn.pack(fill="x", pady=1)
        coincidencias += 1
        if coincidencias >= 15:
          break

    if coincidencias > 0:
      self.frame_sugerencias.pack(fill="x", pady=(4, 0))
    else:
      self.frame_sugerencias.pack_forget()

  def seleccionar_material(self, material, maquina):
    self.frame_sugerencias.pack_forget()
    self.entry_busqueda.delete(0, "end")
    self.focus_set()

    self.material_actual = material
    self.maquina_actual = maquina

    self.lbl_banner_mat.configure(text=material)
    self.lbl_banner_maq.configure(text=maquina)

    self.renderizar_todas_pestañas(material, maquina)

  def renderizar_todas_pestañas(self, material, maquina):
    self.renderizar_tab_info(material, maquina)
    self.renderizar_tab_componentes(material)
    self.renderizar_tab_seguridad(material)
    self.renderizar_tab_alertas(material)
    self.renderizar_tab_pasos(material)
    self.renderizar_tab_historial(material)

  # ==============================================================================
  # PESTAÑA 1: INFO & ESTÁNDAR
  # ==============================================================================
  def renderizar_tab_info(self, material, maquina):
    for child in self.tab_info.winfo_children():
      child.destroy()

    frame_main = ctk.CTkFrame(self.tab_info, fg_color="transparent")
    frame_main.pack(fill="both", expand=True, padx=5, pady=5)

    # 1. Card Datos Generales
    col_izq = ctk.CTkFrame(
        frame_main,
        fg_color=COLOR_CARD_BG,
        corner_radius=12,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    col_izq.pack(side="left", fill="both", expand=True, padx=(0, 10))

    # 2. Card Estándar & Herramientas
    col_centro = ctk.CTkFrame(
        frame_main,
        fg_color=COLOR_CARD_BG,
        corner_radius=12,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    col_centro.pack(side="left", fill="both", expand=True, padx=(0, 10))

    # 3. Card Visualización de Pieza
    col_der = ctk.CTkFrame(
        frame_main,
        fg_color=COLOR_CARD_BG,
        corner_radius=12,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    col_der.pack(side="right", fill="both", expand=True)

    # --- CONTENIDO COLUMNA 1: Datos Generales ---
    col_mat_g = self.data_general.columns[0]
    df_g = self.data_general[
        self.data_general[col_mat_g].apply(limpiar_texto) == material
    ]

    if not df_g.empty:
      row = df_g.iloc[0]
      desc = limpiar_texto(row.iloc[6]) if len(row) > 6 else ""
      self.lbl_banner_desc.configure(text=desc)

      ctk.CTkLabel(
          col_izq,
          text="Datos Generales",
          font=(FONT_FAMILY, 15, "bold"),
          text_color=COLOR_ACCENT,
      ).pack(anchor="w", padx=15, pady=(10, 6))

      frame_kv = ctk.CTkFrame(col_izq, fg_color="transparent")
      frame_kv.pack(fill="both", expand=True, padx=15, pady=(0, 10))

      cols = self.data_general.columns
      for col_name in cols:
        val = limpiar_texto(row[col_name])
        if col_name.lower().startswith("imagen"):
          continue

        nombre_limpio = str(col_name).strip().rstrip(":")

        f_item = ctk.CTkFrame(frame_kv, fg_color="transparent")
        f_item.pack(fill="x", pady=1.5)

        ctk.CTkLabel(
            f_item,
            text=f"{nombre_limpio}",
            font=(FONT_FAMILY, 13, "bold"),
            text_color=COLOR_TEXT_SECONDARY,
            anchor="w",
        ).pack(side="left")
        ctk.CTkLabel(
            f_item,
            text=val if val else "---",
            font=(FONT_FAMILY, 13, "bold"),
            text_color=COLOR_TEXT_PRIMARY,
            anchor="e",
        ).pack(side="right", fill="x", expand=True)

        divider = ctk.CTkFrame(
            frame_kv, height=1, fg_color=COLOR_BORDER, corner_radius=0
        )
        divider.pack(fill="x", pady=0.5)

    # --- CONTENIDO COLUMNA 2: Estándar y Herramientas ---
    col_mat_e = self.data_estandar.columns[0]
    df_e = self.data_estandar[
        self.data_estandar[col_mat_e].apply(limpiar_texto) == material
    ]

    if not df_e.empty:
      ctk.CTkLabel(
          col_centro,
          text="Estándar de Producción",
          font=(FONT_FAMILY, 15, "bold"),
          text_color=COLOR_ACCENT,
      ).pack(anchor="w", padx=15, pady=(10, 6))

      row1 = df_e.iloc[0]
      pzs_hr = limpiar_entero(row1.iloc[1])
      t1 = limpiar_entero(row1.iloc[2])
      t2 = limpiar_entero(row1.iloc[3])
      t3 = limpiar_entero(row1.iloc[4])
      pzs_ciclo = limpiar_entero(row1.iloc[5])
      tiempo_ciclo = limpiar_texto(row1.iloc[6])
      if tiempo_ciclo and not tiempo_ciclo.lower().endswith("seg."):
        tiempo_ciclo = f"{limpiar_entero(tiempo_ciclo)} seg."
      wip_max = limpiar_entero(row1.iloc[7])

      # Stat Card Principal
      stat_main = ctk.CTkFrame(
          col_centro, fg_color=COLOR_BADGE_BG, corner_radius=10
      )
      stat_main.pack(fill="x", padx=15, pady=(0, 8))

      ctk.CTkLabel(
          stat_main,
          text="Rendimiento Estándar",
          font=(FONT_FAMILY, 11),
          text_color=COLOR_TEXT_SECONDARY,
      ).pack(pady=(6, 0))
      ctk.CTkLabel(
          stat_main,
          text=f"{pzs_hr} Pzs/Hr",
          font=(FONT_FAMILY, 24, "bold"),
          text_color=COLOR_ACCENT,
      ).pack(pady=(0, 6))

      # Sub-tarjetas de Turnos
      grid_turnos = ctk.CTkFrame(col_centro, fg_color="transparent")
      grid_turnos.pack(fill="x", padx=15, pady=(0, 8))
      grid_turnos.grid_columnconfigure((0, 1, 2), weight=1)

      turnos_data = [("1er Turno", t1), ("2do Turno", t2), ("3er Turno", t3)]
      for i, (nom_t, val_t) in enumerate(turnos_data):
        card_t = ctk.CTkFrame(
            grid_turnos,
            fg_color="#FFFFFF",
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=8,
        )
        card_t.grid(row=0, column=i, padx=2, sticky="nsew")
        ctk.CTkLabel(
            card_t,
            text=nom_t,
            font=(FONT_FAMILY, 11, "bold"),
            text_color=COLOR_TEXT_SECONDARY,
        ).pack(pady=(4, 0))
        ctk.CTkLabel(
            card_t,
            text=val_t,
            font=(FONT_FAMILY, 16, "bold"),
            text_color=COLOR_TEXT_PRIMARY,
        ).pack(pady=(0, 4))

      # Métricas Secundarias
      f_met = ctk.CTkFrame(col_centro, fg_color="transparent")
      f_met.pack(fill="x", padx=15, pady=(0, 8))

      mets = [
          ("Piezas / Ciclo", pzs_ciclo),
          ("Tiempo Ciclo", tiempo_ciclo),
          ("WIP Máximo", wip_max),
      ]
      for mk, mv in mets:
        fm_row = ctk.CTkFrame(f_met, fg_color="transparent")
        fm_row.pack(fill="x", pady=1)
        ctk.CTkLabel(
            fm_row,
            text=mk,
            font=(FONT_FAMILY, 12),
            text_color=COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        ctk.CTkLabel(
            fm_row,
            text=mv,
            font=(FONT_FAMILY, 12, "bold"),
            text_color=COLOR_TEXT_PRIMARY,
        ).pack(side="right")

      # Sección Herramientas Requeridas
      col_herram = df_e.columns[8] if len(df_e.columns) > 8 else None
      if col_herram:
        herramientas = [
            limpiar_texto(h)
            for h in df_e[col_herram].dropna()
            if limpiar_texto(h)
        ]
        if herramientas:
          ctk.CTkLabel(
              col_centro,
              text="Herramientas Requeridas",
              font=(FONT_FAMILY, 14, "bold"),
              text_color=COLOR_ACCENT,
          ).pack(anchor="w", padx=15, pady=(4, 4))

          frame_pills = ctk.CTkFrame(col_centro, fg_color="transparent")
          frame_pills.pack(fill="both", expand=True, padx=15, pady=(0, 10))

          for h_item in herramientas[:9]:
            pill = ctk.CTkFrame(
                frame_pills, fg_color=COLOR_BADGE_BG, corner_radius=6
            )
            pill.pack(fill="x", pady=1.5)
            ctk.CTkLabel(
                pill,
                text=f"•  {h_item}",
                font=(FONT_FAMILY, 11, "bold"),
                text_color=COLOR_TEXT_PRIMARY,
                anchor="w",
            ).pack(anchor="w", padx=8, pady=2.5)

    # --- CONTENIDO COLUMNA 3: Visualización del Componente ---
    ctk.CTkLabel(
        col_der,
        text="Vista del Componente",
        font=(FONT_FAMILY, 15, "bold"),
        text_color=COLOR_ACCENT,
    ).pack(anchor="w", padx=15, pady=(10, 4))

    imgs_m = self.imgs_general.get(material, [])
    if imgs_m:
      pil_img = imgs_m[0]
      ctk_img = ctk.CTkImage(
          light_image=pil_img, dark_image=pil_img, size=(450, 450)
      )
      lbl_img = ctk.CTkLabel(col_der, image=ctk_img, text="")
      lbl_img.pack(expand=True, fill="both", padx=15, pady=15)
    else:
      ctk.CTkLabel(
          col_der,
          text="📷 Fotografía del Material\n(No disponible)",
          font=(FONT_FAMILY, 14),
          text_color=COLOR_TEXT_SECONDARY,
      ).pack(expand=True)

  # ==============================================================================
  # OTRAS PESTAÑAS
  # ==============================================================================
  def renderizar_tab_componentes(self, material):
    for child in self.tab_comp.winfo_children():
      child.destroy()

    scroll = ctk.CTkScrollableFrame(self.tab_comp, fg_color="transparent")
    scroll.pack(fill="both", expand=True)

    col_mat_c = self.data_componentes.columns[0]
    df_c = self.data_componentes[
        self.data_componentes[col_mat_c].apply(limpiar_texto) == material
    ]

    if df_c.empty:
      ctk.CTkLabel(
          scroll,
          text="No hay componentes registrados.",
          font=(FONT_FAMILY, 14),
      ).pack(pady=40)
      return

    scroll.grid_columnconfigure((0, 1), weight=1)
    imgs_mat = self.imgs_componentes.get(material, [])

    for idx, (_, row) in enumerate(df_c.iterrows()):
      nombre = limpiar_texto(row.iloc[1]) if len(row) > 1 else ""
      no_parte = limpiar_texto(row.iloc[2]) if len(row) > 2 else ""
      codigo = limpiar_texto(row.iloc[3]) if len(row) > 3 else ""

      r, c = divmod(idx, 2)
      card = ctk.CTkFrame(
          scroll,
          corner_radius=10,
          border_width=1,
          border_color=COLOR_BORDER,
          fg_color=COLOR_CARD_BG,
      )
      card.grid(row=r, column=c, padx=10, pady=10, sticky="nsew")

      ctk.CTkLabel(
          card,
          text=codigo if codigo else nombre,
          font=(FONT_FAMILY, 15, "bold"),
          text_color=COLOR_ACCENT,
      ).pack(anchor="w", padx=15, pady=(12, 2))

      if no_parte:
        ctk.CTkLabel(
            card,
            text=f"No. Parte: {no_parte}",
            font=(FONT_FAMILY, 12),
            text_color=COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=15, pady=2)

      if idx < len(imgs_mat):
        pil_img = imgs_mat[idx]
        ctk_img = ctk.CTkImage(
            light_image=pil_img, dark_image=pil_img, size=(140, 140)
        )
        ctk.CTkLabel(card, image=ctk_img, text="").pack(pady=10)

  def renderizar_tab_seguridad(self, material):
    for child in self.tab_seg.winfo_children():
      child.destroy()

    scroll = ctk.CTkScrollableFrame(self.tab_seg, fg_color="transparent")
    scroll.pack(fill="both", expand=True)

    frame_ob = ctk.CTkFrame(
        scroll,
        corner_radius=10,
        fg_color=COLOR_CARD_BG,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    frame_ob.pack(fill="x", padx=10, pady=10)

    ctk.CTkLabel(
        frame_ob,
        text="Equipo de Seguridad Obligatorio",
        font=(FONT_FAMILY, 15, "bold"),
        text_color="#D32F2F",
    ).pack(anchor="w", padx=15, pady=12)

    imgs_ob = self.imgs_epp_ob.get(material, [])
    if imgs_ob:
      grid_ob = ctk.CTkFrame(frame_ob, fg_color="transparent")
      grid_ob.pack(fill="x", padx=15, pady=(0, 10))

      for pil_img in imgs_ob:
        ctk_img = ctk.CTkImage(
            light_image=pil_img, dark_image=pil_img, size=(100, 100)
        )
        card_i = ctk.CTkFrame(
            grid_ob,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=8,
        )
        card_i.pack(side="left", padx=6, pady=6)
        ctk.CTkLabel(card_i, image=ctk_img, text="").pack(padx=8, pady=8)

  def renderizar_tab_alertas(self, material):
    for child in self.tab_alertas.winfo_children():
      child.destroy()

    scroll = ctk.CTkScrollableFrame(self.tab_alertas, fg_color="transparent")
    scroll.pack(fill="both", expand=True)

    col_mat_a = self.data_alertas.columns[0]
    df_a = self.data_alertas[
        self.data_alertas[col_mat_a].apply(limpiar_texto) == material
    ]

    if df_a.empty:
      ctk.CTkLabel(
          scroll,
          text="No hay alertas registradas.",
          font=(FONT_FAMILY, 14),
      ).pack(pady=40)
      return

    frame_a = ctk.CTkFrame(
        scroll,
        corner_radius=10,
        fg_color=COLOR_CARD_BG,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    frame_a.pack(fill="x", padx=10, pady=10)

    ctk.CTkLabel(
        frame_a,
        text="Alertas de Calidad y Documentación",
        font=(FONT_FAMILY, 15, "bold"),
        text_color=COLOR_ACCENT,
    ).pack(anchor="w", padx=15, pady=12)

    col_lista = df_a.columns[1] if len(df_a.columns) > 1 else col_mat_a
    for _, row in df_a.iterrows():
      item_txt = limpiar_texto(row[col_lista])
      if item_txt:
        card_item = ctk.CTkFrame(
            frame_a,
            fg_color=COLOR_BADGE_BG,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=8,
        )
        card_item.pack(fill="x", padx=15, pady=4)
        ctk.CTkLabel(
            card_item,
            text=f"•  {item_txt}",
            font=(FONT_FAMILY, 13),
            text_color=COLOR_TEXT_PRIMARY,
            anchor="w",
        ).pack(padx=14, pady=10, fill="x")

  def _setup_tab_pasos(self):
    self.frame_paso_container = ctk.CTkFrame(
        self.tab_pasos, fg_color="transparent"
    )
    self.frame_paso_container.pack(fill="both", expand=True, padx=5, pady=5)

    self.lbl_paso_num = ctk.CTkLabel(
        self.frame_paso_container,
        text="",
        font=(FONT_FAMILY, 16, "bold"),
        text_color=COLOR_ACCENT,
    )
    self.lbl_paso_num.pack(pady=(5, 5))

    self.frame_contenido_paso = ctk.CTkFrame(
        self.frame_paso_container, fg_color="transparent"
    )
    self.frame_contenido_paso.pack(fill="both", expand=True, padx=10, pady=5)

    self.frame_col_izq = ctk.CTkFrame(
        self.frame_contenido_paso, width=450, fg_color="transparent"
    )
    self.frame_col_izq.pack(side="left", fill="both", padx=(0, 10))

    self.box_paso_desc = ctk.CTkTextbox(
        self.frame_col_izq, font=(FONT_FAMILY, 14), corner_radius=8
    )
    self.box_paso_desc.pack(fill="both", expand=True, pady=(0, 10))

    self.lbl_paso_nota = ctk.CTkLabel(
        self.frame_col_izq,
        text="",
        font=(FONT_FAMILY, 13, "bold"),
        text_color="#B71C1C",
        anchor="w",
        justify="left",
        wraplength=420,
    )
    self.lbl_paso_nota.pack(fill="x", pady=2)

    self.frame_col_der = ctk.CTkFrame(
        self.frame_contenido_paso,
        fg_color=COLOR_CARD_BG,
        corner_radius=10,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    self.frame_col_der.pack(side="right", fill="both", expand=True)

    self.lbl_paso_imagen = ctk.CTkLabel(
        self.frame_col_der, text="Sin imagen de paso"
    )
    self.lbl_paso_imagen.pack(expand=True, fill="both", padx=10, pady=10)

    self.frame_nav_botones = ctk.CTkFrame(
        self.frame_paso_container, fg_color="transparent"
    )
    self.frame_nav_botones.pack(pady=8)

    self.btn_anterior = ctk.CTkButton(
        self.frame_nav_botones,
        text="◄ Paso Anterior",
        font=(FONT_FAMILY, 12, "bold"),
        command=self.paso_anterior,
        width=140,
    )
    self.btn_anterior.pack(side="left", padx=8)

    self.btn_siguiente = ctk.CTkButton(
        self.frame_nav_botones,
        text="Paso Siguiente ►",
        font=(FONT_FAMILY, 12, "bold"),
        command=self.paso_siguiente,
        width=140,
    )
    self.btn_siguiente.pack(side="left", padx=8)

  def renderizar_tab_pasos(self, material):
    col_mat_p = self.data_pasos.columns[0]
    self.pasos_filtrados = (
        self.data_pasos[
            self.data_pasos[col_mat_p].apply(limpiar_texto) == material
        ].to_dict("records")
        if not self.data_pasos.empty
        else []
    )

    self.paso_actual_index = 0
    self.actualizar_vista_paso()

  def actualizar_vista_paso(self):
    total = len(self.pasos_filtrados)
    if total == 0:
      self.lbl_paso_num.configure(text="SOP sin pasos registrados")
      self.box_paso_desc.delete("1.0", "end")
      self.lbl_paso_nota.configure(text="")
      self.lbl_paso_imagen.configure(image=None, text="Sin imagen de paso")
      self.btn_anterior.configure(state="disabled")
      self.btn_siguiente.configure(state="disabled")
      return

    paso_data = self.pasos_filtrados[self.paso_actual_index]
    nombre_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[1]))
    desc_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[2]))
    nota_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[3]))
    simb_paso = limpiar_texto(paso_data.get(self.data_pasos.columns[4]))

    self.lbl_paso_num.configure(text=nombre_paso)

    self.box_paso_desc.delete("1.0", "end")
    self.box_paso_desc.insert("1.0", desc_paso)

    if nota_paso:
      txt_nota = f"[{simb_paso}] {nota_paso}" if simb_paso else nota_paso
      self.lbl_paso_nota.configure(text=txt_nota)
    else:
      self.lbl_paso_nota.configure(text="")

    imgs_p = self.imgs_pasos.get(self.material_actual, [])
    if self.paso_actual_index < len(imgs_p):
      pil_img = imgs_p[self.paso_actual_index]
      ctk_img = ctk.CTkImage(
          light_image=pil_img, dark_image=pil_img, size=(450, 320)
      )
      self.lbl_paso_imagen.configure(image=ctk_img, text="")
    else:
      self.lbl_paso_imagen.configure(image=None, text="Sin imagen disponible")

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

  def renderizar_tab_historial(self, material):
    for child in self.tab_historial.winfo_children():
      child.destroy()

    scroll = ctk.CTkScrollableFrame(self.tab_historial, fg_color="transparent")
    scroll.pack(fill="both", expand=True)

    col_mat_h = self.data_historial.columns[0]
    df_h = self.data_historial[
        self.data_historial[col_mat_h].apply(limpiar_texto) == material
    ]

    if df_h.empty:
      ctk.CTkLabel(
          scroll,
          text="No hay historial registrado.",
          font=(FONT_FAMILY, 14),
      ).pack(pady=40)
      return

    frame_h = ctk.CTkFrame(
        scroll,
        corner_radius=10,
        fg_color=COLOR_CARD_BG,
        border_width=1,
        border_color=COLOR_BORDER,
    )
    frame_h.pack(fill="x", padx=10, pady=10)

    ctk.CTkLabel(
        frame_h,
        text="Historial de Cambios y Revisiones",
        font=(FONT_FAMILY, 15, "bold"),
        text_color=COLOR_ACCENT,
    ).pack(anchor="w", padx=15, pady=12)

    for _, row in df_h.iterrows():
      rev = limpiar_texto(row.iloc[1]) if len(row) > 1 else ""
      fecha = limpiar_texto(row.iloc[2]) if len(row) > 2 else ""
      depto = limpiar_texto(row.iloc[3]) if len(row) > 3 else ""
      desc = limpiar_texto(row.iloc[4]) if len(row) > 4 else ""

      card = ctk.CTkFrame(
          frame_h,
          fg_color=COLOR_BADGE_BG,
          border_width=1,
          border_color=COLOR_BORDER,
          corner_radius=8,
      )
      card.pack(fill="x", padx=15, pady=4)

      lbl_t = f"Revisión: {rev}  │  Fecha: {fecha}  │  Depto: {depto}"
      ctk.CTkLabel(
          card,
          text=lbl_t,
          font=(FONT_FAMILY, 13, "bold"),
          text_color=COLOR_ACCENT,
          anchor="w",
      ).pack(padx=14, pady=(8, 2), fill="x")
      ctk.CTkLabel(
          card,
          text=desc,
          font=(FONT_FAMILY, 13),
          text_color=COLOR_TEXT_PRIMARY,
          anchor="w",
          justify="left",
      ).pack(padx=14, pady=(0, 8), fill="x")


if __name__ == "__main__":
  app = SOPApp()
  app.mainloop()
