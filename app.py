import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import io

# --- 1. CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)
COLOR_GRIS_LINEA = (220, 220, 220)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)
COLOR_CELESTE_PASTEL = (186, 225, 255)       # Celeste pastel para Cambio Categoría
COLOR_AZUL_PASTEL_OSCURO = (120, 180, 235)  # Celeste/Azul más oscuro para Cambio Línea

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = "Resumen de Dotación"

    def header(self):
        self.set_font("Arial", "B", 18)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, self.report_title, 0, 0, "C")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def draw_section_title(self, title):
        self.set_font("Arial", "B", 13)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 8, title, ln=True, align="L")
        self.set_draw_color(*COLOR_AZUL_INSTITUCIONAL)
        self.set_line_width(0.5)
        self.line(self.get_x(), self.get_y(), self.get_x() + self.page_width, self.get_y())
        self.ln(4)

    def draw_kpi_box(self, title, value, color, x, y, width=80):
        kpi_height = 16
        self.set_xy(x, y)
        self.set_fill_color(*color)
        self.cell(width, 1.5, "", fill=True, ln=False, border=0)
        self.set_xy(x, y + 1.5)
        self.set_fill_color(255, 255, 255)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.cell(width, kpi_height - 1.5, "", border=1, fill=True)
        self.set_xy(x, y + 3)
        self.set_font('Arial', '', 9)
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.cell(width, 8, title, align='C')
        self.set_xy(x, y + 8)
        self.set_font('Arial', 'B', 15)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(width, 10, str(value), align='C')

    def draw_table(self, title, df_original, is_crosstab=False, font_size=8):
        if df_original.empty: return
        df = df_original.copy()
        if is_crosstab: 
            df = df.replace(0, '-')
            if df.index.name: df.reset_index(inplace=True)
        
        if self.get_y() + (7 * (len(df) + 1) + 10) > self.h - self.b_margin: 
            self.add_page(orientation=self.cur_orientation)
            
        self.draw_section_title(title)
        
        df_formatted = df.copy()
        for col in df_formatted.columns:
            if pd.api.types.is_numeric_dtype(df_formatted[col]) and col not in ['Nº pers.', 'Antigüedad', 'Edad']:
                if "Prom." in str(col):
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{round(x):.0f}" if isinstance(x, (int, float)) else x)
                else:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) else x)
        
        # Cálculo de anchos proporcionales
        widths = {col: max(self.get_string_width(str(col)) + 6, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x[:35])).max() + 6) for col in df_formatted.columns}
        total_width = sum(widths.values())
        if total_width > self.page_width:
            scaling_factor = self.page_width / total_width
            widths = {k: v * scaling_factor for k, v in widths.items()}
        
        self.set_font("Arial", "B", font_size)
        self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
        self.set_text_color(255, 255, 255)
        for col in df_formatted.columns:
            self.cell(widths[col], 7, str(col)[:30], 0, 0, "C", True)
        self.ln()
        
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.set_line_width(0.2)
        
        for i, (_, row) in enumerate(df_formatted.iterrows()):
            if self.get_y() + 7 > self.h - self.b_margin:
                self.add_page(orientation=self.cur_orientation)
                self.set_font("Arial", "B", font_size)
                self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
                self.set_text_color(255, 255, 255)
                for col in df_formatted.columns:
                    self.cell(widths[col], 7, str(col)[:30], 0, 0, "C", True)
                self.ln()
                self.set_text_color(*COLOR_TEXTO_CUERPO)
            
            fill = i % 2 == 1
            self.set_font("Arial", "B" if "Total" in str(row.iloc[0]) else "", font_size)
            self.set_fill_color(*COLOR_GRIS_FONDO_FILA)
            for col in df_formatted.columns:
                val = str(row[col])
                if len(val) > 40: val = val[:38] + ".."
                self.cell(widths[col], 7, val, 'T', 0, "C", fill)
            self.ln()
        self.ln(8)

# --- 2. FUNCIONES DE CÁLCULO ---
def calcular_años(fecha_inicio, fecha_fin):
    if pd.isna(fecha_inicio) or pd.isna(fecha_fin): return 0
    return (fecha_fin - fecha_inicio).days / 365.25

def generar_resumen_completo(df_datos, index_col='Categoría', columns_col='Línea', incluir_promedios=True):
    if df_datos.empty: return pd.DataFrame()
    resumen = pd.crosstab(df_datos[index_col], df_datos[columns_col], margins=True, margins_name="Total")
    if incluir_promedios and 'Antigüedad' in df_datos.columns and 'Edad' in df_datos.columns:
        promedios = df_datos.groupby(index_col).agg({'Antigüedad': 'mean', 'Edad': 'mean'})
        promedios.loc['Total', 'Antigüedad'] = df_datos['Antigüedad'].mean()
        promedios.loc['Total', 'Edad'] = df_datos['Edad'].mean()
        resumen['Antig. Prom.'] = promedios['Antigüedad']
        resumen['Edad Prom.'] = promedios['Edad']
    return resumen

def normalizar_query(df):
    """Estandariza nombres y tipos de columnas de una query SAP."""
    df = df.copy()
    mapping = {
        'Gr.prof.': 'Categoría', 
        'División de personal': 'Línea', 
        'Division de personal': 'Línea',
        'Motivo de la medida': 'Motivo de Baja'
    }
    df.rename(columns=mapping, inplace=True)
    if 'Nº pers.' in df.columns:
        df['Nº pers.'] = df['Nº pers.'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    for col in ['Fecha', 'Desde', 'Fecha nac.']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
    orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
    if 'Línea' in df.columns: df['Línea'] = pd.Categorical(df['Línea'], categories=orden_lineas, ordered=True)
    if 'Categoría' in df.columns: df['Categoría'] = pd.Categorical(df['Categoría'], categories=orden_categorias, ordered=True)
    return df

def procesar_recategorizaciones(df_hoy, df_ant_activos):
    df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()
    cols_base = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría']
    if 'Desde' in df_act_hoy.columns: cols_base.append('Desde')
    if 'Motivo de Baja' in df_act_hoy.columns: cols_base.append('Motivo de Baja')

    if 'Categoría' in df_act_hoy.columns and 'Categoría' in df_ant_activos.columns:
        df_cmp = pd.merge(
            df_act_hoy[cols_base],
            df_ant_activos[['Nº pers.', 'Categoría']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )
        df_recat = df_cmp[df_cmp['Categoría_Actual'] != df_cmp['Categoría_Anterior']].copy()
        df_recat.rename(columns={
            'Categoría_Anterior': 'Categoría Anterior', 
            'Categoría_Actual': 'Categoría Actual',
            'Motivo de Baja': 'Motivo de la medida'
        }, inplace=True)
        if 'Desde' in df_recat.columns:
            df_recat['Desde'] = pd.to_datetime(df_recat['Desde'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
        else: df_recat['Desde'] = '-'
        if 'Motivo de la medida' not in df_recat.columns: df_recat['Motivo de la medida'] = '-'
        else: df_recat['Motivo de la medida'] = df_recat['Motivo de la medida'].fillna('-')
        return df_recat
    return pd.DataFrame()

def procesar_cambios_linea(df_hoy, df_ant_activos):
    df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()
    cols_base = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea']
    if 'Desde' in df_act_hoy.columns: cols_base.append('Desde')
    if 'Motivo de Baja' in df_act_hoy.columns: cols_base.append('Motivo de Baja')

    if 'Línea' in df_act_hoy.columns and 'Línea' in df_ant_activos.columns:
        df_cmp = pd.merge(
            df_act_hoy[cols_base],
            df_ant_activos[['Nº pers.', 'Línea']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )
        df_cambio_l = df_cmp[df_cmp['Línea_Actual'] != df_cmp['Línea_Anterior']].copy()
        df_cambio_l.rename(columns={
            'Línea_Anterior': 'Línea Anterior', 
            'Línea_Actual': 'Línea Actual',
            'Motivo de Baja': 'Motivo de la medida'
        }, inplace=True)
        if 'Desde' in df_cambio_l.columns:
            df_cambio_l['Desde'] = pd.to_datetime(df_cambio_l['Desde'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
        else: df_cambio_l['Desde'] = '-'
        if 'Motivo de la medida' not in df_cambio_l.columns: df_cambio_l['Motivo de la medida'] = '-'
        else: df_cambio_l['Motivo de la medida'] = df_cambio_l['Motivo de la medida'].fillna('-')
        return df_cambio_l
    return pd.DataFrame()

def procesar_reubicados_con_maestro(df_reub_desaparecidos, df_reub_maestro):
    """Cruza los legajos desaparecidos con el archivo REUBICADOS."""
    if df_reub_desaparecidos.empty:
        return pd.DataFrame()
    
    df_base_reub = df_reub_desaparecidos.copy()
    df_base_reub.rename(columns={'Línea': 'Línea Anterior', 'Categoría': 'Gr.prof. Anterior'}, inplace=True)
    
    if df_reub_maestro.empty:
        df_base_reub['Desde'] = '-'
        df_base_reub['Línea Nueva'] = '-'
        df_base_reub['Gr.prof. Nuevo'] = '-'
        df_base_reub['Área/ Posicion/ Función Nueva'] = '-'
        df_base_reub['Nuevo CCT'] = '-'
        return df_base_reub

    m = df_reub_maestro.copy()
    # Identificar columna de legajo en maestro
    col_leg = 'Legajo' if 'Legajo' in m.columns else ('Nº pers.' if 'Nº pers.' in m.columns else m.columns[0])
    m[col_leg] = m[col_leg].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    cols_a_traer = [col_leg]
    for c in ['Fecha', 'Línea Nueva', 'Gr.prof. Nuevo', 'Área/ Posicion/ Función Nueva', 'Nuevo CCT']:
        if c in m.columns: cols_a_traer.append(c)

    merged = pd.merge(df_base_reub, m[cols_a_traer], left_on='Nº pers.', right_on=col_leg, how='left')

    if 'Fecha' in merged.columns:
        merged['Desde'] = pd.to_datetime(merged['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
    else:
        merged['Desde'] = '-'

    for c in ['Línea Nueva', 'Gr.prof. Nuevo', 'Área/ Posicion/ Función Nueva', 'Nuevo CCT']:
        if c not in merged.columns: merged[c] = '-'
        else: merged[c] = merged[c].fillna('-')

    columnas_ordenadas = [
        'Nº pers.', 'Apellido', 'Nombre de pila', 'Desde',
        'Línea Anterior', 'Gr.prof. Anterior', 
        'Línea Nueva', 'Gr.prof. Nuevo', 
        'Área/ Posicion/ Función Nueva', 'Nuevo CCT'
    ]
    return merged[[col for col in columnas_ordenadas if col in merged.columns]]

# --- 3. GENERADOR DE PDF ---
def crear_pdf_reporte_diario(fecha_str, df_altas, df_bajas, res_altas, res_bajas, res_activos, res_bajas_linea, res_bajas_cat, df_reub=None, df_recat=None, df_cambio_linea=None):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = "Resumen Diario de Dotación"
    pdf.add_page()
    pdf.draw_section_title(f"Indicadores del Día: {fecha_str}")
    
    total_act = f"{res_activos.loc['Total', 'Total']:,}".replace(',', '.') if not res_activos.empty else "0"
    has_reub = df_reub is not None and not df_reub.empty
    has_recat = df_recat is not None and not df_recat.empty
    has_linea = df_cambio_linea is not None and not df_cambio_linea.empty

    # Detección de demoras
    notas_demora = []
    if not df_bajas.empty:
        ayer = pd.to_datetime(datetime.now()) - pd.Timedelta(days=1)
        for _, row in df_bajas.iterrows():
            f_baja = row.get('Desde_DT') if 'Desde_DT' in row else pd.to_datetime(row.get('Desde'), errors='coerce')
            if pd.notna(f_baja) and f_baja.date() < ayer.date():
                legajo = str(row.get('Nº pers.', ''))
                f_str = f_baja.strftime('%d/%m/%Y')
                notas_demora.append(f"El Legajo {legajo} corresponde a una baja del {f_str} registrada con demora.")

    num_kpis = 3 + (1 if has_reub else 0) + (1 if has_recat else 0) + (1 if has_linea else 0)
    k_w = pdf.page_width / (num_kpis + 0.5)
    sp = (pdf.page_width - (k_w * num_kpis)) / max(1, (num_kpis - 1))
    
    y = pdf.get_y()
    curr_x = pdf.l_margin
    
    pdf.draw_kpi_box("Dotación Activa", total_act, (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    pdf.draw_kpi_box("Altas del Período", '-' if len(df_altas) == 0 else str(len(df_altas)), (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    titulo_bajas = "Bajas del Período *" if len(notas_demora) > 0 else "Bajas del Período"
    pdf.draw_kpi_box(titulo_bajas, '-' if len(df_bajas) == 0 else str(len(df_bajas)), (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    if has_reub:
        pdf.draw_kpi_box("Reubicados", str(len(df_reub)), (255, 165, 0), curr_x, y, width=k_w)
        curr_x += k_w + sp
        
    if has_recat:
        pdf.draw_kpi_box("Cambio Categoría", str(len(df_recat)), COLOR_CELESTE_PASTEL, curr_x, y, width=k_w)
        curr_x += k_w + sp

    if has_linea:
        pdf.draw_kpi_box("Cambio Línea", str(len(df_cambio_linea)), COLOR_AZUL_PASTEL_OSCURO, curr_x, y, width=k_w)
    
    pdf.ln(20)

    if len(notas_demora) > 0:
        pdf.set_font("Arial", "I", 8)
        pdf.set_text_color(180, 50, 50)
        for nota in notas_demora:
            pdf.cell(0, 4, f"* Nota: {nota}", ln=True, align="L")
        pdf.ln(2)

    pdf.draw_table("Composición de la Dotación Activa", res_activos, is_crosstab=True)
    pdf.draw_table(f"Resumen de Bajas (Día: {fecha_str})", res_bajas, is_crosstab=True)
    pdf.draw_table("Motivos de Baja por Línea", res_bajas_linea, is_crosstab=True)
    pdf.draw_table("Motivos de Baja por Categoría", res_bajas_cat, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Día: {fecha_str})", res_altas, is_crosstab=True)
    
    if not df_bajas.empty: 
        pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de Baja', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not df_altas.empty: 
        pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    
    # Detalle de Reubicados con datos maestros
    if has_reub: 
        pdf.draw_table("Detalle de Reubicados", df_reub, font_size=7)
    
    if has_recat: 
        pdf.draw_table("Detalle Cambios de Categoría", df_recat[['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría Anterior', 'Categoría Actual', 'Desde', 'Motivo de la medida']])
    if has_linea: 
        pdf.draw_table("Detalle Cambios de Línea", df_cambio_linea[['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea Anterior', 'Línea Actual', 'Desde', 'Motivo de la medida']])
    
    return pdf.output(dest='S').encode('latin-1', 'replace')

# --- 4. INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Control de Dotación - Diario", layout="wide")
st.title("📋 Control Diario de Dotación")
st.markdown("Cargue los archivos correspondientes para procesar las novedades del día.")

c1, c2, c3 = st.columns(3)
file_hoy = c1.file_uploader("1️⃣ Archivo: query-hoy", type=['xlsx'], help="Query de SAP con la foto de hoy (Sheet1)")
file_ant = c2.file_uploader("2️⃣ Archivo: query-ant", type=['xlsx'], help="Query de SAP con la foto anterior/ayer (Sheet1)")
file_reub = c3.file_uploader("3️⃣ Archivo: reubicados", type=['xlsx'], help="Archivo maestro de Reubicados (pestaña REUBICADOS)")

if file_hoy and file_ant:
    try:
        # Carga de archivos individuales
        df_hoy_raw = pd.read_excel(file_hoy, sheet_name=0, engine='openpyxl')
        df_ant_raw = pd.read_excel(file_ant, sheet_name=0, engine='openpyxl')
        
        df_reub_maestro = pd.DataFrame()
        if file_reub:
            try:
                # Intenta leer pestaña REUBICADOS, si no la primera
                try: df_reub_maestro = pd.read_excel(file_reub, sheet_name='REUBICADOS', engine='openpyxl')
                except: df_reub_maestro = pd.read_excel(file_reub, sheet_name=0, engine='openpyxl')
            except Exception as e:
                st.warning(f"No se pudo leer el archivo maestro de reubicados: {e}")

        # Normalización
        df_hoy = normalizar_query(df_hoy_raw)
        df_ant = normalizar_query(df_ant_raw)

        # Foto Activos de Ayer
        df_act_ant = df_ant[df_ant['Status ocupación'] == 'Activo'].copy()
        legs_act_ant = set(df_act_ant['Nº pers.'])
        legs_hoy = set(df_hoy['Nº pers.'])

        # Detección de Altas y Bajas
        df_alt_r = df_hoy[~df_hoy['Nº pers.'].isin(legs_act_ant) & (df_hoy['Status ocupación'] == 'Activo')].copy()
        df_baj_r = df_hoy[df_hoy['Nº pers.'].isin(legs_act_ant) & (df_hoy['Status ocupación'] == 'Dado de baja')].copy()

        if not df_baj_r.empty: 
            df_baj_r['Desde'] = df_baj_r['Desde'] - pd.Timedelta(days=1)
            df_baj_r = df_baj_r.sort_values(by='Desde', ascending=True)
        if not df_alt_r.empty: 
            df_alt_r = df_alt_r.sort_values(by='Fecha', ascending=True)

        # Detección de Reubicados (Desaparecieron de la query hoy pero estaban activos ayer)
        desap_legs = legs_act_ant - legs_hoy
        df_reub_desap = df_act_ant[df_act_ant['Nº pers.'].isin(desap_legs)].copy()
        df_reub_completo = procesar_reubicados_con_maestro(df_reub_desap, df_reub_maestro)

        # Cálculo de Antigüedad y Edad
        hoy = pd.to_datetime(datetime.now())
        
        # Procesar Bajas
        df_baj = df_baj_r.copy()
        if not df_baj.empty:
            df_baj['Antigüedad'] = df_baj.apply(lambda r: calcular_años(r['Fecha'], r['Desde']), axis=1)
            df_baj['Edad'] = df_baj.apply(lambda r: calcular_años(r['Fecha nac.'], r['Desde']), axis=1)
            df_baj_vis = df_baj.copy()
            df_baj_vis['Antigüedad'] = df_baj_vis['Antigüedad'].apply(lambda x: int(round(x)))
            df_baj_vis['Fecha nac.'] = df_baj_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')
            df_baj_vis['Desde_DT'] = df_baj_vis['Desde']
            df_baj_vis['Desde'] = df_baj_vis['Desde'].dt.strftime('%d/%m/%Y')
        else: df_baj_vis = pd.DataFrame()

        # Procesar Altas
        df_alt = df_alt_r.copy()
        if not df_alt.empty:
            df_alt['Antigüedad'] = df_alt.apply(lambda r: calcular_años(r['Fecha'], hoy), axis=1)
            df_alt['Edad'] = df_alt.apply(lambda r: calcular_años(r['Fecha nac.'], hoy), axis=1)
            df_alt_vis = df_alt.copy()
            df_alt_vis['Fecha'] = df_alt_vis['Fecha'].dt.strftime('%d/%m/%Y')
            df_alt_vis['Fecha nac.'] = df_alt_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')
        else: df_alt_vis = pd.DataFrame()

        # Cambios de Categoría y Línea
        df_recat = procesar_recategorizaciones(df_hoy, df_act_ant)
        df_cambio_l = procesar_cambios_linea(df_hoy, df_act_ant)

        # Resúmenes crosstab
        df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()
        df_act_hoy['Antigüedad'] = df_act_hoy.apply(lambda r: calcular_años(r['Fecha'], hoy), axis=1)
        df_act_hoy['Edad'] = df_act_hoy.apply(lambda r: calcular_años(r['Fecha nac.'], hoy), axis=1)

        res_act = generar_resumen_completo(df_act_hoy)
        res_alt = generar_resumen_completo(df_alt, incluir_promedios=False)
        res_baj = generar_resumen_completo(df_baj)
        res_baj_linea = pd.crosstab(df_baj['Motivo de Baja'], df_baj['Línea'], margins=True, margins_name="Total") if not df_baj.empty else pd.DataFrame()
        res_baj_cat = pd.crosstab(df_baj['Motivo de Baja'], df_baj['Categoría'], margins=True, margins_name="Total") if not df_baj.empty else pd.DataFrame()

        # Generar Reporte PDF
        pdf_bytes = crear_pdf_reporte_diario(
            datetime.now().strftime('%d/%m/%Y'),
            df_alt_vis, df_baj_vis,
            res_alt, res_baj, res_act,
            res_baj_linea, res_baj_cat,
            df_reub_completo, df_recat, df_cambio_l
        )

        st.success("✅ Archivos procesados con éxito.")
        st.download_button(
            "📄 Descargar Reporte Diario de Dotación",
            pdf_bytes,
            f"Reporte_Diario_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",
            "application/pdf"
        )

    except Exception as e:
        st.error(f"Error al procesar los datos: {e}")
else:
    st.info("💡 Por favor, cargue al menos 'query-hoy' y 'query-ant' para generar el reporte diario.")
