import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import io
import os

# ==============================================================================
# 0. RUTAS FIJAS CONFIGURADAS
# ==============================================================================
RUTA_QUERY_HOY = r"C:\Users\florencia.flores\Desktop\Cambio Cat-Linea\query-hoy.xlsx"
RUTA_QUERY_ANT = r"C:\Users\florencia.flores\Desktop\Cambio Cat-Linea\query-ant.xlsx"
RUTA_REUBICADOS = r"Q:\GGYDPC\SGyCAPC\Administración del PC\REUBICADOS.xlsx"

# --- 1. CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)[cite: 2]
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)[cite: 2]
COLOR_GRIS_FONDO_FILA = (240, 242, 246)[cite: 2]
COLOR_GRIS_LINEA = (220, 220, 220)[cite: 2]
COLOR_TEXTO_TITULO = (0, 51, 102)[cite: 2]
COLOR_TEXTO_CUERPO = (50, 50, 50)[cite: 2]
COLOR_CELESTE_PASTEL = (186, 225, 255)       # Celeste pastel para Cambio Categoría[cite: 2]
COLOR_AZUL_PASTEL_OSCURO = (120, 180, 235)  # Celeste/Azul más oscuro para Cambio Línea[cite: 2]

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)[cite: 2]
        self.page_width = self.w - 2 * self.l_margin[cite: 2]
        self.report_title = "Resumen de Dotación"[cite: 2]

    def header(self):
        self.set_font("Arial", "B", 18)[cite: 2]
        self.set_text_color(*COLOR_TEXTO_TITULO)[cite: 2]
        self.cell(0, 10, self.report_title, 0, 0, "C")[cite: 2]
        self.ln(15)[cite: 2]

    def footer(self):
        self.set_y(-15)[cite: 2]
        self.set_font("Arial", "I", 8)[cite: 2]
        self.set_text_color(128, 128, 128)[cite: 2]
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")[cite: 2]

    def draw_section_title(self, title):
        self.set_font("Arial", "B", 13)[cite: 2]
        self.set_text_color(*COLOR_TEXTO_TITULO)[cite: 2]
        self.cell(0, 8, title, ln=True, align="L")[cite: 2]
        self.set_draw_color(*COLOR_AZUL_INSTITUCIONAL)[cite: 2]
        self.set_line_width(0.5)[cite: 2]
        self.line(self.get_x(), self.get_y(), self.get_x() + self.page_width, self.get_y())[cite: 2]
        self.ln(4)[cite: 2]

    def draw_kpi_box(self, title, value, color, x, y, width=80):
        kpi_height = 16[cite: 2]
        self.set_xy(x, y)[cite: 2]
        self.set_fill_color(*color)[cite: 2]
        self.cell(width, 1.5, "", fill=True, ln=False, border=0)[cite: 2]
        self.set_xy(x, y + 1.5)[cite: 2]
        self.set_fill_color(255, 255, 255)[cite: 2]
        self.set_draw_color(*COLOR_GRIS_LINEA)[cite: 2]
        self.cell(width, kpi_height - 1.5, "", border=1, fill=True)[cite: 2]
        self.set_xy(x, y + 3)[cite: 2]
        self.set_font('Arial', '', 9)[cite: 2]
        self.set_text_color(*COLOR_TEXTO_CUERPO)[cite: 2]
        self.cell(width, 8, title, align='C')[cite: 2]
        self.set_xy(x, y + 8)[cite: 2]
        self.set_font('Arial', 'B', 15)[cite: 2]
        self.set_text_color(*COLOR_TEXTO_TITULO)[cite: 2]
        self.cell(width, 10, str(value), align='C')[cite: 2]

    def draw_table(self, title, df_original, is_crosstab=False, font_size=8):
        if df_original.empty: return[cite: 2]
        df = df_original.copy()[cite: 2]
        if is_crosstab: 
            df = df.replace(0, '-')[cite: 2]
            if df.index.name: df.reset_index(inplace=True)[cite: 2]
        
        if self.get_y() + (7 * (len(df) + 1) + 10) > self.h - self.b_margin: 
            self.add_page(orientation=self.cur_orientation)[cite: 2]
            
        self.draw_section_title(title)[cite: 2]
        
        df_formatted = df.copy()[cite: 2]
        for col in df_formatted.columns:[cite: 2]
            if pd.api.types.is_numeric_dtype(df_formatted[col]) and col not in ['Nº pers.', 'Antigüedad', 'Edad']:[cite: 2]
                if "Prom." in str(col):[cite: 2]
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{round(x):.0f}" if isinstance(x, (int, float)) else x)[cite: 2]
                else:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) else x)[cite: 2]
        
        widths = {col: max(self.get_string_width(str(col)) + 5, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x[:30])).max() + 5) for col in df_formatted.columns}
        total_width = sum(widths.values())[cite: 2]
        if total_width > self.page_width:[cite: 2]
            scaling_factor = self.page_width / total_width[cite: 2]
            widths = {k: v * scaling_factor for k, v in widths.items()}[cite: 2]
        
        self.set_font("Arial", "B", font_size)[cite: 2]
        self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)[cite: 2]
        self.set_text_color(255, 255, 255)[cite: 2]
        for col in df_formatted.columns:[cite: 2]
            self.cell(widths[col], 7, str(col)[:25], 0, 0, "C", True)[cite: 2]
        self.ln()[cite: 2]
        
        self.set_text_color(*COLOR_TEXTO_CUERPO)[cite: 2]
        self.set_draw_color(*COLOR_GRIS_LINEA)[cite: 2]
        self.set_line_width(0.2)[cite: 2]
        
        for i, (_, row) in enumerate(df_formatted.iterrows()):[cite: 2]
            if self.get_y() + 7 > self.h - self.b_margin:[cite: 2]
                self.add_page(orientation=self.cur_orientation)[cite: 2]
                self.set_font("Arial", "B", font_size)[cite: 2]
                self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)[cite: 2]
                self.set_text_color(255, 255, 255)[cite: 2]
                for col in df_formatted.columns:[cite: 2]
                    self.cell(widths[col], 7, str(col)[:25], 0, 0, "C", True)[cite: 2]
                self.ln()[cite: 2]
                self.set_text_color(*COLOR_TEXTO_CUERPO)[cite: 2]
            
            fill = i % 2 == 1[cite: 2]
            self.set_font("Arial", "B" if "Total" in str(row.iloc[0]) else "", font_size)[cite: 2]
            self.set_fill_color(*COLOR_GRIS_FONDO_FILA)[cite: 2]
            for col in df_formatted.columns:[cite: 2]
                val = str(row[col])[cite: 2]
                if len(val) > 42: val = val[:40] + ".."
                self.cell(widths[col], 7, val, 'T', 0, "C", fill)[cite: 2]
            self.ln()[cite: 2]
        self.ln(8)[cite: 2]

# --- 2. LÓGICA DE CÁLCULO ---
def calcular_años(fecha_inicio, fecha_fin):
    if pd.isna(fecha_inicio) or pd.isna(fecha_fin): return 0[cite: 2]
    return (fecha_fin - fecha_inicio).days / 365.25[cite: 2]

def generar_resumen_completo(df_datos, index_col='Categoría', columns_col='Línea', incluir_promedios=True):
    if df_datos.empty: return pd.DataFrame()[cite: 2]
    resumen = pd.crosstab(df_datos[index_col], df_datos[columns_col], margins=True, margins_name="Total")[cite: 2]
    if incluir_promedios and 'Antigüedad' in df_datos.columns and 'Edad' in df_datos.columns:[cite: 2]
        promedios = df_datos.groupby(index_col).agg({'Antigüedad': 'mean', 'Edad': 'mean'})[cite: 2]
        promedios.loc['Total', 'Antigüedad'] = df_datos['Antigüedad'].mean()[cite: 2]
        promedios.loc['Total', 'Edad'] = df_datos['Edad'].mean()[cite: 2]
        resumen['Antig. Prom.'] = promedios['Antigüedad'][cite: 2]
        resumen['Edad Prom.'] = promedios['Edad'][cite: 2]
    return resumen[cite: 2]

def limpiar_legajo(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def normalizar_query(df):
    df = df.copy()[cite: 2]
    df.columns = [str(c).strip() for c in df.columns]
    mapping = {
        'Gr.prof.': 'Categoría', 
        'División de personal': 'Línea', 
        'Division de personal': 'Línea',
        'Motivo de la medida': 'Motivo de Baja'
    }[cite: 2]
    df.rename(columns=mapping, inplace=True)[cite: 2]
    if 'Nº pers.' in df.columns:
        df['Nº pers.'] = df['Nº pers.'].apply(limpiar_legajo)
    for col in ['Fecha', 'Desde', 'Fecha nac.']:[cite: 2]
        if col in df.columns:[cite: 2]
            df[col] = pd.to_datetime(df[col], errors='coerce')[cite: 2]
    orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL'][cite: 2]
    orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C'][cite: 2]
    if 'Línea' in df.columns: df['Línea'] = pd.Categorical(df['Línea'], categories=orden_lineas, ordered=True)[cite: 2]
    if 'Categoría' in df.columns: df['Categoría'] = pd.Categorical(df['Categoría'], categories=orden_categorias, ordered=True)[cite: 2]
    return df[cite: 2]

def procesar_recategorizaciones(df_hoy, df_ant_activos):
    df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()[cite: 2]
    cols_base = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría'][cite: 2]
    if 'Desde' in df_act_hoy.columns: cols_base.append('Desde')[cite: 2]
    if 'Motivo de Baja' in df_act_hoy.columns: cols_base.append('Motivo de Baja')[cite: 2]

    if 'Categoría' in df_act_hoy.columns and 'Categoría' in df_ant_activos.columns:[cite: 2]
        df_cmp = pd.merge(
            df_act_hoy[cols_base],
            df_ant_activos[['Nº pers.', 'Categoría']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )[cite: 2]
        df_recat = df_cmp[df_cmp['Categoría_Actual'] != df_cmp['Categoría_Anterior']].copy()[cite: 2]
        df_recat.rename(columns={
            'Categoría_Anterior': 'Categoría Anterior', 
            'Categoría_Actual': 'Categoría Actual',
            'Motivo de Baja': 'Motivo de la medida'
        }, inplace=True)[cite: 2]
        if 'Desde' in df_recat.columns:[cite: 2]
            df_recat['Desde'] = pd.to_datetime(df_recat['Desde'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')[cite: 2]
        else: df_recat['Desde'] = '-'[cite: 2]
        if 'Motivo de la medida' not in df_recat.columns: df_recat['Motivo de la medida'] = '-'[cite: 2]
        else: df_recat['Motivo de la medida'] = df_recat['Motivo de la medida'].fillna('-')[cite: 2]
        return df_recat[cite: 2]
    return pd.DataFrame()[cite: 2]

def procesar_cambios_linea(df_hoy, df_ant_activos):
    df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()[cite: 2]
    cols_base = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea'][cite: 2]
    if 'Desde' in df_act_hoy.columns: cols_base.append('Desde')[cite: 2]
    if 'Motivo de Baja' in df_act_hoy.columns: cols_base.append('Motivo de Baja')[cite: 2]

    if 'Línea' in df_act_hoy.columns and 'Línea' in df_ant_activos.columns:[cite: 2]
        df_cmp = pd.merge(
            df_act_hoy[cols_base],
            df_ant_activos[['Nº pers.', 'Línea']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )[cite: 2]
        df_cambio_l = df_cmp[df_cmp['Línea_Actual'] != df_cmp['Línea_Anterior']].copy()[cite: 2]
        df_cambio_l.rename(columns={
            'Línea_Anterior': 'Línea Anterior', 
            'Línea_Actual': 'Línea Actual',
            'Motivo de Baja': 'Motivo de la medida'
        }, inplace=True)[cite: 2]
        if 'Desde' in df_cambio_l.columns:[cite: 2]
            df_cambio_l['Desde'] = pd.to_datetime(df_cambio_l['Desde'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')[cite: 2]
        else: df_cambio_l['Desde'] = '-'[cite: 2]
        if 'Motivo de la medida' not in df_cambio_l.columns: df_cambio_l['Motivo de la medida'] = '-'[cite: 2]
        else: df_cambio_l['Motivo de la medida'] = df_cambio_l['Motivo de la medida'].fillna('-')[cite: 2]
        return df_cambio_l[cite: 2]
    return pd.DataFrame()[cite: 2]

def procesar_reubicados_con_maestro(df_reub_desap, df_reub_maestro):
    if df_reub_desap.empty:
        return pd.DataFrame()[cite: 2]
    
    cols_base = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría']
    base = df_reub_desap[[c for c in cols_base if c in df_reub_desap.columns]].copy()
    base.rename(columns={'Línea': 'Línea Anterior', 'Categoría': 'Gr.prof. Anterior'}, inplace=True)[cite: 2]

    cols_finales = [
        'Nº pers.', 'Apellido', 'Nombre de pila', 'Desde',
        'Línea Anterior', 'Gr.prof. Anterior', 
        'Línea Nueva', 'Gr.prof. Nuevo', 
        'Área/ Posicion/ Función Nueva', 'Nuevo CCT'
    ][cite: 2]
    
    if df_reub_maestro.empty:
        for c in ['Desde', 'Línea Nueva', 'Gr.prof. Nuevo', 'Área/ Posicion/ Función Nueva', 'Nuevo CCT']:
            base[c] = '-'
        return base[cols_finales]

    m = df_reub_maestro.copy()[cite: 2]
    m.columns = [str(c).strip() for c in m.columns]
    
    col_leg = 'Legajo' if 'Legajo' in m.columns else ('Nº pers.' if 'Nº pers.' in m.columns else m.columns[0])
    m['LEG_MATCH'] = m[col_leg].apply(limpiar_legajo)

    col_fecha = 'Fecha' if 'Fecha' in m.columns else next((c for c in m.columns if 'fecha' in c.lower()), None)
    if col_fecha:
        m['Desde'] = pd.to_datetime(m[col_fecha], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y').fillna('-')
    else:
        m['Desde'] = '-'

    cols_m = ['LEG_MATCH', 'Desde']
    for c_std in ['Línea Nueva', 'Gr.prof. Nuevo', 'Área/ Posicion/ Función Nueva', 'Nuevo CCT']:
        match = next((c for c in m.columns if c_std.lower() in c.lower()), None)
        if match:
            m.rename(columns={match: c_std}, inplace=True)
            cols_m.append(c_std)

    merged = pd.merge(base, m[cols_m].drop_duplicates(subset=['LEG_MATCH']), left_on='Nº pers.', right_on='LEG_MATCH', how='left')

    for c in cols_finales:
        if c not in merged.columns:
            merged[c] = '-'
        else:
            merged[c] = merged[c].fillna('-')

    return merged[cols_finales]

# --- 3. GENERADOR DE PDF ---
def crear_pdf_reporte_diario(fecha_str, df_altas, df_bajas, res_altas, res_bajas, res_activos, res_bajas_linea, res_bajas_cat, df_reub=None, df_recat=None, df_cambio_linea=None):
    pdf = PDF(orientation='L', unit='mm', format='A4')[cite: 2]
    pdf.report_title = "Resumen Diario de Dotación"[cite: 2]
    pdf.add_page()[cite: 2]
    pdf.draw_section_title(f"Indicadores del Día: {fecha_str}")[cite: 2]
    
    total_act = f"{res_activos.loc['Total', 'Total']:,}".replace(',', '.') if not res_activos.empty else "0"[cite: 2]
    has_reub = df_reub is not None and not df_reub.empty[cite: 2]
    has_recat = df_recat is not None and not df_recat.empty[cite: 2]
    has_linea = df_cambio_linea is not None and not df_cambio_linea.empty[cite: 2]

    notas_demora = [][cite: 2]
    if not df_bajas.empty:[cite: 2]
        ayer = pd.to_datetime(datetime.now()) - pd.Timedelta(days=1)[cite: 2]
        for _, row in df_bajas.iterrows():[cite: 2]
            f_baja = row.get('Desde_DT') if 'Desde_DT' in row else pd.to_datetime(row.get('Desde'), errors='coerce')[cite: 2]
            if pd.notna(f_baja) and f_baja.date() < ayer.date():[cite: 2]
                legajo = str(row.get('Nº pers.', ''))[cite: 2]
                f_str = f_baja.strftime('%d/%m/%Y')[cite: 2]
                notas_demora.append(f"El Legajo {legajo} corresponde a una baja del {f_str} registrada con demora.")[cite: 2]

    num_kpis = 3 + (1 if has_reub else 0) + (1 if has_recat else 0) + (1 if has_linea else 0)[cite: 2]
    k_w = pdf.page_width / (num_kpis + 0.5)[cite: 2]
    sp = (pdf.page_width - (k_w * num_kpis)) / max(1, (num_kpis - 1))[cite: 2]
    
    y = pdf.get_y()[cite: 2]
    curr_x = pdf.l_margin[cite: 2]
    
    pdf.draw_kpi_box("Dotación Activa", total_act, (200, 200, 200), curr_x, y, width=k_w)[cite: 2]
    curr_x += k_w + sp[cite: 2]
    
    pdf.draw_kpi_box("Altas del Período", '-' if len(df_altas) == 0 else str(len(df_altas)), (200, 200, 200), curr_x, y, width=k_w)[cite: 2]
    curr_x += k_w + sp[cite: 2]
    
    titulo_bajas = "Bajas del Período *" if len(notas_demora) > 0 else "Bajas del Período"[cite: 2]
    pdf.draw_kpi_box(titulo_bajas, '-' if len(df_bajas) == 0 else str(len(df_bajas)), (200, 200, 200), curr_x, y, width=k_w)[cite: 2]
    curr_x += k_w + sp[cite: 2]
    
    if has_reub:
        pdf.draw_kpi_box("Reubicados", str(len(df_reub)), (255, 165, 0), curr_x, y, width=k_w)[cite: 2]
        curr_x += k_w + sp[cite: 2]
        
    if has_recat:
        pdf.draw_kpi_box("Cambio Categoría", str(len(df_recat)), COLOR_CELESTE_PASTEL, curr_x, y, width=k_w)[cite: 2]
        curr_x += k_w + sp[cite: 2]

    if has_linea:
        pdf.draw_kpi_box("Cambio Línea", str(len(df_cambio_linea)), COLOR_AZUL_PASTEL_OSCURO, curr_x, y, width=k_w)[cite: 2]
    
    pdf.ln(20)[cite: 2]

    if len(notas_demora) > 0:[cite: 2]
        pdf.set_font("Arial", "I", 8)[cite: 2]
        pdf.set_text_color(180, 50, 50)[cite: 2]
        for nota in notas_demora:[cite: 2]
            pdf.cell(0, 4, f"* Nota: {nota}", ln=True, align="L")[cite: 2]
        pdf.ln(2)[cite: 2]

    pdf.draw_table("Composición de la Dotación Activa", res_activos, is_crosstab=True)[cite: 2]
    pdf.draw_table(f"Resumen de Bajas (Día: {fecha_str})", res_bajas, is_crosstab=True)[cite: 2]
    pdf.draw_table("Motivos de Baja por Línea", res_bajas_linea, is_crosstab=True)[cite: 2]
    pdf.draw_table("Motivos de Baja por Categoría", res_bajas_cat, is_crosstab=True)[cite: 2]
    pdf.draw_table(f"Resumen de Altas (Día: {fecha_str})", res_altas, is_crosstab=True)[cite: 2]
    
    if not df_bajas.empty: 
        pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de Baja', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])[cite: 2]
    if not df_altas.empty: 
        pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])[cite: 2]
    
    if has_reub: 
        pdf.draw_table("Detalle de Reubicados", df_reub, font_size=7)[cite: 2]
    
    if has_recat: 
        pdf.draw_table("Detalle Cambios de Categoría", df_recat[['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría Anterior', 'Categoría Actual', 'Desde', 'Motivo de la medida']])[cite: 2]
    if has_linea: 
        pdf.draw_table("Detalle Cambios de Línea", df_cambio_linea[['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea Anterior', 'Línea Actual', 'Desde', 'Motivo de la medida']])[cite: 2]
    
    return pdf.output(dest='S').encode('latin-1', 'replace')[cite: 2]

# --- 4. INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Control Diario de Dotación", layout="wide")
st.title("⚡ Control Diario Automatizado")

hay_hoy = os.path.exists(RUTA_QUERY_HOY)
hay_ant = os.path.exists(RUTA_QUERY_ANT)
hay_reub = os.path.exists(RUTA_REUBICADOS)

c1, c2, c3 = st.columns(3)
if hay_hoy: c1.success(f"✅ query-hoy detectado\n\n`{RUTA_QUERY_HOY}`")
else: c1.error(f"❌ No se encuentra query-hoy\n\n`{RUTA_QUERY_HOY}`")

if hay_ant: c2.success(f"✅ query-ant detectado\n\n`{RUTA_QUERY_ANT}`")
else: c2.error(f"❌ No se encuentra query-ant\n\n`{RUTA_QUERY_ANT}`")

if hay_reub: c3.success(f"✅ reubicados detectado\n\n`{RUTA_REUBICADOS}`")
else: c3.warning(f"⚠️ reubicados no detectado (Opcional)\n\n`{RUTA_REUBICADOS}`")

st.markdown("---")

if hay_hoy and hay_ant:
    if st.button("🚀 Procesar y Generar Reporte", type="primary"):
        try:
            with st.spinner("Leyendo archivos y generando reporte..."):
                df_hoy_raw = pd.read_excel(RUTA_QUERY_HOY, sheet_name=0, engine='openpyxl')
                df_ant_raw = pd.read_excel(RUTA_QUERY_ANT, sheet_name=0, engine='openpyxl')
                
                df_reub_maestro = pd.DataFrame()[cite: 2]
                if hay_reub:
                    try:
                        try: df_reub_maestro = pd.read_excel(RUTA_REUBICADOS, sheet_name='REUBICADOS', engine='openpyxl')
                        except: df_reub_maestro = pd.read_excel(RUTA_REUBICADOS, sheet_name=0, engine='openpyxl')
                    except Exception as e:
                        st.warning(f"No se pudo leer el archivo de reubicados: {e}")

                df_hoy = normalizar_query(df_hoy_raw)
                df_ant = normalizar_query(df_ant_raw)

                df_act_ant = df_ant[df_ant['Status ocupación'] == 'Activo'].copy()[cite: 2]
                legs_act_ant = set(df_act_ant['Nº pers.'])[cite: 2]
                legs_hoy = set(df_hoy['Nº pers.'])[cite: 2]

                df_alt_r = df_hoy[~df_hoy['Nº pers.'].isin(legs_act_ant) & (df_hoy['Status ocupación'] == 'Activo')].copy()[cite: 2]
                df_baj_r = df_hoy[df_hoy['Nº pers.'].isin(legs_act_ant) & (df_hoy['Status ocupación'] == 'Dado de baja')].copy()[cite: 2]

                if not df_baj_r.empty: 
                    df_baj_r['Desde'] = df_baj_r['Desde'] - pd.Timedelta(days=1)[cite: 2]
                    df_baj_r = df_baj_r.sort_values(by='Desde', ascending=True)[cite: 2]
                if not df_alt_r.empty: 
                    df_alt_r = df_alt_r.sort_values(by='Fecha', ascending=True)[cite: 2]

                desap_legs = legs_act_ant - legs_hoy[cite: 2]
                df_reub_desap = df_act_ant[df_act_ant['Nº pers.'].isin(desap_legs)].copy()[cite: 2]
                df_reub_completo = procesar_reubicados_con_maestro(df_reub_desap, df_reub_maestro)

                hoy = pd.to_datetime(datetime.now())[cite: 2]
                
                df_baj = df_baj_r.copy()[cite: 2]
                if not df_baj.empty:[cite: 2]
                    df_baj['Antigüedad'] = df_baj.apply(lambda r: calcular_años(r['Fecha'], r['Desde']), axis=1)[cite: 2]
                    df_baj['Edad'] = df_baj.apply(lambda r: calcular_años(r['Fecha nac.'], r['Desde']), axis=1)[cite: 2]
                    df_baj_vis = df_baj.copy()[cite: 2]
                    df_baj_vis['Antigüedad'] = df_baj_vis['Antigüedad'].apply(lambda x: int(round(x)))[cite: 2]
                    df_baj_vis['Fecha nac.'] = df_baj_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')[cite: 2]
                    df_baj_vis['Desde_DT'] = df_baj_vis['Desde'][cite: 2]
                    df_baj_vis['Desde'] = df_baj_vis['Desde'].dt.strftime('%d/%m/%Y')[cite: 2]
                else: df_baj_vis = pd.DataFrame()[cite: 2]

                df_alt = df_alt_r.copy()[cite: 2]
                if not df_alt.empty:[cite: 2]
                    df_alt['Antigüedad'] = df_alt.apply(lambda r: calcular_años(r['Fecha'], hoy), axis=1)[cite: 2]
                    df_alt['Edad'] = df_alt.apply(lambda r: calcular_años(r['Fecha nac.'], hoy), axis=1)[cite: 2]
                    df_alt_vis = df_alt.copy()[cite: 2]
                    df_alt_vis['Fecha'] = df_alt_vis['Fecha'].dt.strftime('%d/%m/%Y')[cite: 2]
                    df_alt_vis['Fecha nac.'] = df_alt_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')[cite: 2]
                else: df_alt_vis = pd.DataFrame()[cite: 2]

                df_recat = procesar_recategorizaciones(df_hoy, df_act_ant)[cite: 2]
                df_cambio_l = procesar_cambios_linea(df_hoy, df_act_ant)[cite: 2]

                df_act_hoy = df_hoy[df_hoy['Status ocupación'] == 'Activo'].copy()[cite: 2]
                df_act_hoy['Antigüedad'] = df_act_hoy.apply(lambda r: calcular_años(r['Fecha'], hoy), axis=1)[cite: 2]
                df_act_hoy['Edad'] = df_act_hoy.apply(lambda r: calcular_años(r['Fecha nac.'], hoy), axis=1)[cite: 2]

                res_act = generar_resumen_completo(df_act_hoy)[cite: 2]
                res_alt = generar_resumen_completo(df_alt, incluir_promedios=False)[cite: 2]
                res_baj = generar_resumen_completo(df_baj)[cite: 2]
                res_baj_linea = pd.crosstab(df_baj['Motivo de Baja'], df_baj['Línea'], margins=True, margins_name="Total") if not df_baj.empty else pd.DataFrame()[cite: 2]
                res_baj_cat = pd.crosstab(df_baj['Motivo de Baja'], df_baj['Categoría'], margins=True, margins_name="Total") if not df_baj.empty else pd.DataFrame()[cite: 2]

                pdf_bytes = crear_pdf_reporte_diario(
                    datetime.now().strftime('%d/%m/%Y'),[cite: 2]
                    df_alt_vis, df_baj_vis,[cite: 2]
                    res_alt, res_baj, res_act,[cite: 2]
                    res_baj_linea, res_baj_cat,[cite: 2]
                    df_reub_completo, df_recat, df_cambio_l[cite: 2]
                )

                st.success("✅ Reporte procesado exitosamente.")
                st.download_button(
                    "📄 Descargar Reporte Diario de Dotación",[cite: 2]
                    pdf_bytes,[cite: 2]
                    f"Reporte_Diario_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf",[cite: 2]
                    "application/pdf"[cite: 2]
                )

        except Exception as e:
            st.error(f"Error al procesar los archivos: {e}")
else:
    st.info("💡 Asegurate de que los archivos 'query-hoy' y 'query-ant' existan en tu Escritorio.")
