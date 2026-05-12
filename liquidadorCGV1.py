import streamlit as st
import pandas as pd
import io
from datetime import datetime
import calendar
import re

# --- CONFIGURACIÓN DE PÁGINA Y TÍTULO ---
st.set_page_config(page_title="Liquidador CG", page_icon="💰", layout="wide")

def check_password():
    """Devuelve True si el usuario ingresó la contraseña correcta."""
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Introduzca la contraseña para acceder", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Contraseña incorrecta. Reintente:", type="password", on_change=password_entered, key="password")
        return False
    else:
        return True

if not check_password():
    st.stop()

# --- INTERFAZ DE USUARIO ---
st.title("Liquidador de Haberes CG - Claudio Coronel V.1")

with st.sidebar:
    st.header("Parámetros de Liquidación")
    mes_liq = st.number_input("Mes", min_value=1, max_value=12, value=datetime.now().month)
    anio_liq = st.number_input("Año", min_value=2024, max_value=2030, value=datetime.now().year)
    ultimo_dia = calendar.monthrange(anio_liq, mes_liq)[1]
    fecha_corte = datetime(anio_liq, mes_liq, ultimo_dia)
    
    st.divider()
    st.info("Suba el archivo Excel con las pestañas: MAESTRO_AGENTES, CATEGORIAS_AGENTES, REGLAS_LIQUIDACION, MATRIZ_VALORES, MATRIZ_FUNCIONES y MATRIZ_PARAMETROS.")

archivo_subido = st.file_uploader("Seleccione la Base de Liquidación (Excel)", type=["xlsx"])

# --- FUNCIONES NÚCLEO (Preservando tu lógica) ---
def safe_float(val):
    try: return float(val) if pd.notna(val) and str(val).strip() != "" else 0.0
    except: return 0.0

def buscar_con_jerarquia(df, condicion_base, id_escalafon):
    # Filtro por Organismo
    filtro_esc = (df['ID_ESCALAFON'].isna()) | (df['ID_ESCALAFON'].astype(str).str.strip() == "") | (df['ID_ESCALAFON'] == id_escalafon)
    resultado = df[condicion_base & filtro_esc]
    if not resultado.empty:
        return resultado.iloc[-1] # Prioriza la fila con ID_ESCALAFON si existe
    return pd.Series(dtype='float64')

# --- PROCESO DE LIQUIDACIÓN ---
if archivo_subido:
    try:
        with st.spinner("Procesando liquidación..."):
            hojas = pd.read_excel(archivo_subido, sheet_name=None, engine='openpyxl')
            
            # Carga de hojas
            df_agentes = hojas["MAESTRO_AGENTES"]
            df_cat_agentes = hojas["CATEGORIAS_AGENTES"]
            df_reglas = hojas["REGLAS_LIQUIDACION"]
            
            resultados_est = []
            resultados_aud = []

            for _, agente in df_agentes.iterrows():
                dni = agente['DNI']
                id_escalafon = str(agente.get('ID_ESCALAFON', '')).strip().upper()
                
                # Búsqueda individualizada (DNI + ID_ESCALAFON para Adscriptos)
                cat_info_base = df_cat_agentes[df_cat_agentes['DOCUMENTO'] == dni]
                cat_info_exacta = cat_info_base[cat_info_base['ID_ESCALAFON'].astype(str).str.strip().str.upper() == id_escalafon]
                cat_info = cat_info_exacta if not cat_info_exacta.empty else cat_info_base
                
                if cat_info.empty: continue
                
                info = cat_info.iloc[0]
                categoria_norm = str(info.get('CATEGORIA', '')).strip()
                desc_clase = str(info.get('DESC_CLASE', '')).strip().upper()
                
                # Contexto de Liquidación
                # El blindaje: inicializamos todos los códigos en 0.0
                ctx = {"DNI": dni, "ID_ESCALAFON": id_escalafon, "CATEGORIA": categoria_norm}
                for cod_r in df_reglas['CODIGO'].unique():
                    ctx[str(cod_r).strip().upper()] = 0.0

                # Definición de funciones internas para las fórmulas
                def CAT(codigo):
                    df = hojas['MATRIZ_VALORES']
                    fila = df[df['CATEGORIA'].astype(str).str.strip() == categoria_norm]
                    if not fila.empty: return safe_float(fila.iloc[0].get(codigo, 0))
                    return 0.0

                def CAT_REF(cat, codigo):
                    df = hojas['MATRIZ_VALORES']
                    fila = df[df['CATEGORIA'].astype(str).str.strip() == str(cat)]
                    if not fila.empty: return safe_float(fila.iloc[0].get(codigo, 0))
                    return 0.0

                def FUNC(col):
                    func_agente = str(agente.get('FUNCION', '')).strip().upper()
                    if func_agente in ["", "NAN"]: func_agente = "SIN FUNCION"
                    
                    df = hojas['MATRIZ_FUNCIONES']
                    condicion = (df['FUNCION'].astype(str).str.upper() == func_agente)
                    res = buscar_con_jerarquia(df, condicion, id_escalafon)
                    valor = safe_float(res.get(col, 0))
                    
                    # Cascada
                    if valor == 0.0 and func_agente == "SIN FUNCION" and desc_clase != "":
                        cond_clase = (df['FUNCION'].astype(str).str.upper() == desc_clase)
                        res_clase = buscar_con_jerarquia(df, cond_clase, id_escalafon)
                        valor = safe_float(res_clase.get(col, 0))
                    return valor

                def AGENTE(col): return agente.get(col, 0)

                # Ejecución de reglas
                df_reglas_ord = df_reglas.sort_values("ORDEN")
                for _, regla in df_reglas_ord.iterrows():
                    cod = str(regla['CODIGO']).strip().upper()
                    condicion = str(regla.get('CONDICION', 'True')).strip()
                    if condicion == "" or pd.isna(regla['CONDICION']): condicion = "True"
                    
                    try:
                        if eval(condicion, {"AGENTE": AGENTE, "CATEGORIA": safe_float(categoria_norm)}, ctx):
                            formula = str(regla['FORMULA'])
                            # El corazón matemático
                            ctx[cod] = eval(formula, {"CAT": CAT, "CAT_REF": CAT_REF, "FUNC": FUNC, "AGENTE": AGENTE, "min": min, "max": max}, ctx)
                    except:
                        ctx[cod] = 0.0
                
                resultados_est.append(ctx.copy())

            # Generación de Excel de salida
            df_final = pd.DataFrame(resultados_est)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='LIQUIDACION')
            
            st.success("✅ Procesamiento completado.")
            st.download_button(
                label="📥 Descargar Resultado",
                data=output.getvalue(),
                file_name=f"LIQUIDACION_{mes_liq}_{anio_liq}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")