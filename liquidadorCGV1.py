import streamlit as st
import pandas as pd
import io
from datetime import datetime
import calendar
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Liquidador CG", page_icon="💰", layout="wide")

def check_password():
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

# --- FUNCIONES DE AYUDA ORIGINALES ---
def parse_date(fecha_str):
    if pd.isna(fecha_str) or str(fecha_str).strip() in ["", "NaT"]: return None
    try: return fecha_str if isinstance(fecha_str, datetime) else pd.to_datetime(fecha_str)
    except: return None

def safe_float(val):
    try: return float(val) if pd.notna(val) and str(val).strip() != "" else 0.0
    except: return 0.0

def calcular_anios_antiguedad(f_ingreso, f_corte):
    if not f_ingreso: return 0
    anios = f_corte.year - f_ingreso.year
    if (f_corte.month, f_corte.day) < (f_ingreso.month, f_ingreso.day): anios -= 1
    return max(0, anios)

def calcular_diferencia_decimal(f_antigua, f_nueva):
    if not f_antigua or not f_nueva: return 0.0
    return round((f_nueva - f_antigua).days / 365.25, 2)

def normalizar_categoria(texto):
    texto = str(texto).strip().upper()
    if texto in ["", "NAN", "NONE"]: return None
    match = re.search(r'\d+', texto)
    return int(match.group()) if match else texto

# --- GENERADOR DEL CONTEXTO (Con protección anti-KeyError) ---
def crear_contexto(agente, categoria_norm, desc_clase, anios_antig, hojas, recibo_actual, id_escalafon, id_general, df_reglas_agente):
    jerarquia_esc = [str(id_escalafon).strip().upper()]
    if id_general == 'SI': jerarquia_esc.append('EG')
    jerarquia_esc.append('')

    def buscar_con_jerarquia(df, condicion_extra):
        # Normalizamos nombres de columnas para evitar espacios fantasmas en las matrices
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        for esc_buscado in jerarquia_esc:
            if 'ID_ESCALAFON' in df.columns:
                filtro_esc = df['ID_ESCALAFON'].fillna('').astype(str).str.strip().str.upper() == esc_buscado
            else:
                # Si la matriz de Excel no posee la columna, matchea únicamente el caso base vacío
                filtro_esc = pd.Series([True if esc_buscado == '' else False] * len(df), index=df.index)
                
            resultado = df[filtro_esc & condicion_extra]
            if not resultado.empty: return resultado.iloc[-1]
        return pd.Series(dtype='float64')

    def AGENTE(col): return safe_float(agente.get(col, 0))

    def CAT_REF(cat_id, concepto):
        cat_clean = str(cat_id).strip().upper()
        if cat_clean.endswith('.0'): cat_clean = cat_clean[:-2]
        df = hojas['MATRIZ_VALORES_FIJOS']
        df.columns = [str(c).strip().upper() for c in df.columns]
        condicion = (df['CATEGORIA'].astype(str).str.upper() == cat_clean) & (df['CONCEPTO'].astype(str).str.upper() == str(concepto).upper())
        res = buscar_con_jerarquia(df, condicion)
        return safe_float(res.get('VALORES', 0))

    def CAT(concepto): return CAT_REF(categoria_norm, concepto)

    def FUNC(col):
        funcion = str(agente.get('FUNCION', '')).strip().upper()
        df = hojas['MATRIZ_FUNCIONES']
        df.columns = [str(c).strip().upper() for c in df.columns]
        condicion = (df['FUNCION'].astype(str).str.upper() == funcion)
        res = buscar_con_jerarquia(df, condicion)
        valor = safe_float(res.get(col.strip().upper(), 0))
        
        if valor == 0.0 and funcion in ["", "NAN"]:
            cond_sin = (df['FUNCION'].astype(str).str.upper() == "SIN FUNCION")
            res_sin = buscar_con_jerarquia(df, cond_sin)
            valor = safe_float(res_sin.get(col.strip().upper(), 0))
            if valor == 0.0 and desc_clase != "":
                cond_clase = (df['FUNCION'].astype(str).str.upper() == desc_clase)
                res_clase = buscar_con_jerarquia(df, cond_clase)
                valor = safe_float(res_clase.get(col.strip().upper(), 0))
        return valor
        
    def VALOR_CATEGORIA(col):
        cat_str = str(categoria_norm).strip().upper()
        df = hojas['MATRIZ_FUNCIONES']
        df.columns = [str(c).strip().upper() for c in df.columns]
        condicion = (df['FUNCION'].astype(str).str.upper() == cat_str)
        res = buscar_con_jerarquia(df, condicion)
        return safe_float(res.get(col.strip().upper(), 0))

    def TITULO(col):
        titulo = str(agente.get('TITULOS', '')).strip().upper()
        if desc_clase == "PROFESIONAL UNIVERSITARIO": titulo = 'PROFESIONAL UNIVERSITARIO'
        df = hojas['MATRIZ_TITULOS']
        df.columns = [str(c).strip().upper() for c in df.columns]
        condicion = (df['TITULO_NOMBRE'].astype(str).str.upper() == titulo)
        res = buscar_con_jerarquia(df, condicion)
        return safe_float(res.get(col.strip().upper(), 0))

    def PCT_ANTIGUEDAD():
        df = hojas['MATRIZ_PARAMETROS']
        df.columns = [str(c).strip().upper() for c in df.columns]
        condicion = (df['TIPO_PARAMETRO'].astype(str).str.upper() == 'ANTIGUEDAD') & (df['CLAVE_MIN'] <= anios_antig) & (df['CLAVE_MAX'] >= anios_antig)
        res = buscar_con_jerarquia(df, condicion)
        return safe_float(res.get('VALOR_RESULTADO', 0))

    def MATRIZ(hoja, col_busqueda, valor, col_resultado):
        df = hojas[hoja]
        df.columns = [str(c).strip().upper() for c in df.columns]
        filtro = df[col_busqueda.strip().upper()].astype(str).str.contains(str(valor), case=False, na=False)
        return safe_float(df[filtro].iloc[0].get(col_resultado.strip().upper(), 0)) if filtro.any() else 0.0

    def SUMAR_ATRIBUTO(nombre_columna):
        suma = 0.0
        # Normalizar columnas de la regla actual
        df_reglas_agente.columns = [str(c).strip().upper() for c in df_reglas_agente.columns]
        col_busqueda = nombre_columna.strip().upper()
        
        if col_busqueda not in df_reglas_agente.columns: return 0.0
        for cod_recibo, monto in recibo_actual.items():
            if isinstance(monto, (int, float)) and str(cod_recibo).startswith('COD_'):
                filtro = df_reglas_agente['CODIGO'].astype(str).str.strip().str.upper() == str(cod_recibo).upper()
                if filtro.any():
                    es_valido = str(df_reglas_agente[filtro].iloc[-1].get(col_busqueda, 'NO')).strip().upper()
                    if es_valido == 'SI': suma += monto
        return round(suma, 2)

    def SUMAR_BONIFICABLES(): return SUMAR_ATRIBUTO('BONIFICABLE')
    def SUMAR_REMUNERATIVOS(): return SUMAR_ATRIBUTO('REMUNERATIVO')

    es_prof = desc_clase == "PROFESIONAL UNIVERSITARIO"
    aplica_146 = isinstance(categoria_norm, int) or "PRESIDENTE" in str(categoria_norm).upper()

    ctx = recibo_actual.copy()
    ctx.update({
        'AGENTE': AGENTE, 'CAT': CAT, 'CAT_REF': CAT_REF, 'FUNC': FUNC, 
        'VALOR_CATEGORIA': VALOR_CATEGORIA, 'TITULO': TITULO, 'PCT_ANTIGUEDAD': PCT_ANTIGUEDAD, 
        'MATRIZ': MATRIZ, 'SUMAR_BONIFICABLES': SUMAR_BONIFICABLES, 'SUMAR_REMUNERATIVOS': SUMAR_REMUNERATIVOS,
        'ES_PROFESIONAL': es_prof, 'NO_ES_PROFESIONAL': not es_prof, 'APLICA_146': aplica_146,
        'min': min, 'max': max
    })
    return ctx

# --- INTERFAZ WEB ---
st.title("Liquidador de Haberes CG - Claudio Coronel V.1")

with st.sidebar:
    st.header("Parámetros de Liquidación")
    mes_liq = st.number_input("Mes", min_value=1, max_value=12, value=datetime.now().month)
    anio_liq = st.number_input("Año", min_value=2024, max_value=2030, value=datetime.now().year)
    ultimo_dia = calendar.monthrange(anio_liq, mes_liq)[1]
    fecha_corte = datetime(anio_liq, mes_liq, ultimo_dia)
    st.divider()

archivo_subido = st.file_uploader("📂 Seleccione la Base de Liquidación (Excel)", type=["xlsx"])

if archivo_subido:
    try:
        with st.spinner("⏳ Leyendo Excel y calculando matrices..."):
            hojas = pd.read_excel(archivo_subido, sheet_name=None, engine='openpyxl')
            
            # Normalizar los nombres de las columnas de las hojas principales de entrada
            for name in hojas:
                hojas[name].columns = [str(c).strip().upper() for c in hojas[name].columns]
            
            df_agentes = hojas["MAESTRO_AGENTES"]
            df_cat_agentes = hojas["CATEGORIAS_AGENTES"]
            df_reglas_global = hojas["REGLAS_LIQUIDACION"]
            
            recibos_est = []
            recibos_aud = []
            agentes_ignorados = [] 

            # BUCLE PRINCIPAL
            for _, agente in df_agentes.iterrows():
                dni = agente['DNI']
                id_escalafon = str(agente.get('ID_ESCALAFON', '')).strip().upper()
                id_general = str(agente.get('ID_GENERAL', 'NO')).strip().upper()
                
                # Lógica Adscriptos con control de existencia de columna
                cat_info_base = df_cat_agentes[df_cat_agentes['DOCUMENTO'] == dni]
                if cat_info_base.empty: 
                    agentes_ignorados.append({"DNI": dni, "MOTIVO": "DNI inexistente en CATEGORIAS_AGENTES"})
                    continue
                
                if 'ID_ESCALAFON' in cat_info_base.columns:
                    cat_info_exacta = cat_info_base[cat_info_base['ID_ESCALAFON'].astype(str).str.strip().str.upper() == id_escalafon]
                    cat_info = cat_info_exacta if not cat_info_exacta.empty else cat_info_base
                else:
                    cat_info = cat_info_base
                
                raw_escalafon = cat_info.iloc[0].get('ESCALAFON', '')
                es_funcionario_eg = False
                try:
                    if str(raw_escalafon).strip() != '' and float(raw_escalafon) == 0.0: es_funcionario_eg = True
                except: pass

                desc_cat_base = cat_info.iloc[0].get('DESCCATEGORIA', '')
                desc_cat_equip = cat_info.iloc[0].get('DESCCATEQUIP', '')
                desc_clase = str(cat_info.iloc[0].get('DESCCLASE', '')).strip().upper()
                
                categoria_texto = desc_cat_equip if pd.notna(desc_cat_equip) and str(desc_cat_equip).strip() not in ["", "NAN", "NAT"] else desc_cat_base
                categoria_norm = normalizar_categoria(categoria_texto)
                
                if categoria_norm is None: 
                    agentes_ignorados.append({"DNI": dni, "MOTIVO": "Categoría vacía o inválida"})
                    continue

                # Filtrado de Reglas
                reglas_validas = []
                for _, regla in df_reglas_global.iterrows():
                    regla_esc = str(regla.get('ID_ESCALAFON', '')).strip().upper()
                    if regla_esc == 'NAN': regla_esc = ''
                    es_regla_func = str(regla.get('FUNCIONARIO_EG', 'NO')).strip().upper() == 'SI'
                    
                    if es_funcionario_eg:
                        if es_regla_func: reglas_validas.append(regla)
                    else:
                        if not es_regla_func:
                            if regla_esc == id_escalafon or regla_esc == '': reglas_validas.append(regla)
                            elif regla_esc == 'EG' and id_general == 'SI': reglas_validas.append(regla)
                        
                df_reglas_agente = pd.DataFrame(reglas_validas)
                if df_reglas_agente.empty:
                    agentes_ignorados.append({"DNI": dni, "MOTIVO": "Sin reglas válidas"})
                    continue
                    
                df_reglas_agente = df_reglas_agente.sort_values(by="ORDEN")

                f_normal = parse_date(agente.get('AL ESTADO NO AL ORGANISMO'))
                f_base = parse_date(agente.get('INGRESO_BASE'))
                
                # --- LIQUIDACIÓN ESTÁNDAR ---
                anios_est = calcular_anios_antiguedad(f_normal, fecha_corte)
                recibo_est = {"DNI": dni, "FUNCION": str(agente.get('FUNCION','')), "CATEGORIA": categoria_norm, "ESCALAFON": id_escalafon, "ANIOS_ANTIGUEDAD": anios_est}
                for col in agente.index:
                    if str(col).startswith("COD_"): recibo_est[str(col).strip().upper()] = safe_float(agente.get(col, 0))
                
                # Blindaje Estándar
                for cod_r in df_reglas_agente['CODIGO']:
                    cod_limpio = str(cod_r).strip().upper()
                    if cod_limpio not in recibo_est: recibo_est[cod_limpio] = 0.0

                for _, regla in df_reglas_agente.iterrows():
                    cod = str(regla['CODIGO']).strip().upper()
                    cond = str(regla.get('CONDICION', ''))
                    form = str(regla['FORMULA'])
                    ctx = crear_contexto(agente, categoria_norm, desc_clase, anios_est, hojas, recibo_est, id_escalafon, id_general, df_reglas_agente)
                    try:
                        cumple = eval(cond, {}, ctx) if cond and cond.strip() not in ['nan', ''] else True
                        recibo_est[cod] = round(eval(form, {}, ctx), 2) if cumple and form and form.strip() != 'nan' else 0.0
                    except:
                        recibo_est[cod] = 0.0
                recibos_est.append(recibo_est)

                # --- LIQUIDACIÓN AUDITORÍA ---
                f_auditoria = f_normal
                uso_base, dif = "NO", 0.0
                if f_base and f_normal and (f_base < f_normal):
                    f_auditoria, uso_base = f_base, "SI"
                    dif = calcular_diferencia_decimal(f_base, f_normal)

                anios_aud = calcular_anios_antiguedad(f_auditoria, fecha_corte)
                recibo_aud = {"DNI": dni, "FUNCION": str(agente.get('FUNCION','')), "CATEGORIA": categoria_norm, "ESCALAFON": id_escalafon, "ANIOS_ANTIGUEDAD": anios_aud}
                for col in agente.index:
                    if str(col).startswith("COD_"): recibo_aud[str(col).strip().upper()] = safe_float(agente.get(col, 0))
                
                # Blindaje Auditoría
                for cod_r in df_reglas_agente['CODIGO']:
                    cod_limpio = str(cod_r).strip().upper()
                    if cod_limpio not in recibo_aud: recibo_aud[cod_limpio] = 0.0
                    
                for _, regla in df_reglas_agente.iterrows():
                    cod = str(regla['CODIGO']).strip().upper()
                    cond = str(regla.get('CONDICION', ''))
                    form = str(regla['FORMULA'])
                    ctx = crear_contexto(agente, categoria_norm, desc_clase, anios_aud, hojas, recibo_aud, id_escalafon, id_general, df_reglas_agente)
                    try:
                        cumple = eval(cond, {}, ctx) if cond and cond.strip() not in ['nan', ''] else True
                        recibo_aud[cod] = round(eval(form, {}, ctx), 2) if cumple and form and form.strip() != 'nan' else 0.0
                    except:
                        recibo_aud[cod] = 0.0
                    
                recibo_aud["INGRESO_BASE"] = uso_base
                recibo_aud["DIFERENCIA_ANIOS"] = dif
                recibos_aud.append(recibo_aud)

            # --- EXPORTACIÓN ---
            df_est = pd.DataFrame(recibos_est)
            df_aud = pd.DataFrame(recibos_aud)

            orden_columnas = [
                "DNI", "FUNCION", "CATEGORIA", "ESCALAFON", "ANIOS_ANTIGUEDAD",
                "COD_001", "COD_003", "COD_004", "COD_005", "COD_006", "COD_008", "COD_010", "COD_011", "COD_012", "COD_013",
                "COD_014", "COD_017", "COD_019", "COD_020", "COD_022", "COD_024", "COD_029", "COD_036", "COD_038", "COD_041",
                "COD_042", "COD_050", "COD_054", "COD_056", "COD_057", "COD_058", "COD_060", "COD_081", "COD_090", "COD_092",
                "COD_100", "COD_101", "COD_105", "COD_109", "COD_110", "COD_116", "COD_117", "COD_118", "COD_126", "COD_130",
                "COD_136", "COD_138", "COD_140", "COD_142", "COD_146", "COD_148", "COD_166", "COD_170", "COD_174", "COD_176",
                "COD_177", "COD_178", "COD_179", "COD_180", "COD_181", "COD_183", "COD_184", "COD_185", "COD_186", "COD_188",
                "COD_192", "COD_194", "COD_195", "COD_201", "COD_204", "COD_206", "COD_208", "COD_210", "COD_212", "COD_214",
                "COD_219", "COD_220", "COD_221", "COD_222", "COD_225", "COD_226", "COD_228", "COD_232", "COD_234", "COD_235",
                "COD_236", "COD_239", "COD_240", "COD_243", "COD_244", "COD_246", "COD_248", "COD_264", "COD_265", "COD_266",
                "COD_272", "COD_280", "COD_282", "COD_285", "COD_286", "COD_288", "BRUTO", "SUMA_BONIFIC", "PORC_ANTIG"
            ]

            if not df_est.empty:
                for col in orden_columnas:
                    if col not in df_est.columns: df_est[col] = 0.0
                df_est = df_est[[c for c in orden_columnas if c in df_est.columns]]
                
            if not df_aud.empty:
                orden_aud = orden_columnas + ["INGRESO_BASE", "DIFERENCIA_ANIOS"]
                for col in orden_aud:
                    if col not in df_aud.columns: df_aud[col] = 0.0
                df_aud = df_aud[[c for c in orden_aud if c in df_aud.columns]]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_est.empty: df_est.to_excel(writer, sheet_name='LIQ_ESTANDAR', index=False)
                if not df_aud.empty: df_aud.to_excel(writer, sheet_name='LIQ_AUDITORIA', index=False)
                if agentes_ignorados: pd.DataFrame(agentes_ignorados).to_excel(writer, sheet_name='AGENTES_IGNORADOS', index=False)

            st.success(f"✅ Éxito. Se liquidaron {len(recibos_est)} agentes.")
            if agentes_ignorados: st.warning(f"⚠️ Atención: Se ignoraron {len(agentes_ignorados)} agentes.")
            
            st.download_button(
                label="📥 Descargar Excel Resultante",
                data=output.getvalue(),
                file_name=f"LIQUIDACION_{mes_liq}_{anio_liq}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
