import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="Outlook → Contactos (Excel multi-hoja + País por TLD)", layout="wide")
st.title("📤 Outlook → 🧰 Contactos Limpios (Excel multi-hoja + País por dominio)")

st.markdown("""
Subí un **CSV** exportado de Outlook. La app:
- Busca correos en **todas** las columnas.
- Deduplica por email y conserva la **fecha más reciente** (si hay fecha).
- Infere **Nombre/Apellido** y **Empresa** desde el dominio.
- **Pais** por TLD del dominio (ej.: `.cr`, `.mx`, `.com.mx`, `.hn`). Si no se puede inferir, queda **vacío**.
- Clasifica **Cliente reciente** (últimos *N* meses) vs **Seguimiento**.
- **Excluye** correos no deseados según reglas (ventas@, info@, etc.).
- Descarga **un solo Excel** con 3 hojas: **Contactos**, **Empresas**, **Excluidos**.
""")

# ===== Utilidades =====

PERSONAL_DOMAINS = {
    "gmail.com","hotmail.com","outlook.com","yahoo.com","icloud.com","proton.me","live.com","msn.com"
}

COMPANY_SUFFIXES = [
    "consulting","consultores","consultora","solutions","soluciones","group","grupo",
    "corp","corporation","company","compania","compañia","co","ltda","ltd","llc",
    "srl","sa","saa","sac","inc","ag","gmbh","bv","plc","pty","sas"
]

DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %H:%M",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %I:%M:%S %p",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
]

EMAIL_REGEX = re.compile(r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})")
DATE_CANDIDATES = ["Sent","Enviado","Fecha","Date","Fecha de envío","Sent On","Date Sent","Enviados el","Fecha de envío:" ]

DEFAULT_ROLE_PREFIXES = [
    "ventas","sales","info","contact","admin","hr","hello","support",
    "marketing","billing","accounts","compras","noreply","no-reply","noresponder","no-responder"
]

# Mapa simple de TLDs a país (agregá los que necesites)
COUNTRY_MAP = {
    "ar":"Argentina", "bo":"Bolivia", "br":"Brasil", "cl":"Chile", "co":"Colombia",
    "cr":"Costa Rica", "do":"República Dominicana", "ec":"Ecuador", "es":"España",
    "gt":"Guatemala", "hn":"Honduras", "mx":"México", "ni":"Nicaragua", "pa":"Panamá",
    "pe":"Perú", "py":"Paraguay", "sv":"El Salvador", "uy":"Uruguay", "ve":"Venezuela",
    "pr":"Puerto Rico"
}

def parse_date(s):
    if pd.isna(s):
        return None
    s = str(s).strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def split_domain(domain):
    parts = domain.lower().split(".")
    if len(parts) >= 2:
        return parts[-2], parts[-1]
    return parts[0], ""

def prettify_company_from_domain(domain):
    d = domain.lower()
    if d in PERSONAL_DOMAINS:
        return "Particular"
    sld, _ = split_domain(d)
    for suf in COMPANY_SUFFIXES:
        if sld.endswith(suf):
            core = sld[:-len(suf)]
            if core and core[-1].isalpha():
                company = f"{core.upper() if len(core) <= 3 else core.capitalize()} {suf.capitalize()}"
            else:
                company = suf.capitalize()
            return company.strip()
    return sld.upper() if len(sld) <= 3 else sld.capitalize()

def infer_name_parts(local_part):
    lp = local_part.lower().replace("-", ".").replace("_", ".")
    tokens = [t for t in lp.split(".") if t and not t.isdigit()]
    nombre = tokens[0].capitalize() if tokens else ""
    apellido = tokens[-1].capitalize() if len(tokens) >= 2 else ""
    return nombre, apellido

def harvest_emails_from_row(row: pd.Series):
    emails = set(); cols = set()
    for col, val in row.items():
        if pd.isna(val): 
            continue
        text = str(val)
        for m in EMAIL_REGEX.finditer(text):
            em = m.group(1).lower()
            if "@" in em:
                emails.add(em); cols.add(col)
    return emails, cols

def infer_country_from_domain(domain: str) -> str:
    """
    Intenta inferir el país desde el TLD:
    - TLD final de 2 letras (p.ej. .cr, .mx, .hn)
    - o TLD compuesto: *.com.mx → usa 'mx'
    Si no se puede inferir, devuelve '' (vacío).
    """
    if not domain:
        return ""
    labels = domain.lower().split(".")
    if len(labels) < 2:
        return ""
    last = labels[-1]
    # Caso 1: el último label es un ccTLD (2 letras)
    if len(last) == 2 and last.isalpha():
        return COUNTRY_MAP.get(last, "")
    # Caso 2: dominios como *.com.mx → usar el penúltimo si es de 2 letras
    if len(labels) >= 2:
        penultimo = labels[-2]
        if len(penultimo) == 2 and penultimo.isalpha():
            return COUNTRY_MAP.get(penultimo, "")
    return ""

def make_excel_bytes(df_contacts: pd.DataFrame, df_companies: pd.DataFrame, df_excluded: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Contactos
        if not df_contacts.empty:
            cols = ["Email","Nombre","Apellido","Dominio","Empresa","Pais","UltimoEnvio","EstadoCliente","AsuntoUltimo","ColumnasOrigen"]
            cols = [c for c in cols if c in df_contacts.columns] + [c for c in df_contacts.columns if c not in cols]
            df_contacts[cols].to_excel(writer, sheet_name="Contactos", index=False)
        else:
            pd.DataFrame(columns=["Email","Nombre","Apellido","Dominio","Empresa","Pais","UltimoEnvio","EstadoCliente","AsuntoUltimo","ColumnasOrigen"]).to_excel(writer, sheet_name="Contactos", index=False)
        # Empresas
        if not df_companies.empty:
            df_companies.to_excel(writer, sheet_name="Empresas", index=False)
        else:
            pd.DataFrame(columns=["Empresa","Dominio","Pais","ContactosUnicos","TotalEmails","UltimoEnvio"]).to_excel(writer, sheet_name="Empresas", index=False)
        # Excluidos
        if not df_excluded.empty:
            df_excluded.to_excel(writer, sheet_name="Excluidos", index=False)
        else:
            pd.DataFrame(columns=["Email","Motivo","FilaOrigen","ColumnasOrigen"]).to_excel(writer, sheet_name="Excluidos", index=False)
    buffer.seek(0)
    return buffer.getvalue()

# ===== Sidebar: parámetros =====
uploaded = st.file_uploader("Subí tu CSV exportado de Outlook", type=["csv"])
months = st.sidebar.slider("Meses para 'Cliente reciente'", min_value=1, max_value=18, value=6, step=1)
st.sidebar.write("Ventana para clasificar: últimos", months, "meses")

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros de exclusión")

use_role = st.sidebar.toggle("Excluir cuentas genéricas (role-based)", value=True,
                             help="ventas@, info@, hr@, support@, marketing@, etc.")
custom_list = st.sidebar.text_area(
    "Prefijos de email a excluir (uno por línea)", 
    value="",
    placeholder="ej.:\nventas\ninfo\nnoreply\nno-responder\natencion"
)
use_regex = st.sidebar.toggle("Tratar la lista anterior como expresiones REGEX", value=False,
                              help="Si está activado, cada línea se interpreta como patrón regex (aplica sobre el local-part antes de @).")

if uploaded is None:
    st.info("Esperando archivo CSV…")
    st.stop()

# Leer CSV
df = pd.read_csv(uploaded, dtype=str, encoding="utf-8", keep_default_na=False)
st.success(f"Archivo cargado: {uploaded.name} — {df.shape[0]} filas, {df.shape[1]} columnas")
st.dataframe(df.head(20), use_container_width=True)

# Detectar columna de fecha (opcional)
lower_cols = {c.lower(): c for c in df.columns}
auto_date = None
for cand in DATE_CANDIDATES:
    if cand.lower() in lower_cols:
        auto_date = lower_cols[cand.lower()]
        break

date_col = st.selectbox(
    "Selecciona la columna de FECHA (opcional, mejora la clasificación):",
    ["(ninguna)"] + list(df.columns),
    index=(0 if auto_date is None else (list(df.columns).index(auto_date)+1))
)
use_date = None if date_col == "(ninguna)" else date_col

# Armado de filtros
role_prefixes = set([p.strip().lower() for p in DEFAULT_ROLE_PREFIXES]) if use_role else set()
custom_prefixes = [line.strip().lower() for line in custom_list.splitlines() if line.strip()]

def is_excluded_local(local_part: str) -> bool:
    lp = local_part.lower()
    # role-based
    if role_prefixes and any(lp == pref or lp.startswith(pref + "+") or lp.startswith(pref + ".") for pref in role_prefixes):
        return True
    # custom
    if custom_prefixes:
        if use_regex:
            try:
                return any(re.search(patt, lp) for patt in custom_prefixes)
            except re.error:
                return any(lp == pref or lp.startswith(pref + "+") or lp.startswith(pref + ".") for pref in custom_prefixes)
        else:
            return any(lp == pref or lp.startswith(pref + "+") or lp.startswith(pref + ".") for pref in custom_prefixes)
    return False

# Procesar
with st.spinner("Procesando…"):
    records = {}
    cols_by_email = defaultdict(set)
    excluded_rows = []

    for idx, row in df.iterrows():
        sent_dt = parse_date(row[use_date]) if use_date else None
        subject = row.get("Subject", "") or row.get("Asunto", "") or ""
        found, cols = harvest_emails_from_row(row)
        for em in found:
            try:
                local, domain = em.split("@", 1)
            except ValueError:
                continue

            # Excluir por filtros
            if is_excluded_local(local):
                excluded_rows.append({
                    "Email": em,
                    "Motivo": "Filtro de exclusión",
                    "FilaOrigen": idx + 1,
                    "ColumnasOrigen": ";".join(sorted(cols))
                })
                continue

            nombre, apellido = infer_name_parts(local)
            empresa = prettify_company_from_domain(domain)
            pais = infer_country_from_domain(domain)  # <= NUEVO: país desde TLD
            prev = records.get(em)
            if prev is None:
                records[em] = {
                    "Email": em,
                    "Nombre": nombre,
                    "Apellido": apellido,
                    "Dominio": domain.lower(),
                    "Empresa": empresa,
                    "Pais": pais,               # <= NUEVO
                    "UltimoEnvio": sent_dt,
                    "AsuntoUltimo": subject
                }
                cols_by_email[em] |= set(cols)
            else:
                if sent_dt and (prev["UltimoEnvio"] is None or sent_dt > prev["UltimoEnvio"]):
                    prev["UltimoEnvio"] = sent_dt
                    prev["AsuntoUltimo"] = subject
                if not prev["Nombre"] and nombre:
                    prev["Nombre"] = nombre
                if not prev["Apellido"] and apellido:
                    prev["Apellido"] = apellido
                if not prev.get("Pais") and pais:       # completa País si estaba vacío
                    prev["Pais"] = pais
                cols_by_email[em] |= set(cols)

    # Contacts DF
    cutoff = datetime.now() - timedelta(days=30*months)
    contacts = []
    for em, data in records.items():
        last_str = data["UltimoEnvio"].strftime("%Y-%m-%d %H:%M:%S") if data["UltimoEnvio"] else ""
        estado = "Cliente para seguimiento"
        if data["UltimoEnvio"] and data["UltimoEnvio"] >= cutoff:
            estado = "Cliente reciente"
        contacts.append({
            **data,
            "UltimoEnvio": last_str,
            "EstadoCliente": estado,
            "ColumnasOrigen": ";".join(sorted(cols_by_email[em]))
        })
    df_contacts = pd.DataFrame(contacts)

    # Excluidos DF
    df_excluded = pd.DataFrame(excluded_rows)

    # Company rollup (solo contactos válidos) — incluimos País si hay mayoría, si no queda vacío
    agg = defaultdict(lambda: {"Dominio":"","Empresa":"", "Pais":"", "ContactosUnicos":set(),"TotalEmails":0,"UltimoEnvio":None})
    for row in contacts:
        key = (row["Empresa"], row["Dominio"])
        d = agg[key]
        d["Dominio"] = row["Dominio"]
        d["Empresa"] = row["Empresa"]
        # Para País: el primero que aparezca (en dominios .com se quedará vacío, como pediste)
        if not d["Pais"] and row.get("Pais"):
            d["Pais"] = row["Pais"]
        d["TotalEmails"] += 1
        d["ContactosUnicos"].add(row["Email"])
        if row["UltimoEnvio"]:
            dt = parse_date(row["UltimoEnvio"])
            if dt and (d["UltimoEnvio"] is None or dt > d["UltimoEnvio"]):
                d["UltimoEnvio"] = dt

    rows_company = []
    for (empresa, dominio), d in agg.items():
        last_str = d["UltimoEnvio"].strftime("%Y-%m-%d %H:%M:%S") if d["UltimoEnvio"] else ""
        rows_company.append({
            "Empresa": empresa, "Dominio": dominio, "Pais": d["Pais"],
            "ContactosUnicos": len(d["ContactosUnicos"]), "TotalEmails": d["TotalEmails"],
            "UltimoEnvio": last_str
        })
    df_companies = pd.DataFrame(rows_company).sort_values(["Empresa","Dominio"])

# KPIs
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Contactos únicos", len(df_contacts))
col2.metric("Recientes", int((df_contacts["EstadoCliente"]=="Cliente reciente").sum()))
col3.metric("Seguimiento", int((df_contacts["EstadoCliente"]=="Cliente para seguimiento").sum()))
col4.metric("Empresas", df_companies["Empresa"].nunique() if not df_companies.empty else 0)
col5.metric("Excluidos", len(df_excluded))

# Vistas
tab1, tab2, tab3 = st.tabs(["✅ Contactos", "🏢 Empresas", "🚫 Excluidos"])
with tab1:
    st.dataframe(df_contacts, use_container_width=True)
with tab2:
    st.dataframe(df_companies, use_container_width=True)
with tab3:
    st.dataframe(df_excluded if not df_excluded.empty else pd.DataFrame(columns=["Email","Motivo","FilaOrigen","ColumnasOrigen"]), use_container_width=True)

# Descargar EXCEL multi-hoja
def make_excel_bytes(df_contacts: pd.DataFrame, df_companies: pd.DataFrame, df_excluded: pd.DataFrame) -> bytes:
    from io import BytesIO
    buffer = BytesIO()
    # 👇 cambiá el engine a xlsxwriter
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        # Contactos
        if not df_contacts.empty:
            cols = ["Email","Nombre","Apellido","Dominio","Empresa","Pais","UltimoEnvio","EstadoCliente","AsuntoUltimo","ColumnasOrigen"]
            cols = [c for c in cols if c in df_contacts.columns] + [c for c in df_contacts.columns if c not in cols]
            df_contacts[cols].to_excel(writer, sheet_name="Contactos", index=False)
        else:
            pd.DataFrame(columns=["Email","Nombre","Apellido","Dominio","Empresa","Pais","UltimoEnvio","EstadoCliente","AsuntoUltimo","ColumnasOrigen"]).to_excel(writer, sheet_name="Contactos", index=False)

        # Empresas
        if not df_companies.empty:
            df_companies.to_excel(writer, sheet_name="Empresas", index=False)
        else:
            pd.DataFrame(columns=["Empresa","Dominio","Pais","ContactosUnicos","TotalEmails","UltimoEnvio"]).to_excel(writer, sheet_name="Empresas", index=False)

        # Excluidos
        if not df_excluded.empty:
            df_excluded.to_excel(writer, sheet_name="Excluidos", index=False)
        else:
            pd.DataFrame(columns=["Email","Motivo","FilaOrigen","ColumnasOrigen"]).to_excel(writer, sheet_name="Excluidos", index=False)

    buffer.seek(0)
    return buffer.getvalue()
