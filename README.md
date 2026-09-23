import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Validador NAT OP - Devoluções", layout="wide"
)
st.title("📦 Consulta Automática de NAT OP (Coluna C)")


@st.cache_data
def carregar_regras():
    df = pd.read_excel("DEV.xlsx", sheet_name="Natureza de op")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def mapear_cfop_entrada(cfop_saida):
    cfop = str(cfop_saida).strip()
    if cfop.startswith("5"):
        return (
            "1201" if cfop in ["5101", "5103", "5118", "5122"] else "1202"
        )
    elif cfop.startswith("6"):
        if cfop in ["6101", "6103", "6118", "6122"]:
            return "2201"
        elif cfop in ["6103", "6203"]:
            return "2203"
        elif cfop in ["6104", "6204"]:
            return "2204"
        else:
            return "2202"
    return "1201"


def obter_icms(icms_elem):
    if icms_elem is None:
        return "T"
    pRedBC = icms_elem.find("{*}pRedBC")
    if pRedBC is not None and float(pRedBC.text) > 0:
        return "R"
    cst_elem = icms_elem.find("{*}CST")
    cst = cst_elem.text if cst_elem is not None else ""
    if cst in ["00", "10"]:
        return "T"
    elif cst in ["20"]:
        return "R"
    elif cst in ["40", "41"]:
        return "I"
    elif cst in ["50", "51", "60", "90"]:
        return "O"
    return "T"


def obter_ipi(ipi_elem):
    if ipi_elem is None:
        return 1.0
    cst_elem = ipi_elem.find(".//{*}CST")
    cst = cst_elem.text if cst_elem is not None else ""
    if cst in ["00", "49", "50"]:
        return 0.0
    elif cst in ["03", "04", "05"]:
        return 3.0
    else:
        return 1.0


def obter_pis_cofins(pis_elem, cofins_elem):
    cst_pis = (
        pis_elem.find(".//{*}CST").text
        if pis_elem is not None and pis_elem.find(".//{*}CST") is not None
        else "01"
    )
    pPIS_elem = pis_elem.find(".//{*}pPIS") if pis_elem is not None else None
    pPIS = float(pPIS_elem.text) if pPIS_elem is not None else 0.0
    pCOFINS_elem = (
        cofins_elem.find(".//{*}pCOFINS") if cofins_elem is not None else None
    )
    pCOFINS = float(pCOFINS_elem.text) if pCOFINS_elem is not None else 0.0

    if cst_pis in ["01", "02", "50", "51", "52", "53", "54", "55", "56", "60"]:
        cst_pc = 54.0
    elif cst_pis in ["04", "06", "07", "08", "09", "73"]:
        cst_pc = 73.0
    else:
        cst_pc = 70.0

    if pPIS > 2.1 or pCOFINS > 10.0:
        aliq = "0,022 e 0,103"
    elif pPIS > 1.8 or pCOFINS > 8.0:
        aliq = "0,021 e 0,099"
    elif pPIS > 0 or pCOFINS > 0:
        aliq = "0,0165 e 0,076"
    else:
        aliq = "0 e 0"

    return cst_pc, aliq


def processar_xml(xml_file, df_regras):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    n_nf = root.find(".//{*}ide/{*}nNF")
    n_nf = n_nf.text if n_nf is not None else "N/A"
    resultados = []

    # Identifica coluna de alíquota do PIS/COFINS
    col_pct = [
        c for c in df_regras.columns if "PIS e COFINS" in c and "%" in c
    ][0]

    for det in root.findall(".//{*}det"):
        n_item = det.attrib.get("nItem", "1")
        cProd = det.find(".//{*}prod/{*}cProd").text
        xProd = det.find(".//{*}prod/{*}xProd").text
        cfop_xml = det.find(".//{*}prod/{*}CFOP").text

        cfop_ref = float(mapear_cfop_entrada(cfop_xml))
        icms_val = obter_icms(det.find(".//{*}imposto/{*}ICMS/*"))
        ipi_val = obter_ipi(det.find(".//{*}imposto/{*}IPI"))
        pis_cofins_cst, aliq_str = obter_pis_cofins(
            det.find(".//{*}imposto/{*}PIS"),
            det.find(".//{*}imposto/{*}COFINS"),
        )

        match = df_regras[
            (df_regras["CFOP"] == cfop_ref)
            & (df_regras["ICMS"].astype(str).str.strip() == icms_val)
            & (df_regras["IPI"] == ipi_val)
            & (df_regras["PIS COFINS"] == pis_cofins_cst)
            & (
                df_regras[col_pct]
                .astype(str)
                .str.strip()
                .str.contains(aliq_str)
            )
        ]

        nat_op_coluna_c = (
            match["NAT OP"].values[0] if not match.empty else "NÃO ENCONTRADO"
        )

        resultados.append({
            "NF Origem": n_nf,
            "Item": n_item,
            "Código Produto": cProd,
            "Descrição do Produto": xProd,
            "NAT OP (Coluna C)": nat_op_coluna_c,
            "CFOP Saída (XML)": cfop_xml,
            "ICMS": icms_val,
            "IPI": int(ipi_val),
            "CST PIS/COFINS": int(pis_cofins_cst),
            "Alíquota PIS/COFINS": aliq_str,
        })
    return resultados


df_regras = carregar_regras()
uploaded_files = st.file_uploader(
    "Importe os arquivos XML das Notas Fiscais",
    type=["xml"],
    accept_multiple_files=True,
)

if uploaded_files:
    dados = []
    for f in uploaded_files:
        dados.extend(processar_xml(f, df_regras))

    df_res = pd.DataFrame(dados)
    st.dataframe(df_res, use_container_width=True)

    df_res.to_excel("Resultado_NAT_OP.xlsx", index=False)
    with open("Resultado_NAT_OP.xlsx", "rb") as f_out:
        st.download_button(
            "📥 Baixar Relatório Excel",
            f_out,
            file_name="Resultado_NAT_OP_Coluna_C.xlsx",
        )# OPER
