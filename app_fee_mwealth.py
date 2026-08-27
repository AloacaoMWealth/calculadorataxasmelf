from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Iterable, Optional

import pandas as pd
import requests
import streamlit as st

# ============================================================
# M WEALTH - MOTOR DE FEE
# ============================================================
# Regras principais:
# 1) O PL do GRUPO FAMILIAR na data de referência define a faixa.
# 2) Cada conta usa a Tabela Fee cadastrada no Controle de Clientes.
# 3) Contas com PL diário: Fee dia = PL_BRL * taxa_aa / 252.
# 4) Contas mensais: Fee mês = PL_BRL * taxa_aa / 12.
# 5) Offshore (Charles Schwab / XP US) é convertido para BRL pela PTAX.
# ============================================================

OFFSHORE_BROKERS = {"CHARLES SCHWAB", "XP US"}
NO_FEE_LABELS = {"SEM COBRANÇA", "SEM COBRANCA", "NÃO COBRAR", "NAO COBRAR"}


def norm_text(v) -> str:
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", " ", str(v).strip())


def norm_account(v) -> str:
    """Normaliza contas removendo máscara e zeros à esquerda."""
    s = re.sub(r"\D", "", norm_text(v))
    if not s:
        return ""
    return str(int(s))


def excel_date(v) -> Optional[pd.Timestamp]:
    if pd.isna(v) or v == "":
        return None
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).normalize()
    try:
        # Data serial do Excel
        if isinstance(v, (int, float)) or (isinstance(v, str) and re.fullmatch(r"\d+(\.\d+)?", v)):
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(v), unit="D")
        return pd.to_datetime(v, dayfirst=True).normalize()
    except Exception:
        return None


def month_end_before(year: int, month: int) -> date:
    first = date(year, month, 1)
    return first - timedelta(days=1)


def month_bounds(year: int, month: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(year=year, month=month, day=1)
    end = start + pd.offsets.MonthEnd(1)
    return start, end


def previous_business_day(d: date) -> date:
    d = d - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_ptax_usd(d: date) -> tuple[float, date]:
    """Busca PTAX de venda no BCB. Se a data não tiver cotação, volta dias."""
    probe = d
    for _ in range(10):
        ds = probe.strftime("%m-%d-%Y")
        url = (
            "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
            "CotacaoDolarDia(dataCotacao=@dataCotacao)"
            f"?@dataCotacao='{ds}'&$top=100&$format=json"
        )
        try:
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            vals = r.json().get("value", [])
            if vals:
                # último boletim do dia; cotacaoVenda = BRL por USD
                return float(vals[-1]["cotacaoVenda"]), probe
        except Exception:
            pass
        probe = previous_business_day(probe + timedelta(days=1))
    raise RuntimeError(f"Não foi possível obter a PTAX para {d:%d/%m/%Y} nem dias anteriores.")


def ptax_with_manual(d: date, manual: dict[date, float]) -> tuple[float, date, str]:
    if d in manual:
        return float(manual[d]), d, "Manual"
    value, used_date = fetch_ptax_usd(d)
    return value, used_date, "BCB"


def read_control(file) -> tuple[pd.DataFrame, dict[str, list[tuple[float, Optional[float], float]]], list[str]]:
    ctrl = pd.read_excel(file, sheet_name="Controle Contas", dtype=object)
    ctrl.columns = [norm_text(c) for c in ctrl.columns]

    required = ["Corretora", "Grupo Familiar", "Cliente", "Conta", "Tabela Fee"]
    missing = [c for c in required if c not in ctrl.columns]
    if missing:
        raise ValueError(f"Controle de Clientes sem colunas obrigatórias: {missing}")

    ctrl["Corretora"] = ctrl["Corretora"].map(lambda x: norm_text(x).upper())
    ctrl["Grupo Familiar"] = ctrl["Grupo Familiar"].map(norm_text)
    ctrl["Cliente"] = ctrl["Cliente"].map(norm_text)
    ctrl["Conta_norm"] = ctrl["Conta"].map(norm_account)
    ctrl["Tabela Fee"] = ctrl["Tabela Fee"].map(norm_text)

    fee_raw = pd.read_excel(file, sheet_name="Tabelas Fee", header=None, dtype=object)
    fee_tables = parse_fee_tables(fee_raw)
    pl_cols = [c for c in ctrl.columns if re.fullmatch(r"PL \d{2}/\d{2}/\d{4}", str(c))]
    return ctrl, fee_tables, pl_cols


def parse_fee_tables(raw: pd.DataFrame) -> dict[str, list[tuple[float, Optional[float], float]]]:
    """Lê blocos do tipo [Taxa | Início | Término] da aba Tabelas Fee."""
    tables: dict[str, list[tuple[float, Optional[float], float]]] = {}
    for col in range(raw.shape[1]):
        title = norm_text(raw.iat[0, col]) if raw.shape[0] > 0 else ""
        if not title:
            continue
        # Só blocos de 3 colunas com cabeçalho Segmentos / Início / Término
        if raw.shape[0] > 2 and norm_text(raw.iat[1, col]).lower() == "segmentos":
            rows = []
            for r in range(2, raw.shape[0]):
                rate = raw.iat[r, col]
                if pd.isna(rate) or rate == "":
                    continue
                try:
                    rate = float(rate)
                except Exception:
                    continue
                begin = raw.iat[r, col + 1] if col + 1 < raw.shape[1] else None
                end = raw.iat[r, col + 2] if col + 2 < raw.shape[1] else None
                try:
                    begin = float(begin) if not pd.isna(begin) and begin != "" else 0.0
                except Exception:
                    begin = 0.0
                try:
                    end = float(end) if not pd.isna(end) and end != "" else None
                except Exception:
                    end = None
                # A primeira faixa é tratada como aberta para baixo.
                # Nas planilhas históricas, Tabela 1/2 começam visualmente em R$ 1MM,
                # mas a lógica usada no legado cobra a primeira taxa também abaixo disso.
                if not rows:
                    begin = 0.0
                rows.append((begin, end, rate))
            if rows:
                tables[title.strip()] = rows
                tables[title.strip().upper()] = rows
    return tables


def choose_reference_pl_column(pl_cols: list[str], year: int, month: int) -> tuple[str, date]:
    target = month_end_before(year, month)
    candidates = []
    for c in pl_cols:
        try:
            d = datetime.strptime(c[3:], "%d/%m/%Y").date()
            if d <= target:
                candidates.append((d, c))
        except Exception:
            continue
    if not candidates:
        raise ValueError("Não encontrei coluna de PL anterior ao mês selecionado no Controle de Clientes.")
    return max(candidates)[1], max(candidates)[0]


def fixed_rate_from_label(label: str) -> Optional[float]:
    s = norm_text(label).upper()
    if s in NO_FEE_LABELS or "SEM COBR" in s:
        return 0.0
    # Exemplos: Tesouraria (0,35), Tesouraria (0,40%)
    if "TESOURARIA" in s:
        m = re.search(r"(\d+[,.]\d+)\s*%?", s)
        if m:
            pct = float(m.group(1).replace(",", "."))
            return pct / 100.0
    return None


def lookup_fee_rate(table_label: str, group_pl_brl: float, fee_tables) -> tuple[Optional[float], str]:
    fixed = fixed_rate_from_label(table_label)
    if fixed is not None:
        return fixed, "Taxa fixa/sem cobrança"

    candidates = [table_label, table_label.upper(), table_label.strip(), table_label.strip().upper()]
    rows = None
    for c in candidates:
        if c in fee_tables:
            rows = fee_tables[c]
            break
    if rows is None:
        return None, "Tabela não encontrada"

    for begin, end, rate in rows:
        if group_pl_brl >= begin and (end is None or group_pl_brl <= end):
            return rate, "Tabela regressiva"
    return None, "PL fora das faixas"


def build_fee_registry(ctrl: pd.DataFrame, fee_tables, ref_col: str, ref_date: date,
                       manual_ptax: dict[date, float]) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = ctrl.copy()
    df["PL_ref_original"] = pd.to_numeric(df[ref_col], errors="coerce").fillna(0.0)

    ref_ptax, ref_ptax_date, ref_ptax_src = ptax_with_manual(ref_date, manual_ptax)
    df["Moeda_ref"] = df["Corretora"].apply(lambda x: "USD" if x in OFFSHORE_BROKERS else "BRL")
    df["PTAX_ref"] = df["Moeda_ref"].apply(lambda x: ref_ptax if x == "USD" else 1.0)
    df["PL_ref_BRL"] = df["PL_ref_original"] * df["PTAX_ref"]

    # Grupo Familiar vazio cai no Cliente para não misturar vazios.
    df["Grupo_chave"] = df["Grupo Familiar"].where(df["Grupo Familiar"].ne(""), df["Cliente"])
    group = df.groupby("Grupo_chave", dropna=False)["PL_ref_BRL"].sum().rename("PL_Grupo_Ref_BRL")
    df = df.merge(group, left_on="Grupo_chave", right_index=True, how="left")
    df["Participacao_no_Grupo"] = df["PL_ref_BRL"] / df["PL_Grupo_Ref_BRL"].replace(0, pd.NA)

    rates, rules = [], []
    for _, r in df.iterrows():
        rate, rule = lookup_fee_rate(r["Tabela Fee"], float(r["PL_Grupo_Ref_BRL"] or 0), fee_tables)
        rates.append(rate)
        rules.append(rule)
    df["Fee_aa"] = rates
    df["Regra_Fee"] = rules
    df["Fee_Projetado_Anual_Ref"] = df["PL_ref_BRL"] * df["Fee_aa"].fillna(0)
    df["Fee_Projetado_Mensal_Ref"] = df["Fee_Projetado_Anual_Ref"] / 12

    group_view = (
        df.groupby("Grupo_chave", as_index=False)
          .agg(PL_Grupo_Ref_BRL=("PL_ref_BRL", "sum"),
               Contas=("Conta_norm", "count"),
               Fee_Projetado_Anual=("Fee_Projetado_Anual_Ref", "sum"))
    )
    group_view["ROA_Contratado"] = group_view["Fee_Projetado_Anual"] / group_view["PL_Grupo_Ref_BRL"].replace(0, pd.NA)
    group_view["Data_Referencia"] = pd.Timestamp(ref_date)
    group_view["PTAX_Referencia"] = ref_ptax
    group_view["Data_PTAX_Utilizada"] = pd.Timestamp(ref_ptax_date)
    group_view["Fonte_PTAX"] = ref_ptax_src

    keep = [
        "Grupo_chave", "Grupo Familiar", "Cliente", "Corretora", "Conta_norm", "Tabela Fee",
        "PL_ref_original", "Moeda_ref", "PTAX_ref", "PL_ref_BRL", "PL_Grupo_Ref_BRL",
        "Participacao_no_Grupo", "Fee_aa", "Regra_Fee", "Fee_Projetado_Mensal_Ref"
    ]
    return df[keep].copy(), group_view


def read_xp_daily(file, year: int, month: int) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="Taxa Clientes", dtype=object)
    ren = {"Codigo Cliente": "Conta", "Patrimonio Liquido": "PL_original", "Data Posicao": "Data"}
    df = df.rename(columns=ren)
    req = ["Conta", "PL_original", "Data"]
    if any(c not in df.columns for c in req):
        raise ValueError("Arquivo XP não possui Codigo Cliente / Patrimonio Liquido / Data Posicao.")
    df["Conta_norm"] = df["Conta"].map(norm_account)
    df["Data"] = df["Data"].map(excel_date)
    df["PL_original"] = pd.to_numeric(df["PL_original"], errors="coerce")
    start, end = month_bounds(year, month)
    df = df[df["Data"].between(start, end)].copy()
    df["Corretora"] = "XP"
    df["Moeda"] = "BRL"
    df["PTAX"] = 1.0
    df["PL_BRL"] = df["PL_original"]
    df["Metodo"] = "Diário / 252"
    return df[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]]


def read_btg_daily(files: Iterable, year: int, month: int) -> pd.DataFrame:
    parts = []
    for f in files:
        df = pd.read_excel(f, sheet_name=0, dtype=object)
        # Alguns exports podem vir sem headers reconhecidos; pega as 3 primeiras colunas.
        df = df.iloc[:, :3].copy()
        df.columns = ["Conta", "Nome_Fonte", "PL_original"]
        df["Conta_norm"] = df["Conta"].map(norm_account)
        df["PL_original"] = pd.to_numeric(df["PL_original"], errors="coerce")
        df = df[df["Conta_norm"].ne("") & df["PL_original"].notna()].copy()
        name = getattr(f, "name", "")
        m = re.search(r"(\d{1,2})[.\-_](\d{1,2})(?:[.\-_](\d{2,4}))?", name)
        if not m:
            raise ValueError(f"Não consegui identificar a data no nome do arquivo BTG: {name}")
        day, mon = int(m.group(1)), int(m.group(2))
        yr = int(m.group(3)) if m.group(3) else year
        if yr < 100:
            yr += 2000
        d = pd.Timestamp(yr, mon, day)
        if d.year != year or d.month != month:
            continue
        df["Data"] = d
        df["Corretora"] = "BTG"
        df["Moeda"] = "BRL"
        df["PTAX"] = 1.0
        df["PL_BRL"] = df["PL_original"]
        df["Metodo"] = "Diário / 252"
        parts.append(df[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]])
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def read_cs_daily(files: Iterable, year: int, month: int, manual_ptax: dict[date, float]) -> pd.DataFrame:
    parts = []
    for f in files:
        raw = f.getvalue() if hasattr(f, "getvalue") else f.read()
        text = raw.decode("utf-8-sig", errors="replace") if isinstance(raw, bytes) else str(raw)
        lines = text.splitlines()
        date_found = None
        for line in lines[:5]:
            m = re.search(r"as of (\d{1,2})/(\d{1,2})/(\d{2,4})", line, re.I)
            if m:
                mm, dd, yy = map(int, m.groups())
                if yy < 100:
                    yy += 2000
                date_found = date(yy, mm, dd)
                break
        if date_found is None:
            name = getattr(f, "name", "")
            m = re.search(r"(\d{1,2})[.\-_](\d{1,2})(?:[.\-_](\d{2,4}))?", name)
            if m:
                dd, mm = int(m.group(1)), int(m.group(2))
                yy = int(m.group(3)) if m.group(3) else year
                if yy < 100:
                    yy += 2000
                date_found = date(yy, mm, dd)
        if date_found is None:
            raise ValueError(f"Não consegui identificar a data do arquivo CS: {getattr(f, 'name', '')}")
        if date_found.year != year or date_found.month != month:
            continue

        header_idx = next((i for i, line in enumerate(lines) if line.startswith("Account,")), None)
        if header_idx is None:
            raise ValueError(f"Cabeçalho 'Account' não encontrado em {getattr(f, 'name', '')}")
        df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), dtype=object)
        df = df.rename(columns={"Account": "Conta", "Total account value": "PL_original"})
        df["Conta_norm"] = df["Conta"].map(norm_account)
        df["PL_original"] = (
            df["PL_original"].astype(str)
              .str.replace("$", "", regex=False)
              .str.replace(",", "", regex=False)
              .str.replace("(", "-", regex=False)
              .str.replace(")", "", regex=False)
        )
        df["PL_original"] = pd.to_numeric(df["PL_original"], errors="coerce")
        ptax, ptax_date, src = ptax_with_manual(date_found, manual_ptax)
        df["Data"] = pd.Timestamp(date_found)
        df["Corretora"] = "CHARLES SCHWAB"
        df["Moeda"] = "USD"
        df["PTAX"] = ptax
        df["Data_PTAX"] = pd.Timestamp(ptax_date)
        df["Fonte_PTAX"] = src
        df["PL_BRL"] = df["PL_original"] * ptax
        df["Metodo"] = "Diário / 252"
        parts.append(df[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]])
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def read_monthly_standardized(file, broker: str, year: int, month: int,
                              manual_ptax: dict[date, float]) -> pd.DataFrame:
    """Arquivo mensal opcional padronizado: Conta | Valor | Data | Moeda(opcional)."""
    df = pd.read_excel(file, sheet_name=0, dtype=object)
    cols = {norm_text(c).lower(): c for c in df.columns}
    conta = cols.get("conta")
    valor = cols.get("valor") or cols.get("pl") or cols.get("patrimonio liquido")
    data_col = cols.get("data") or cols.get("data posicao")
    moeda_col = cols.get("moeda")
    if not conta or not valor:
        raise ValueError(f"Arquivo mensal {broker}: use colunas Conta e Valor (Data/Moeda opcionais).")
    out = pd.DataFrame()
    out["Conta_norm"] = df[conta].map(norm_account)
    out["PL_original"] = pd.to_numeric(df[valor], errors="coerce")
    close = month_bounds(year, month)[1].date()
    if data_col:
        out["Data"] = df[data_col].map(excel_date)
    else:
        out["Data"] = pd.Timestamp(close)
    out["Corretora"] = broker
    default_currency = "USD" if broker in OFFSHORE_BROKERS else "BRL"
    if moeda_col:
        out["Moeda"] = df[moeda_col].map(lambda x: norm_text(x).upper() or default_currency)
    else:
        out["Moeda"] = default_currency
    ptax, _, _ = ptax_with_manual(close, manual_ptax) if default_currency == "USD" else (1.0, close, "BRL")
    out["PTAX"] = out["Moeda"].apply(lambda x: ptax if x == "USD" else 1.0)
    out["PL_BRL"] = out["PL_original"] * out["PTAX"]
    out["Metodo"] = "Mensal / 12"
    return out[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]]


def calculate_fees(mov: pd.DataFrame, registry: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if mov.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    reg = registry[["Conta_norm", "Corretora", "Grupo_chave", "Grupo Familiar", "Cliente", "Tabela Fee",
                    "PL_ref_BRL", "PL_Grupo_Ref_BRL", "Participacao_no_Grupo", "Fee_aa"]].copy()
    out = mov.merge(reg, on=["Conta_norm", "Corretora"], how="left", indicator=True)
    out["Status_Match"] = out["_merge"].map({"both": "OK", "left_only": "Conta não encontrada no Controle", "right_only": ""})
    out = out.drop(columns=["_merge"])
    out["Fee_Calculado"] = 0.0
    daily = out["Metodo"].eq("Diário / 252")
    monthly = out["Metodo"].eq("Mensal / 12")
    out.loc[daily, "Fee_Calculado"] = out.loc[daily, "PL_BRL"] * out.loc[daily, "Fee_aa"].fillna(0) / 252
    out.loc[monthly, "Fee_Calculado"] = out.loc[monthly, "PL_BRL"] * out.loc[monthly, "Fee_aa"].fillna(0) / 12

    summary = (
        out.groupby(["Corretora", "Conta_norm", "Grupo_chave", "Cliente", "Tabela Fee", "Fee_aa", "Metodo"], dropna=False)
           .agg(PL_Medio_BRL=("PL_BRL", "mean"),
                PL_Fechamento_BRL=("PL_BRL", "last"),
                Dias_ou_Registros=("Data", "nunique"),
                Fee_Mes=("Fee_Calculado", "sum"),
                Data_Inicial=("Data", "min"),
                Data_Final=("Data", "max"))
           .reset_index()
    )
    summary["Fee_Projetado_Anual_sobre_PL_Medio"] = summary["PL_Medio_BRL"] * summary["Fee_aa"].fillna(0)

    unmatched = out[out["Status_Match"] != "OK"][["Corretora", "Conta_norm", "Data", "PL_BRL", "Status_Match"]].drop_duplicates()
    return out, summary, unmatched


def expected_business_days(year: int, month: int, until: Optional[date] = None) -> int:
    start, end = month_bounds(year, month)
    if until:
        end = min(end, pd.Timestamp(until))
    return len(pd.bdate_range(start, end))


def export_excel(registry, groups, detail, summary, unmatched, year, month) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        registry.to_excel(writer, sheet_name="Cadastro Fee", index=False)
        groups.to_excel(writer, sheet_name="Grupos", index=False)
        if not summary.empty:
            summary.to_excel(writer, sheet_name="Resumo Cobranca", index=False)
        if not detail.empty:
            detail.to_excel(writer, sheet_name="Memoria Calculo", index=False)
            for broker in sorted(detail["Corretora"].dropna().unique()):
                safe = re.sub(r"[^A-Za-z0-9 _-]", "", broker)[:22]
                detail[detail["Corretora"] == broker].to_excel(writer, sheet_name=f"Calc {safe}"[:31], index=False)
        if not unmatched.empty:
            unmatched.to_excel(writer, sheet_name="Pendencias Match", index=False)

        # Formatação enxuta e financeira
        wb = writer.book
        fmt_money = wb.add_format({"num_format": 'R$ #,##0.00;[Red](R$ #,##0.00);-'})
        fmt_pct = wb.add_format({"num_format": '0.00%;[Red](0.00%);-'})
        fmt_hdr = wb.add_format({"bold": True, "bg_color": "#141C24", "font_color": "#FFFFFF", "border": 0})
        for sheet_name, ws in writer.sheets.items():
            ws.hide_gridlines(2)
            # Cabeçalho
            if sheet_name in ["Cadastro Fee", "Grupos", "Resumo Cobranca", "Memoria Calculo", "Pendencias Match"] or sheet_name.startswith("Calc "):
                # largura genérica
                ws.set_row(0, 22, fmt_hdr)
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, max(1, ws.dim_rowmax), max(0, ws.dim_colmax))
                ws.set_column(0, max(0, ws.dim_colmax), 17)
        # formatos por busca de nomes de coluna
        datasets = {
            "Cadastro Fee": registry, "Grupos": groups, "Resumo Cobranca": summary,
            "Memoria Calculo": detail, "Pendencias Match": unmatched
        }
        for sname, data in datasets.items():
            if sname not in writer.sheets or data is None or data.empty:
                continue
            ws = writer.sheets[sname]
            for idx, col in enumerate(data.columns):
                if any(k in str(col) for k in ["PL_", "Fee_"]) and "Fee_aa" not in str(col):
                    ws.set_column(idx, idx, 18, fmt_money)
                if col in ["Fee_aa", "Participacao_no_Grupo", "ROA_Contratado"]:
                    ws.set_column(idx, idx, 16, fmt_pct)
                if "Cliente" in str(col) or "Grupo" in str(col):
                    ws.set_column(idx, idx, 28)
    return output.getvalue()


# ----------------------------- UI -----------------------------
st.set_page_config(page_title="Cálculo de Fee | M Wealth", layout="wide")
st.title("Cálculo de Fee Mensal — M Wealth")
st.caption("Faixa por PL do grupo familiar · cálculo diário/252 · mensal/12 · offshore por PTAX")

with st.sidebar:
    st.header("Período")
    today = date.today()
    year = st.number_input("Ano da cobrança", min_value=2024, max_value=2100, value=today.year, step=1)
    month = st.number_input("Mês da cobrança", min_value=1, max_value=12, value=today.month, step=1)
    st.divider()
    st.subheader("PTAX manual (opcional)")
    st.caption("Use apenas se quiser travar uma cotação. Formato: DD/MM/AAAA;5,0770")
    ptax_text = st.text_area("Cotações", value="", height=100, label_visibility="collapsed")

manual_ptax = {}
for line in ptax_text.splitlines():
    if ";" not in line:
        continue
    ds, vs = line.split(";", 1)
    try:
        manual_ptax[datetime.strptime(ds.strip(), "%d/%m/%Y").date()] = float(vs.strip().replace(".", "").replace(",", "."))
    except Exception:
        pass

st.subheader("1. Base cadastral")
control_file = st.file_uploader("Controle de Clientes MWealth", type=["xlsx"], key="control")

st.subheader("2. PL para cobrança")
c1, c2, c3 = st.columns(3)
with c1:
    xp_file = st.file_uploader("XP — relatório com todos os dias", type=["xlsx"], key="xp")
with c2:
    btg_files = st.file_uploader("BTG — um arquivo por dia", type=["xlsx"], accept_multiple_files=True, key="btg")
with c3:
    cs_files = st.file_uploader("Charles Schwab — um CSV por dia", type=["csv"], accept_multiple_files=True, key="cs")

with st.expander("Fontes mensais (Safra / XP US)"):
    st.caption("Enquanto não fechamos o layout original dessas fontes, use um Excel com colunas Conta e Valor; Data e Moeda são opcionais.")
    safra_file = st.file_uploader("Safra — mensal", type=["xlsx"], key="safra")
    xpus_file = st.file_uploader("XP US — mensal", type=["xlsx"], key="xpus")

if control_file:
    try:
        ctrl, fee_tables, pl_cols = read_control(control_file)
        ref_col, ref_date = choose_reference_pl_column(pl_cols, int(year), int(month))
        registry, groups = build_fee_registry(ctrl, fee_tables, ref_col, ref_date, manual_ptax)

        total_pl = registry["PL_ref_BRL"].sum()
        fee_ann = registry["Fee_Projetado_Mensal_Ref"].sum() * 12
        roa = fee_ann / total_pl if total_pl else 0
        a, b, c, d = st.columns(4)
        a.metric("PL de referência", f"R$ {total_pl:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        b.metric("Data referência", ref_date.strftime("%d/%m/%Y"))
        c.metric("ROA médio contratado", f"{roa:.3%}")
        d.metric("Fee mensal projetado", f"R$ {fee_ann/12:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.caption("O PL do grupo usa a coluna de fechamento anterior e converte offshore para BRL pela PTAX da data de referência.")
        st.dataframe(registry[["Grupo_chave", "Cliente", "Corretora", "Conta_norm", "Tabela Fee", "PL_ref_BRL", "PL_Grupo_Ref_BRL", "Participacao_no_Grupo", "Fee_aa", "Regra_Fee"]], use_container_width=True, hide_index=True)

        if st.button("Calcular cobrança", type="primary", use_container_width=True):
            parts = []
            errors = []
            if xp_file:
                try: parts.append(read_xp_daily(xp_file, int(year), int(month)))
                except Exception as e: errors.append(f"XP: {e}")
            if btg_files:
                try: parts.append(read_btg_daily(btg_files, int(year), int(month)))
                except Exception as e: errors.append(f"BTG: {e}")
            if cs_files:
                try: parts.append(read_cs_daily(cs_files, int(year), int(month), manual_ptax))
                except Exception as e: errors.append(f"Charles Schwab: {e}")
            if safra_file:
                try: parts.append(read_monthly_standardized(safra_file, "SAFRA", int(year), int(month), manual_ptax))
                except Exception as e: errors.append(f"Safra: {e}")
            if xpus_file:
                try: parts.append(read_monthly_standardized(xpus_file, "XP US", int(year), int(month), manual_ptax))
                except Exception as e: errors.append(f"XP US: {e}")

            for e in errors:
                st.error(e)

            valid = [x for x in parts if x is not None and not x.empty]
            if not valid:
                st.warning("Nenhum PL válido foi carregado para o período.")
            else:
                mov = pd.concat(valid, ignore_index=True)
                detail, summary, unmatched = calculate_fees(mov, registry)

                st.subheader("Resultado da cobrança")
                m1, m2, m3 = st.columns(3)
                m1.metric("Fee calculado", f"R$ {detail['Fee_Calculado'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                m2.metric("Contas com cálculo", f"{summary['Conta_norm'].nunique():,}".replace(",", "."))
                m3.metric("Pendências de match", f"{len(unmatched):,}".replace(",", "."))

                st.dataframe(summary, use_container_width=True, hide_index=True)

                # Cobertura diária por corretora
                st.subheader("Cobertura de arquivos diários")
                max_date = detail["Data"].dropna().max()
                expected = expected_business_days(int(year), int(month), max_date.date() if pd.notna(max_date) else None)
                cov = (detail[detail["Metodo"].eq("Diário / 252")]
                       .groupby("Corretora")["Data"].nunique().rename("Dias encontrados").reset_index())
                cov["Dias úteis esperados até a última data"] = expected
                cov["Cobertura"] = cov["Dias encontrados"] / expected if expected else 0
                st.dataframe(cov, use_container_width=True, hide_index=True)

                if not unmatched.empty:
                    st.warning("Há contas nas fontes que não foram encontradas no Controle. Veja a aba Pendências Match no download.")

                excel = export_excel(registry, groups, detail, summary, unmatched, int(year), int(month))
                st.download_button(
                    "Baixar memória completa (.xlsx)",
                    data=excel,
                    file_name=f"Fee_MWealth_{int(year)}_{int(month):02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                st.subheader("Downloads por corretora")
                for broker in sorted(detail["Corretora"].dropna().unique()):
                    sub = detail[detail["Corretora"] == broker]
                    csv_bytes = sub.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
                    st.download_button(
                        f"Baixar cálculo — {broker}",
                        data=csv_bytes,
                        file_name=f"Calculo_Fee_{broker.replace(' ', '_')}_{int(year)}_{int(month):02d}.csv",
                        mime="text/csv",
                        key=f"dl_{broker}",
                    )

    except Exception as e:
        st.exception(e)
else:
    st.info("Comece pelo Controle de Clientes. Ele é a base-mãe do cálculo.")
