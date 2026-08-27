from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Iterable, Optional

import pandas as pd
import requests
import streamlit as st
from PIL import Image

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
MONTHLY_FROM_CONTROL_BROKERS = {"SAFRA", "XP US"}
APP_DIR = Path(__file__).resolve().parent
CONTROL_CANDIDATES = [
    APP_DIR / "data" / "controle_clientes.xlsx",
    APP_DIR / "controle_clientes.xlsx",
    APP_DIR / "Controle de Clientes MWealth 2026.xlsx",
]
NO_FEE_LABELS = {"SEM COBRANÇA", "SEM COBRANCA", "NÃO COBRAR", "NAO COBRAR"}
ICON_CANDIDATES = [APP_DIR / "M_light.png", APP_DIR / "M light.png"]


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


def choose_billing_pl_column(pl_cols: list[str], year: int, month: int) -> tuple[str, date, bool]:
    """Última posição disponível dentro do mês de cobrança.

    Retorna (coluna, data, eh_fechamento). O fechamento é considerado definitivo
    quando a data é o último dia útil do mês ou posterior ao penúltimo dia corrido.
    Enquanto o mês estiver aberto, a posição é marcada como prévia.
    """
    candidates = []
    for c in pl_cols:
        try:
            d = datetime.strptime(c[3:], "%d/%m/%Y").date()
            if d.year == year and d.month == month:
                candidates.append((d, c))
        except Exception:
            continue
    if not candidates:
        raise ValueError(f"Não encontrei coluna de PL no Controle para {month:02d}/{year}.")
    d, c = max(candidates)
    month_end = (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(1)).date()
    last_business = pd.bdate_range(pd.Timestamp(year, month, 1), pd.Timestamp(month_end))[-1].date()
    is_close = d >= last_business
    return c, d, is_close


def find_fixed_control_path() -> Optional[Path]:
    for candidate in CONTROL_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


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
        return fixed, "Sem cobrança" if fixed == 0 else "Taxa fixa"

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
    """Lê posições diárias em USD e converte todo o mês pela PTAX de fechamento.

    Enquanto o mês estiver incompleto, usa a data mais recente entre os arquivos
    carregados como PTAX de prévia. No fechamento, essa data será o último dia
    útil disponível do mês.
    """
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
        df["Data"] = pd.Timestamp(date_found)
        df["Corretora"] = "CHARLES SCHWAB"
        df["Moeda"] = "USD"
        parts.append(df[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda"]])

    if not parts:
        return pd.DataFrame()

    out = pd.concat(parts, ignore_index=True)
    close_date = out["Data"].max().date()
    ptax, ptax_date, src = ptax_with_manual(close_date, manual_ptax)
    out["PTAX"] = ptax
    out["Data_PTAX"] = pd.Timestamp(ptax_date)
    out["Fonte_PTAX"] = src
    out["PL_BRL"] = out["PL_original"] * ptax
    out["Metodo"] = "Diário / 252"
    return out[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]]

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



def build_monthly_from_control(registry: pd.DataFrame, ctrl: pd.DataFrame, pl_cols: list[str],
                               year: int, month: int, manual_ptax: dict[date, float]) -> tuple[pd.DataFrame, dict]:
    """Gera cobrança mensal de Safra e XP US diretamente do Controle de Clientes."""
    billing_col, billing_date, is_close = choose_billing_pl_column(pl_cols, year, month)
    base = ctrl[ctrl["Corretora"].isin(MONTHLY_FROM_CONTROL_BROKERS)].copy()
    base["PL_original"] = pd.to_numeric(base[billing_col], errors="coerce").fillna(0.0)
    base["Data"] = pd.Timestamp(billing_date)
    base["Moeda"] = base["Corretora"].apply(lambda x: "USD" if x == "XP US" else "BRL")

    if (base["Moeda"] == "USD").any():
        ptax, ptax_date, src = ptax_with_manual(billing_date, manual_ptax)
    else:
        ptax, ptax_date, src = 1.0, billing_date, "BRL"
    base["PTAX"] = base["Moeda"].apply(lambda x: ptax if x == "USD" else 1.0)
    base["PL_BRL"] = base["PL_original"] * base["PTAX"]
    base["Metodo"] = "Mensal / 12 (Controle)"

    out = base[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]].copy()
    meta = {
        "billing_col": billing_col,
        "billing_date": billing_date,
        "is_close": is_close,
        "ptax": ptax if (base["Moeda"] == "USD").any() else None,
        "ptax_date": ptax_date if (base["Moeda"] == "USD").any() else None,
        "ptax_source": src if (base["Moeda"] == "USD").any() else None,
    }
    return out, meta

def calculate_fees(mov: pd.DataFrame, registry: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calcula somente contas casadas com o Controle.

    Contas extras das fontes nunca entram na cobrança; o diagnóstico de cadastro
    é produzido separadamente por build_match_diagnostics().
    """
    if mov.empty:
        return pd.DataFrame(), pd.DataFrame()

    reg = registry[[
        "Conta_norm", "Corretora", "Grupo_chave", "Grupo Familiar", "Cliente", "Tabela Fee",
        "PL_ref_BRL", "PL_Grupo_Ref_BRL", "Participacao_no_Grupo", "Fee_aa", "Regra_Fee"
    ]].copy()

    out = mov.merge(reg, on=["Conta_norm", "Corretora"], how="inner")
    if out.empty:
        return pd.DataFrame(), pd.DataFrame()

    out["Fee_Calculado"] = 0.0
    daily = out["Metodo"].eq("Diário / 252")
    monthly = out["Metodo"].str.startswith("Mensal / 12", na=False)
    out.loc[daily, "Fee_Calculado"] = out.loc[daily, "PL_BRL"] * out.loc[daily, "Fee_aa"].fillna(0) / 252
    out.loc[monthly, "Fee_Calculado"] = out.loc[monthly, "PL_BRL"] * out.loc[monthly, "Fee_aa"].fillna(0) / 12

    out = out.sort_values(["Corretora", "Conta_norm", "Data"]).reset_index(drop=True)

    summary = (
        out.groupby([
            "Corretora", "Conta_norm", "Grupo_chave", "Cliente", "Tabela Fee", "Fee_aa", "Regra_Fee", "Metodo", "Moeda"
        ], dropna=False)
        .agg(
            PL_Medio_Original=("PL_original", "mean"),
            PL_Fechamento_Original=("PL_original", "last"),
            PL_Medio_BRL=("PL_BRL", "mean"),
            PL_Fechamento_BRL=("PL_BRL", "last"),
            Dias_ou_Registros=("Data", "nunique"),
            Fee_Mes=("Fee_Calculado", "sum"),
            Data_Inicial=("Data", "min"),
            Data_Final=("Data", "max"),
        )
        .reset_index()
    )
    summary["PL_Medio_USD"] = summary["PL_Medio_Original"].where(summary["Moeda"].eq("USD"))
    summary["PL_Fechamento_USD"] = summary["PL_Fechamento_Original"].where(summary["Moeda"].eq("USD"))
    summary["Fee_Projetado_Anual_sobre_PL_Medio"] = summary["PL_Medio_BRL"] * summary["Fee_aa"].fillna(0)
    return out, summary


def build_match_diagnostics(mov: pd.DataFrame, registry: pd.DataFrame) -> pd.DataFrame:
    """Diagnóstico objetivo de cadastro para XP e BTG.

    - PL -> Controle: conta apareceu na fonte e não existe no Controle.
    - Controle -> PL: conta existe no Controle e não apareceu em nenhuma posição carregada da corretora.
    - Charles Schwab extra é ignorada integralmente, conforme regra operacional.
    """
    rows = []
    for broker in ["XP", "BTG"]:
        src = mov[mov["Corretora"].eq(broker)].copy()
        if src.empty:
            continue
        reg = registry[registry["Corretora"].eq(broker)].copy()
        source_accounts = set(src["Conta_norm"].dropna().astype(str)) - {""}
        control_accounts = set(reg["Conta_norm"].dropna().astype(str)) - {""}

        for account in sorted(source_accounts - control_accounts):
            ss = src[src["Conta_norm"].eq(account)]
            dates = sorted(pd.to_datetime(ss["Data"].dropna()).dt.date.unique())
            rows.append({
                "Corretora": broker,
                "Conta": account,
                "Cliente": "",
                "Grupo Familiar": "",
                "Problema": "Está no arquivo de PL, mas não está no Controle",
                "Origem": "PL → Controle",
                "Registros": int(len(ss)),
                "Primeira Data": pd.Timestamp(min(dates)) if dates else pd.NaT,
                "Última Data": pd.Timestamp(max(dates)) if dates else pd.NaT,
            })

        missing = reg[reg["Conta_norm"].isin(control_accounts - source_accounts)].copy()
        for _, r in missing.iterrows():
            rows.append({
                "Corretora": broker,
                "Conta": r["Conta_norm"],
                "Cliente": r["Cliente"],
                "Grupo Familiar": r["Grupo_chave"],
                "Problema": "Está no Controle, mas não apareceu no arquivo de PL",
                "Origem": "Controle → PL",
                "Registros": 0,
                "Primeira Data": pd.NaT,
                "Última Data": pd.NaT,
            })
    return pd.DataFrame(rows)


def build_coverage(mov: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    rows = []
    start, month_end = month_bounds(year, month)
    today = pd.Timestamp(date.today())
    if month_end < today.normalize():
        expected_end = month_end
    elif start <= today.normalize() <= month_end:
        expected_end = today.normalize()
    else:
        expected_end = month_end

    expected_dates = list(pd.bdate_range(start, expected_end).date)
    for broker in ["XP", "BTG", "CHARLES SCHWAB"]:
        sub = mov[mov["Corretora"].eq(broker)]
        if sub.empty:
            continue
        found_dates = sorted(pd.to_datetime(sub["Data"].dropna()).dt.date.unique())
        found_set = set(found_dates)
        missing = [d for d in expected_dates if d not in found_set]
        rows.append({
            "Corretora": broker,
            "Primeira Data": pd.Timestamp(min(found_dates)) if found_dates else pd.NaT,
            "Última Data": pd.Timestamp(max(found_dates)) if found_dates else pd.NaT,
            "Dias Encontrados": len(found_dates),
            "Dias Úteis Esperados": len(expected_dates),
            "Cobertura": len(found_dates) / len(expected_dates) if expected_dates else 0,
            "Datas Ausentes": ", ".join(d.strftime("%d/%m/%Y") for d in missing) if missing else "Nenhuma",
        })
    return pd.DataFrame(rows)


def build_account_fee_summary(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame()
    return (
        summary.groupby(["Corretora", "Conta_norm", "Cliente", "Grupo_chave", "Fee_aa", "Regra_Fee"], dropna=False)
        .agg(Fee_Total=("Fee_Mes", "sum"), PL_Base_BRL=("PL_Medio_BRL", "mean"))
        .reset_index()
        .sort_values(["Corretora", "Cliente", "Conta_norm"])
    )


def find_icon_path() -> Optional[Path]:
    for p in ICON_CANDIDATES:
        if p.exists():
            return p
    return None


def parse_manual_ptax(text: str) -> dict[date, float]:
    manual = {}
    if not text:
        return manual
    for item in re.split(r"\s*\|\s*|\n+", text.strip()):
        if ";" not in item:
            continue
        ds, vs = item.split(";", 1)
        try:
            cleaned = vs.strip().replace(" ", "")
            if "," in cleaned:
                cleaned = cleaned.replace(".", "").replace(",", ".")
            manual[datetime.strptime(ds.strip(), "%d/%m/%Y").date()] = float(cleaned)
        except Exception:
            continue
    return manual


def fmt_brl(v) -> str:
    if pd.isna(v):
        return "—"
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_usd(v) -> str:
    if pd.isna(v):
        return "—"
    return f"$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v, decimals=2) -> str:
    if pd.isna(v):
        return "—"
    return f"{float(v) * 100:.{decimals}f}%".replace(".", ",")


def registry_display(registry: pd.DataFrame) -> pd.DataFrame:
    d = registry.copy()
    d["PL Referência USD"] = d["PL_ref_original"].where(d["Moeda_ref"].eq("USD")).map(fmt_usd)
    d["PL Referência BRL"] = d["PL_ref_BRL"].map(fmt_brl)
    d["PL Grupo Familiar"] = d["PL_Grupo_Ref_BRL"].map(fmt_brl)
    d["Participação"] = d["Participacao_no_Grupo"].map(fmt_pct)
    d["Fee a.a."] = d["Fee_aa"].map(lambda x: fmt_pct(x, 3))
    d["Fee Mensal Teórico"] = d["Fee_Projetado_Mensal_Ref"].map(fmt_brl)
    return d.rename(columns={
        "Grupo_chave": "Grupo Familiar", "Conta_norm": "Conta", "Tabela Fee": "Tabela",
        "Regra_Fee": "Regra do Fee", "Corretora": "Corretora", "Cliente": "Cliente"
    })[["Grupo Familiar", "Cliente", "Corretora", "Conta", "Tabela", "Regra do Fee",
        "PL Referência USD", "PL Referência BRL", "PL Grupo Familiar", "Participação", "Fee a.a.", "Fee Mensal Teórico"]]


def summary_display(summary: pd.DataFrame) -> pd.DataFrame:
    d = summary.copy()
    d["PL Médio USD"] = d["PL_Medio_USD"].map(fmt_usd)
    d["PL Médio BRL"] = d["PL_Medio_BRL"].map(fmt_brl)
    d["PL Fechamento USD"] = d["PL_Fechamento_USD"].map(fmt_usd)
    d["PL Fechamento BRL"] = d["PL_Fechamento_BRL"].map(fmt_brl)
    d["Fee a.a."] = d["Fee_aa"].map(lambda x: fmt_pct(x, 3))
    d["Fee do Mês"] = d["Fee_Mes"].map(fmt_brl)
    d["Período"] = d["Data_Inicial"].dt.strftime("%d/%m/%Y") + " a " + d["Data_Final"].dt.strftime("%d/%m/%Y")
    return d.rename(columns={
        "Grupo_chave": "Grupo Familiar", "Conta_norm": "Conta", "Tabela Fee": "Tabela",
        "Regra_Fee": "Regra do Fee", "Metodo": "Metodologia", "Dias_ou_Registros": "Dias/Registros"
    })[["Grupo Familiar", "Cliente", "Corretora", "Conta", "Tabela", "Regra do Fee", "Metodologia",
        "PL Médio USD", "PL Médio BRL", "PL Fechamento USD", "PL Fechamento BRL", "Fee a.a.", "Dias/Registros", "Fee do Mês", "Período"]]


def diagnostics_display(diag: pd.DataFrame) -> pd.DataFrame:
    if diag.empty:
        return diag
    d = diag.copy()
    d["Primeira Data"] = pd.to_datetime(d["Primeira Data"]).dt.strftime("%d/%m/%Y").fillna("—")
    d["Última Data"] = pd.to_datetime(d["Última Data"]).dt.strftime("%d/%m/%Y").fillna("—")
    return d[["Corretora", "Conta", "Cliente", "Grupo Familiar", "Origem", "Problema", "Registros", "Primeira Data", "Última Data"]]


def coverage_display(cov: pd.DataFrame) -> pd.DataFrame:
    if cov.empty:
        return cov
    d = cov.copy()
    d["Primeira Data"] = pd.to_datetime(d["Primeira Data"]).dt.strftime("%d/%m/%Y")
    d["Última Data"] = pd.to_datetime(d["Última Data"]).dt.strftime("%d/%m/%Y")
    d["Cobertura"] = d["Cobertura"].map(fmt_pct)
    return d


def _pretty_export_frames(registry, groups, detail, summary, diagnostics, coverage, account_summary):
    reg = registry.copy().rename(columns={
        "Grupo_chave": "Grupo Familiar (Chave)", "Conta_norm": "Conta", "Tabela Fee": "Tabela Fee",
        "PL_ref_original": "PL Referência Original", "Moeda_ref": "Moeda", "PTAX_ref": "PTAX Referência",
        "PL_ref_BRL": "PL Referência BRL", "PL_Grupo_Ref_BRL": "PL Grupo Familiar BRL",
        "Participacao_no_Grupo": "Participação no Grupo", "Fee_aa": "Fee a.a.", "Regra_Fee": "Regra do Fee",
        "Fee_Projetado_Mensal_Ref": "Fee Mensal Teórico"
    })
    reg["PL Referência USD"] = reg["PL Referência Original"].where(reg["Moeda"].eq("USD"))

    grp = groups.copy().rename(columns={
        "Grupo_chave": "Grupo Familiar", "PL_Grupo_Ref_BRL": "PL Grupo Familiar BRL", "Contas": "Contas",
        "Fee_Projetado_Anual": "Fee Projetado Anual", "ROA_Contratado": "ROA Contratado",
        "Data_Referencia": "Data de Referência", "PTAX_Referencia": "PTAX Referência",
        "Data_PTAX_Utilizada": "Data PTAX Utilizada", "Fonte_PTAX": "Fonte PTAX"
    })

    summ = summary.copy().rename(columns={
        "Conta_norm": "Conta", "Grupo_chave": "Grupo Familiar", "Tabela Fee": "Tabela Fee", "Fee_aa": "Fee a.a.",
        "Regra_Fee": "Regra do Fee", "Metodo": "Metodologia", "PL_Medio_Original": "PL Médio Original",
        "PL_Fechamento_Original": "PL Fechamento Original", "PL_Medio_BRL": "PL Médio BRL",
        "PL_Fechamento_BRL": "PL Fechamento BRL", "Dias_ou_Registros": "Dias/Registros", "Fee_Mes": "Fee do Mês",
        "Data_Inicial": "Data Inicial", "Data_Final": "Data Final", "PL_Medio_USD": "PL Médio USD",
        "PL_Fechamento_USD": "PL Fechamento USD", "Fee_Projetado_Anual_sobre_PL_Medio": "Fee Anual sobre PL Médio"
    })

    det = detail.copy().rename(columns={
        "Conta_norm": "Conta", "PL_original": "PL Original", "PL_BRL": "PL BRL", "Metodo": "Metodologia",
        "Grupo_chave": "Grupo Familiar", "Tabela Fee": "Tabela Fee", "PL_ref_BRL": "PL Referência BRL",
        "PL_Grupo_Ref_BRL": "PL Grupo Familiar BRL", "Participacao_no_Grupo": "Participação no Grupo",
        "Fee_aa": "Fee a.a.", "Regra_Fee": "Regra do Fee", "Fee_Calculado": "Fee Calculado"
    })
    det["PL USD"] = det["PL Original"].where(det["Moeda"].eq("USD"))

    acct = account_summary.copy().rename(columns={
        "Conta_norm": "Conta", "Grupo_chave": "Grupo Familiar", "Fee_aa": "Fee a.a.", "Regra_Fee": "Regra do Fee",
        "Fee_Total": "Fee Total", "PL_Base_BRL": "PL Base BRL"
    })
    return reg, grp, summ, det, diagnostics.copy(), coverage.copy(), acct


def export_excel(registry, groups, detail, summary, diagnostics, coverage, account_summary, year, month) -> bytes:
    output = io.BytesIO()
    reg, grp, summ, det, diag, cov, acct = _pretty_export_frames(
        registry, groups, detail, summary, diagnostics, coverage, account_summary
    )
    daily = det[det["Metodologia"].eq("Diário / 252")].copy() if not det.empty else pd.DataFrame()
    monthly = det[det["Metodologia"].str.startswith("Mensal / 12", na=False)].copy() if not det.empty else pd.DataFrame()

    datasets = {
        "Resumo por Conta": acct,
        "Resumo Cobranca": summ,
        "Cadastro Fee": reg,
        "Grupos": grp,
        "Cobertura Diaria": cov,
        "Pendencias Cadastro": diag,
        "Memoria Calculo": det,
        "Calculo Diario": daily,
        "Calculo Mensal": monthly,
    }
    if not det.empty:
        for broker in sorted(det["Corretora"].dropna().unique()):
            safe = re.sub(r"[^A-Za-z0-9 _-]", "", broker)[:22]
            datasets[f"Calc {safe}"[:31]] = det[det["Corretora"].eq(broker)].copy()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        for sname, data in datasets.items():
            if data is not None and not data.empty:
                data.to_excel(writer, sheet_name=sname, index=False)

        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#141C24", "font_color": "#FFFFFF", "border": 0, "align": "center", "valign": "vcenter"})
        fmt_brl_x = wb.add_format({"num_format": 'R$ #,##0.00;[Red](R$ #,##0.00);-'})
        fmt_usd_x = wb.add_format({"num_format": '$ #,##0.00;[Red]($ #,##0.00);-'})
        fmt_pct_x = wb.add_format({"num_format": '0.000%;[Red](0.000%);-'})
        fmt_ptax = wb.add_format({"num_format": '0.0000'})
        fmt_int = wb.add_format({"num_format": '#,##0'})

        for sname, data in datasets.items():
            if sname not in writer.sheets or data is None or data.empty:
                continue
            ws = writer.sheets[sname]
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_row(0, 24, fmt_header)
            ws.autofilter(0, 0, len(data), len(data.columns)-1)
            for i, col in enumerate(data.columns):
                name = str(col)
                width = 16
                fmt = None
                if "Cliente" in name or "Grupo Familiar" in name or "Problema" in name or "Datas Ausentes" in name:
                    width = 30 if "Datas Ausentes" not in name else 55
                elif "Corretora" in name or "Metodologia" in name or "Regra" in name or "Tabela" in name:
                    width = 20
                elif "Data" in name or name == "Período":
                    width = 15
                elif "Conta" in name:
                    width = 15
                if "USD" in name or ("Original" in name and "Moeda" in data.columns):
                    fmt = fmt_usd_x if "USD" in name else None
                if "BRL" in name or name.startswith("Fee ") or name in {"Fee Total", "Fee Calculado", "Fee do Mês", "Fee Mensal Teórico", "Fee Projetado Anual", "Fee Anual sobre PL Médio"}:
                    fmt = fmt_brl_x
                if "%" in name or name in {"Fee a.a.", "ROA Contratado", "Participação no Grupo", "Cobertura"}:
                    fmt = fmt_pct_x
                if "PTAX" in name:
                    fmt = fmt_ptax
                if name in {"Contas", "Registros", "Dias/Registros", "Dias Encontrados", "Dias Úteis Esperados"}:
                    fmt = fmt_int
                ws.set_column(i, i, width, fmt)
    return output.getvalue()


# ----------------------------- UI -----------------------------
icon_path = find_icon_path()
page_icon = Image.open(icon_path) if icon_path else "M"
st.set_page_config(page_title="Cálculo de Fee | M Wealth", page_icon=page_icon, layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 2.0rem; padding-bottom: 3rem; max-width: 1500px;}
    h1 {font-size: 2rem !important; margin-bottom: .35rem !important;}
    h2, h3 {letter-spacing: -0.02em;}
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.035);
        border: 1px solid rgba(128,128,128,0.18);
        padding: 14px 16px;
        border-radius: 12px;
        min-height: 105px;
    }
    div[data-testid="stMetricLabel"] {font-size: .82rem; opacity: .72;}
    div[data-testid="stMetricValue"] {font-size: 1.35rem;}
    div[data-testid="stFileUploader"] {border-radius: 12px;}
    .small-note {font-size: .82rem; opacity: .68; margin-top: -6px;}
    .section-gap {height: .35rem;}
    [data-testid="stDataFrame"] {border: 1px solid rgba(128,128,128,.16); border-radius: 10px; overflow: hidden;}
</style>
""", unsafe_allow_html=True)

st.title("Cálculo de Fee Mensal — M Wealth")

today = date.today()
month_names = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
sel1, sel2, sel3, spacer = st.columns([1.0, 1.25, 3.1, 4.65])
with sel1:
    year = st.selectbox("Ano", list(range(today.year-2, today.year+3)), index=2)
with sel2:
    month_name = st.selectbox("Mês", list(month_names.values()), index=today.month-1)
    month = next(k for k,v in month_names.items() if v == month_name)
with sel3:
    ptax_text = st.text_input("PTAX manual · opcional", placeholder="31/07/2026;5,0770")
manual_ptax = parse_manual_ptax(ptax_text)

control_path = find_fixed_control_path()
if control_path is None:
    st.error("Base fixa não encontrada. Mantenha 'Controle de Clientes MWealth 2026.xlsx' na raiz do projeto.")
    st.stop()

try:
    ctrl, fee_tables, pl_cols = read_control(control_path)
    ref_col, ref_date = choose_reference_pl_column(pl_cols, int(year), int(month))
    registry, groups = build_fee_registry(ctrl, fee_tables, ref_col, ref_date, manual_ptax)
except Exception as e:
    st.error(f"Erro ao carregar a base fixa: {e}")
    st.stop()

latest_pl_date = max(datetime.strptime(c[3:], "%d/%m/%Y").date() for c in pl_cols)
accounts_count = int(ctrl["Conta_norm"].ne("").sum())
groups_count = int(ctrl["Grupo Familiar"].where(ctrl["Grupo Familiar"].ne(""), ctrl["Cliente"]).nunique())
total_pl = registry["PL_ref_BRL"].sum()
fee_ann = registry["Fee_Projetado_Mensal_Ref"].sum() * 12
roa = fee_ann / total_pl if total_pl else 0

st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
st.subheader("Visão geral")
r1, r2, r3, r4, r5 = st.columns(5)
r1.metric("PL de referência", fmt_brl(total_pl))
r2.metric("ROA médio contratado", fmt_pct(roa, 3))
r3.metric("Fee mensal teórico", fmt_brl(fee_ann/12))
r4.metric("Contas cadastradas", f"{accounts_count:,}".replace(",", "."))
r5.metric("Grupos familiares", f"{groups_count:,}".replace(",", "."))
st.markdown(f'<div class="small-note">Referência da faixa: {ref_date:%d/%m/%Y} · Último PL disponível no Controle: {latest_pl_date:%d/%m/%Y}</div>', unsafe_allow_html=True)

with st.expander("Cadastro de fee e segmentação", expanded=False):
    st.dataframe(registry_display(registry), use_container_width=True, hide_index=True, height=420)

st.subheader("Arquivos operacionais do mês")
st.caption("Safra e XP US não precisam de upload: o PL mensal é lido diretamente do Controle de Clientes.")
c1, c2, c3 = st.columns(3)
with c1:
    xp_file = st.file_uploader("XP — relatório com todos os dias", type=["xlsx"], key="xp")
with c2:
    btg_files = st.file_uploader("BTG — um arquivo por dia", type=["xlsx"], accept_multiple_files=True, key="btg")
with c3:
    cs_files = st.file_uploader("Charles Schwab — um CSV por dia", type=["csv"], accept_multiple_files=True, key="cs")

monthly_preview = None
monthly_meta = None
try:
    monthly_preview, monthly_meta = build_monthly_from_control(registry, ctrl, pl_cols, int(year), int(month), manual_ptax)
except Exception:
    monthly_preview = None

if st.button("Calcular cobrança", type="primary", use_container_width=True):
    parts = []
    errors = []
    if xp_file:
        try:
            parts.append(read_xp_daily(xp_file, int(year), int(month)))
        except Exception as e:
            errors.append(f"XP: {e}")
    if btg_files:
        try:
            parts.append(read_btg_daily(btg_files, int(year), int(month)))
        except Exception as e:
            errors.append(f"BTG: {e}")
    if cs_files:
        try:
            parts.append(read_cs_daily(cs_files, int(year), int(month), manual_ptax))
        except Exception as e:
            errors.append(f"Charles Schwab: {e}")
    if monthly_preview is not None and not monthly_preview.empty:
        parts.append(monthly_preview)

    for e in errors:
        st.error(e)

    valid = [x for x in parts if x is not None and not x.empty]
    if not valid:
        st.warning("Nenhum PL válido foi encontrado para o período.")
    else:
        mov = pd.concat(valid, ignore_index=True)
        diagnostics = build_match_diagnostics(mov, registry)
        coverage = build_coverage(mov[mov["Metodo"].eq("Diário / 252")].copy(), int(year), int(month))
        detail, summary = calculate_fees(mov, registry)
        account_summary = build_account_fee_summary(summary)

        daily_summary = summary[summary["Metodo"].eq("Diário / 252")].copy() if not summary.empty else pd.DataFrame()
        monthly_summary = summary[summary["Metodo"].str.startswith("Mensal / 12", na=False)].copy() if not summary.empty else pd.DataFrame()
        daily_detail = detail[detail["Metodo"].eq("Diário / 252")].copy() if not detail.empty else pd.DataFrame()
        monthly_detail = detail[detail["Metodo"].str.startswith("Mensal / 12", na=False)].copy() if not detail.empty else pd.DataFrame()

        st.subheader("Resultado da cobrança")
        m1, m2, m3, m4 = st.columns(4)
        total_fee = detail["Fee_Calculado"].sum() if not detail.empty else 0
        m1.metric("Fee total", fmt_brl(total_fee))
        m2.metric("Fee cálculo diário", fmt_brl(daily_detail["Fee_Calculado"].sum() if not daily_detail.empty else 0))
        m3.metric("Fee cálculo mensal", fmt_brl(monthly_detail["Fee_Calculado"].sum() if not monthly_detail.empty else 0))
        m4.metric("Pendências de cadastro", f"{len(diagnostics):,}".replace(",", "."))

        st.markdown("#### Fee total por conta")
        if account_summary.empty:
            st.info("Sem contas calculadas.")
        else:
            acct_disp = account_summary.rename(columns={"Conta_norm":"Conta", "Grupo_chave":"Grupo Familiar", "Fee_aa":"Fee a.a.", "Regra_Fee":"Regra do Fee"}).copy()
            acct_disp["PL Base BRL"] = acct_disp["PL_Base_BRL"].map(fmt_brl)
            acct_disp["Fee a.a."] = acct_disp["Fee a.a."].map(lambda x: fmt_pct(x,3))
            acct_disp["Fee Total"] = acct_disp["Fee_Total"].map(fmt_brl)
            st.dataframe(acct_disp[["Grupo Familiar","Cliente","Corretora","Conta","Regra do Fee","Fee a.a.","PL Base BRL","Fee Total"]], use_container_width=True, hide_index=True, height=360)

        tab1, tab2, tab3 = st.tabs(["Cobrança consolidada", "Cálculo diário", "Cálculo mensal"])
        with tab1:
            st.dataframe(summary_display(summary), use_container_width=True, hide_index=True, height=440)
        with tab2:
            if daily_summary.empty:
                st.info("Nenhum arquivo diário carregado para o período.")
            else:
                st.dataframe(summary_display(daily_summary), use_container_width=True, hide_index=True, height=440)
        with tab3:
            if monthly_summary.empty:
                st.info("Safra/XP US sem posição mensal disponível no Controle.")
            else:
                st.dataframe(summary_display(monthly_summary), use_container_width=True, hide_index=True, height=440)

        st.subheader("Cobertura dos arquivos diários")
        if coverage.empty:
            st.info("Nenhuma fonte diária carregada.")
        else:
            st.dataframe(coverage_display(coverage), use_container_width=True, hide_index=True)
            missing_brokers = coverage[coverage["Datas Ausentes"].ne("Nenhuma")]
            if not missing_brokers.empty:
                st.warning("Há datas úteis sem arquivo/posição. Confira a coluna 'Datas Ausentes' antes do fechamento.")

        st.subheader("Conferência de contas")
        st.caption("Charles Schwab: contas que aparecem no CS Total e não existem no Controle são ignoradas automaticamente e não viram pendência.")
        if diagnostics.empty:
            st.success("XP e BTG sem divergências entre contas carregadas e Controle.")
        else:
            st.dataframe(diagnostics_display(diagnostics), use_container_width=True, hide_index=True, height=380)

        excel = export_excel(registry, groups, detail, summary, diagnostics, coverage, account_summary, int(year), int(month))
        st.download_button(
            "Baixar memória completa (.xlsx)", data=excel,
            file_name=f"Fee_MWealth_{int(year)}_{int(month):02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.subheader("Downloads por corretora")
        if not detail.empty:
            cols = st.columns(min(4, max(1, detail["Corretora"].nunique())))
            for i, broker in enumerate(sorted(detail["Corretora"].dropna().unique())):
                sub = detail[detail["Corretora"].eq(broker)].copy()
                csv_bytes = sub.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
                with cols[i % len(cols)]:
                    st.download_button(
                        f"Baixar — {broker}", data=csv_bytes,
                        file_name=f"Calculo_Fee_{broker.replace(' ', '_')}_{int(year)}_{int(month):02d}.csv",
                        mime="text/csv", key=f"dl_{broker}", use_container_width=True,
                    )
