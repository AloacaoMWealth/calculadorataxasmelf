from __future__ import annotations

import io
import re
from pathlib import Path
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
# 1) O PL do GRUPO FAMILIAR no fechamento do mês calculado define a faixa.
# 2) Cada conta usa a Tabela Fee cadastrada no Controle de Clientes.
# 3) Contas com PL diário: Fee dia = PL_BRL * taxa_aa / 250.
# 4) Contas mensais: Fee mês = PL_BRL * taxa_aa / 12.
# 5) Offshore (Charles Schwab / XP US) é convertido para BRL pela PTAX.
# ============================================================

OFFSHORE_BROKERS = {"CHARLES SCHWAB", "XP US", "XP GLOBAL"}
MONTHLY_FROM_CONTROL_BROKERS = {"SAFRA", "XP US", "XP GLOBAL"}
DAILY_BASIS = 250
SEGMENTATION_AUX_SHEETS = {"Contas Atívore": "ATIVORE"}
WEALTH_NAVY = "#141C24"
WEALTH_NAVY_2 = "#131925"
WEALTH_SAND = "#DCC9B1"
WEALTH_BLUE = "#5D73AF"
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
    """Normaliza contas sem perder identificadores alfanuméricos (ex.: XP Global)."""
    raw = norm_text(v).upper()
    if not raw:
        return ""
    # Contas alfanuméricas precisam manter as letras; remove apenas máscara/espaços.
    if re.search(r"[A-Z]", raw):
        return re.sub(r"[^A-Z0-9]", "", raw)
    digits = re.sub(r"\D", "", raw)
    if not digits:
        return ""
    return str(int(digits))


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
                # Usa o último boletim disponível do dia; cotacaoVenda = BRL por USD.
                last = max(vals, key=lambda x: str(x.get("dataHoraCotacao", "")))
                return float(last["cotacaoVenda"]), probe
        except Exception:
            pass
        probe = previous_business_day(probe)
    raise RuntimeError(f"Não foi possível obter a PTAX para {d:%d/%m/%Y} nem dias anteriores.")


def ptax_with_manual(d: date, manual: dict[date, float]) -> tuple[float, date, str]:
    if d in manual:
        return float(manual[d]), d, "Manual"
    value, used_date = fetch_ptax_usd(d)
    return value, used_date, "BCB"


def read_control(file) -> tuple[pd.DataFrame, dict[str, list[tuple[float, Optional[float], float]]], list[str], pd.DataFrame]:
    """Lê a base principal e custódias auxiliares usadas apenas na segmentação.

    A aba ``Controle Contas`` contém as contas efetivamente cobradas. Abas auxiliares
    como ``Contas Atívore`` entram no PL do grupo familiar para definição da faixa,
    mas não geram fee nem exigem uma conta de cobrança.
    """
    xls = pd.ExcelFile(file)
    ctrl = pd.read_excel(xls, sheet_name="Controle Contas", dtype=object)
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
    ctrl = ctrl[ctrl["Corretora"].ne("") & ctrl["Cliente"].ne("")].copy()

    fee_raw = pd.read_excel(xls, sheet_name="Tabelas Fee", header=None, dtype=object)
    fee_tables = parse_fee_tables(fee_raw)
    pl_cols = [c for c in ctrl.columns if re.fullmatch(r"PL \d{2}/\d{2}/\d{4}", str(c))]

    aux_parts = []
    for sheet_name, broker_name in SEGMENTATION_AUX_SHEETS.items():
        if sheet_name not in xls.sheet_names:
            continue
        aux = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
        aux.columns = [norm_text(c) for c in aux.columns]
        if "Grupo Familiar" not in aux.columns or "Cliente" not in aux.columns:
            raise ValueError(
                f"A aba '{sheet_name}' precisa ter ao menos as colunas Grupo Familiar e Cliente."
            )
        if "Corretora" not in aux.columns:
            aux["Corretora"] = broker_name
        aux["Corretora"] = aux["Corretora"].map(lambda x: norm_text(x).upper() or broker_name)
        aux["Grupo Familiar"] = aux["Grupo Familiar"].map(norm_text)
        aux["Cliente"] = aux["Cliente"].map(norm_text)
        if "Moeda" not in aux.columns:
            aux["Moeda"] = "BRL"
        aux["Moeda"] = aux["Moeda"].map(lambda x: norm_text(x).upper() or "BRL")
        aux = aux[aux["Grupo Familiar"].ne("") | aux["Cliente"].ne("")].copy()
        aux_parts.append(aux)

    segmentation_aux = pd.concat(aux_parts, ignore_index=True, sort=False) if aux_parts else pd.DataFrame()
    return ctrl, fee_tables, pl_cols, segmentation_aux

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
    """Usa a última posição disponível DENTRO do mês selecionado para segmentação.

    Regra operacional: o mês calculado e a segmentação usam a mesma fotografia de
    fechamento. Ex.: cobrança 07/2026 -> PL de fechamento de 07/2026 e PTAX dessa
    mesma data. Enquanto o fechamento definitivo não estiver disponível, usa a
    posição mais recente do próprio mês como prévia.
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
        raise ValueError(
            f"Não encontrei coluna de PL no Controle para {month:02d}/{year}. "
            "A segmentação precisa usar uma posição do próprio mês calculado."
        )
    d, c = max(candidates)
    return c, d


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
    """Reconhece sem cobrança e taxas fixas descritas entre parênteses.

    Exemplos aceitos: ``Tesouraria (0,35)``, ``Tesouraria (0,40%)`` e
    ``Fee fixo (0,50%)``. Rótulos como ``Tabela 1`` não são interpretados
    como taxa fixa.
    """
    s = norm_text(label).upper()
    if s in NO_FEE_LABELS or "SEM COBR" in s:
        return 0.0
    m = re.search(r"\((\d+(?:[,.]\d+)?)\s*%?\)", s)
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
                       manual_ptax: dict[date, float], segmentation_aux: Optional[pd.DataFrame] = None
                       ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Monta cadastro de fee e PL consolidado para segmentação.

    Custódias auxiliares (hoje Atívore) entram exclusivamente no PL total do grupo.
    Elas não criam uma linha de cobrança, mas podem alterar a faixa regressiva.
    """
    df = ctrl.copy()
    df["PL_ref_original"] = pd.to_numeric(df[ref_col], errors="coerce").fillna(0.0)

    ref_ptax, ref_ptax_date, ref_ptax_src = ptax_with_manual(ref_date, manual_ptax)
    df["Moeda_ref"] = df["Corretora"].apply(lambda x: "USD" if x in OFFSHORE_BROKERS else "BRL")
    df["PTAX_ref"] = df["Moeda_ref"].apply(lambda x: ref_ptax if x == "USD" else 1.0)
    df["PL_ref_BRL"] = df["PL_ref_original"] * df["PTAX_ref"]
    df["Grupo_chave"] = df["Grupo Familiar"].where(df["Grupo Familiar"].ne(""), df["Cliente"])

    core_group = df.groupby("Grupo_chave", dropna=False)["PL_ref_BRL"].sum()

    aux_group = pd.Series(dtype="float64", name="PL_Aux_Grupo_BRL")
    ativore_group = pd.Series(dtype="float64", name="PL_Ativore_Grupo_BRL")
    if segmentation_aux is not None and not segmentation_aux.empty:
        if ref_col not in segmentation_aux.columns:
            raise ValueError(
                f"A aba de custódias auxiliares não possui '{ref_col}'. "
                "Atualize a Atívore para o mesmo fechamento usado na segmentação."
            )
        aux = segmentation_aux.copy()
        aux["PL_aux_original"] = pd.to_numeric(aux[ref_col], errors="coerce").fillna(0.0)
        aux["Grupo_chave"] = aux["Grupo Familiar"].where(aux["Grupo Familiar"].ne(""), aux["Cliente"])
        aux["Moeda"] = aux.get("Moeda", "BRL").map(lambda x: norm_text(x).upper() or "BRL")
        aux["PTAX_aux"] = aux["Moeda"].apply(lambda x: ref_ptax if x == "USD" else 1.0)
        aux["PL_aux_BRL"] = aux["PL_aux_original"] * aux["PTAX_aux"]
        aux_group = aux.groupby("Grupo_chave", dropna=False)["PL_aux_BRL"].sum().rename("PL_Aux_Grupo_BRL")
        ativore_group = (
            aux[aux["Corretora"].eq("ATIVORE")]
            .groupby("Grupo_chave", dropna=False)["PL_aux_BRL"].sum()
            .rename("PL_Ativore_Grupo_BRL")
        )

    all_groups = core_group.index.union(aux_group.index)
    total_group = core_group.reindex(all_groups, fill_value=0.0).add(
        aux_group.reindex(all_groups, fill_value=0.0), fill_value=0.0
    ).rename("PL_Grupo_Ref_BRL")

    df["PL_Grupo_Ref_BRL"] = df["Grupo_chave"].map(total_group).fillna(0.0)
    df["PL_Ativore_Grupo_BRL"] = df["Grupo_chave"].map(ativore_group).fillna(0.0)
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

    fee_group = df.groupby("Grupo_chave", dropna=False)["Fee_Projetado_Anual_Ref"].sum()
    count_group = df.groupby("Grupo_chave", dropna=False)["Conta_norm"].count()
    group_view = pd.DataFrame({"Grupo_chave": all_groups})
    group_view["PL_Cobrado_Ref_BRL"] = group_view["Grupo_chave"].map(core_group).fillna(0.0)
    group_view["PL_Ativore_Grupo_BRL"] = group_view["Grupo_chave"].map(ativore_group).fillna(0.0)
    group_view["PL_Grupo_Ref_BRL"] = group_view["Grupo_chave"].map(total_group).fillna(0.0)
    group_view["Contas"] = group_view["Grupo_chave"].map(count_group).fillna(0).astype(int)
    group_view["Fee_Projetado_Anual"] = group_view["Grupo_chave"].map(fee_group).fillna(0.0)
    group_view["ROA_Contratado"] = group_view["Fee_Projetado_Anual"] / group_view["PL_Grupo_Ref_BRL"].replace(0, pd.NA)
    group_view["Data_Referencia"] = pd.Timestamp(ref_date)
    group_view["PTAX_Referencia"] = ref_ptax
    group_view["Data_PTAX_Utilizada"] = pd.Timestamp(ref_ptax_date)
    group_view["Fonte_PTAX"] = ref_ptax_src

    keep = [
        "Grupo_chave", "Grupo Familiar", "Cliente", "Corretora", "Conta_norm", "Tabela Fee",
        "PL_ref_original", "Moeda_ref", "PTAX_ref", "PL_ref_BRL", "PL_Ativore_Grupo_BRL",
        "PL_Grupo_Ref_BRL", "Participacao_no_Grupo", "Fee_aa", "Regra_Fee",
        "Fee_Projetado_Mensal_Ref"
    ]
    return df[keep].copy(), group_view.sort_values("Grupo_chave").reset_index(drop=True)

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
    df["Metodo"] = "Diário / 250"
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
        df["Metodo"] = "Diário / 250"
        parts.append(df[["Data", "Corretora", "Conta_norm", "PL_original", "Moeda", "PTAX", "PL_BRL", "Metodo"]])
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def read_cs_daily(files: Iterable, year: int, month: int, manual_ptax: dict[date, float], reference_date: Optional[date] = None) -> pd.DataFrame:
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
    # Toda a posição offshore do mês usa a PTAX da mesma data de referência
    # utilizada na segmentação. Isso evita mudança de faixa causada por usar
    # câmbios diferentes entre segmentação e cobrança.
    close_date = reference_date or out["Data"].max().date()
    ptax, ptax_date, src = ptax_with_manual(close_date, manual_ptax)
    out["PTAX"] = ptax
    out["Data_PTAX"] = pd.Timestamp(ptax_date)
    out["Fonte_PTAX"] = src
    out["PL_BRL"] = out["PL_original"] * ptax
    out["Metodo"] = "Diário / 250"
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
                               year: int, month: int, manual_ptax: dict[date, float],
                               reference_col: Optional[str] = None,
                               reference_date: Optional[date] = None) -> tuple[pd.DataFrame, dict]:
    """Gera cobrança mensal de Safra, XP US e XP Global usando o fechamento da segmentação."""
    if reference_col is not None and reference_date is not None:
        billing_col, billing_date = reference_col, reference_date
        month_end = (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(1)).date()
        last_business = pd.bdate_range(pd.Timestamp(year, month, 1), pd.Timestamp(month_end))[-1].date()
        is_close = billing_date >= last_business
    else:
        billing_col, billing_date, is_close = choose_billing_pl_column(pl_cols, year, month)
    base = ctrl[ctrl["Corretora"].isin(MONTHLY_FROM_CONTROL_BROKERS)].copy()
    base["PL_original"] = pd.to_numeric(base[billing_col], errors="coerce").fillna(0.0)
    base["Data"] = pd.Timestamp(billing_date)
    base["Moeda"] = base["Corretora"].apply(lambda x: "USD" if x in {"XP US", "XP GLOBAL"} else "BRL")

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
    daily = out["Metodo"].eq("Diário / 250")
    monthly = out["Metodo"].str.startswith("Mensal / 12", na=False)
    out.loc[daily, "Fee_Calculado"] = out.loc[daily, "PL_BRL"] * out.loc[daily, "Fee_aa"].fillna(0) / DAILY_BASIS
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


def _xp_global_target_map(registry: pd.DataFrame) -> dict[tuple[str, str], str]:
    """Mapeia XP Global para a conta XP Brasil onde a cobrança será lançada.

    Prioridade: (1) mesmo Cliente; (2) se não houver match nominal, usa a conta XP
    do mesmo grupo familiar somente quando existir uma única candidata. Quando houver
    ambiguidade, não força o vínculo e deixa a linha para revisão.
    """
    mapping: dict[tuple[str, str], str] = {}
    global_rows = registry[registry["Corretora"].eq("XP GLOBAL")].copy()
    for _, row in global_rows.iterrows():
        client = norm_text(row.get("Cliente", ""))
        group = norm_text(row.get("Grupo_chave", ""))
        source_account = norm_account(row.get("Conta_norm", ""))
        local = registry[registry["Corretora"].eq("XP")].copy()
        local = local[local["Conta_norm"].astype(str).ne("")]

        candidates = local[local["Cliente"].eq(client)].copy()
        if candidates.empty:
            group_candidates = local[local["Grupo_chave"].eq(group)].copy()
            if group_candidates["Conta_norm"].nunique() == 1:
                candidates = group_candidates
        if candidates.empty:
            continue

        candidates["_pl"] = pd.to_numeric(candidates["PL_ref_BRL"], errors="coerce").fillna(0.0)
        target = candidates.sort_values(["_pl", "Conta_norm"], ascending=[False, True]).iloc[0]["Conta_norm"]
        mapping[(client, source_account)] = str(target)
    return mapping

def annotate_billing_destination(detail: pd.DataFrame, registry: pd.DataFrame) -> pd.DataFrame:
    """Acrescenta o destino operacional da cobrança sem apagar a origem do cálculo."""
    if detail.empty:
        return detail.copy()
    out = detail.copy()
    out["Corretora_Cobranca"] = out["Corretora"]
    out["Conta_Cobranca"] = out["Conta_norm"].astype(str)
    out["XP_Global_Incorporado"] = False
    out["Observacao_Cobranca"] = ""

    mapping = _xp_global_target_map(registry)
    mask = out["Corretora"].eq("XP GLOBAL")
    for i in out.index[mask]:
        client = norm_text(out.at[i, "Cliente"])
        source_account = norm_account(out.at[i, "Conta_norm"])
        target = mapping.get((client, source_account))
        if target:
            out.at[i, "Corretora_Cobranca"] = "XP"
            out.at[i, "Conta_Cobranca"] = target
            out.at[i, "XP_Global_Incorporado"] = True
            out.at[i, "Observacao_Cobranca"] = f"Fee XP Global ({source_account}) incorporado à conta XP Brasil {target}."
        else:
            out.at[i, "Observacao_Cobranca"] = "XP Global sem conta XP Brasil correspondente no Controle; revisar antes da cobrança."
    return out


def build_billing_summary(detail: pd.DataFrame, registry: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cria o resumo operacional final, incorporando XP Global à conta XP Brasil."""
    annotated = annotate_billing_destination(detail, registry)
    if annotated.empty:
        return annotated, pd.DataFrame()

    # Primeiro resume cada conta de origem; assim PL diário vira PL médio e não soma dias.
    source = (
        annotated.groupby([
            "Corretora", "Conta_norm", "Corretora_Cobranca", "Conta_Cobranca",
            "Grupo_chave", "Cliente", "Fee_aa", "Regra_Fee", "Moeda",
            "XP_Global_Incorporado", "Observacao_Cobranca"
        ], dropna=False)
        .agg(
            PL_Base_Original=("PL_original", "mean"),
            PL_Base_BRL=("PL_BRL", "mean"),
            Fee_Origem=("Fee_Calculado", "sum"),
        )
        .reset_index()
    )
    source["PL_XP_Global_USD"] = source["PL_Base_Original"].where(source["Corretora"].eq("XP GLOBAL"), 0.0)
    source["PL_XP_Global_BRL"] = source["PL_Base_BRL"].where(source["Corretora"].eq("XP GLOBAL"), 0.0)
    source["PL_Local_BRL"] = source["PL_Base_BRL"].where(~source["Corretora"].eq("XP GLOBAL"), 0.0)
    source["Fee_XP_Global"] = source["Fee_Origem"].where(source["Corretora"].eq("XP GLOBAL"), 0.0)
    source["Fee_Local"] = source["Fee_Origem"].where(~source["Corretora"].eq("XP GLOBAL"), 0.0)
    source["Conta_XP_Global"] = source["Conta_norm"].where(source["Corretora"].eq("XP GLOBAL"), "")

    rows = []
    keys = ["Corretora_Cobranca", "Conta_Cobranca", "Grupo_chave", "Cliente"]
    for key, g in source.groupby(keys, dropna=False, sort=True):
        rates = [float(x) for x in pd.to_numeric(g["Fee_aa"], errors="coerce").dropna().unique()]
        rate = rates[0] if rates else None
        global_accounts = sorted({str(x) for x in g["Conta_XP_Global"] if norm_text(x)})
        fee_global = float(g["Fee_XP_Global"].sum())
        notes = []
        if fee_global:
            notes.append("Fee XP Global incorporado à cobrança XP Brasil")
        if any("sem conta XP Brasil" in str(x).lower() for x in g["Observacao_Cobranca"]):
            notes.append("Revisar vínculo da XP Global")
        if len(rates) > 1:
            notes.append("Mais de uma taxa na composição")
        rows.append({
            "Corretora_Cobranca": key[0],
            "Conta_Cobranca": key[1],
            "Grupo_chave": key[2],
            "Cliente": key[3],
            "Fee_aa": rate,
            "PL_Local_BRL": float(g["PL_Local_BRL"].sum()),
            "PL_XP_Global_USD": float(g["PL_XP_Global_USD"].sum()),
            "PL_XP_Global_BRL": float(g["PL_XP_Global_BRL"].sum()),
            "Fee_Local": float(g["Fee_Local"].sum()),
            "Fee_XP_Global": fee_global,
            "Fee_Total": float(g["Fee_Origem"].sum()),
            "Contas_XP_Global": ", ".join(global_accounts),
            "Observacao": " · ".join(notes),
        })
    final = pd.DataFrame(rows)
    if not final.empty:
        value_cols = ["PL_Local_BRL", "PL_XP_Global_BRL", "Fee_Total"]
        keep_mask = pd.Series(False, index=final.index)
        for col in value_cols:
            keep_mask |= pd.to_numeric(final[col], errors="coerce").fillna(0).abs().gt(1e-9)
        final = final[keep_mask].copy()
        final = final.sort_values(["Corretora_Cobranca", "Cliente", "Conta_Cobranca"]).reset_index(drop=True)
    return annotated, final


def billing_summary_display(billing: pd.DataFrame) -> pd.DataFrame:
    if billing.empty:
        return billing
    d = billing.copy()
    out = pd.DataFrame(index=d.index)
    out["Grupo Familiar"] = d["Grupo_chave"]
    out["Cliente"] = d["Cliente"]
    out["Corretora"] = d["Corretora_Cobranca"]
    out["Conta para Cobrança"] = d["Conta_Cobranca"]
    out["Fee a.a."] = d["Fee_aa"].map(lambda x: fmt_pct(x, 3))
    out["PL Base Local"] = d["PL_Local_BRL"].map(fmt_brl)
    out["PL XP Global USD"] = d["PL_XP_Global_USD"].map(lambda x: fmt_usd(x) if float(x or 0) != 0 else "—")
    out["PL XP Global BRL"] = d["PL_XP_Global_BRL"].map(lambda x: fmt_brl(x) if float(x or 0) != 0 else "—")
    out["Fee Local"] = d["Fee_Local"].map(fmt_brl)
    out["Fee XP Global"] = d["Fee_XP_Global"].map(lambda x: fmt_brl(x) if float(x or 0) != 0 else "—")
    out["Fee Total"] = d["Fee_Total"].map(fmt_brl)
    out["Observação"] = d["Observacao"]
    return out.reset_index(drop=True)


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


def _group_label(df: pd.DataFrame) -> pd.Series:
    """Retorna uma única série de grupo sem criar nomes de coluna duplicados."""
    if "Grupo_chave" in df.columns:
        return df["Grupo_chave"]
    if "Grupo Familiar" in df.columns:
        return df["Grupo Familiar"]
    return pd.Series("", index=df.index, dtype="object")


def registry_display(registry: pd.DataFrame) -> pd.DataFrame:
    d = registry.copy()
    out = pd.DataFrame(index=d.index)
    out["Grupo Familiar"] = _group_label(d)
    out["Cliente"] = d.get("Cliente", "")
    out["Corretora"] = d.get("Corretora", "")
    out["Conta"] = d.get("Conta_norm", d.get("Conta", ""))
    out["Tabela"] = d.get("Tabela Fee", d.get("Tabela", ""))
    out["Regra do Fee"] = d.get("Regra_Fee", d.get("Regra do Fee", ""))
    out["PL Referência USD"] = d["PL_ref_original"].where(d["Moeda_ref"].eq("USD")).map(fmt_usd)
    out["PL Referência BRL"] = d["PL_ref_BRL"].map(fmt_brl)
    if "PL_Ativore_Grupo_BRL" in d.columns and pd.to_numeric(d["PL_Ativore_Grupo_BRL"], errors="coerce").fillna(0).ne(0).any():
        out["PL Atívore no Grupo"] = d["PL_Ativore_Grupo_BRL"].map(fmt_brl)
    out["PL Grupo Familiar"] = d["PL_Grupo_Ref_BRL"].map(fmt_brl)
    out["Participação"] = d["Participacao_no_Grupo"].map(fmt_pct)
    out["Fee a.a."] = d["Fee_aa"].map(lambda x: fmt_pct(x, 3))
    out["Fee Mensal Teórico"] = d["Fee_Projetado_Mensal_Ref"].map(fmt_brl)
    return out.reset_index(drop=True)


def summary_display(summary: pd.DataFrame) -> pd.DataFrame:
    d = summary.copy()
    out = pd.DataFrame(index=d.index)
    out["Grupo Familiar"] = _group_label(d)
    out["Cliente"] = d.get("Cliente", "")
    out["Corretora"] = d.get("Corretora", "")
    out["Conta"] = d.get("Conta_norm", d.get("Conta", ""))
    out["Tabela"] = d.get("Tabela Fee", d.get("Tabela", ""))
    out["Regra do Fee"] = d.get("Regra_Fee", d.get("Regra do Fee", ""))
    out["Metodologia"] = d.get("Metodo", d.get("Metodologia", ""))
    out["PL Médio USD"] = d["PL_Medio_USD"].map(fmt_usd)
    out["PL Médio BRL"] = d["PL_Medio_BRL"].map(fmt_brl)
    out["PL Fechamento USD"] = d["PL_Fechamento_USD"].map(fmt_usd)
    out["PL Fechamento BRL"] = d["PL_Fechamento_BRL"].map(fmt_brl)
    out["Fee a.a."] = d["Fee_aa"].map(lambda x: fmt_pct(x, 3))
    out["Dias/Registros"] = d.get("Dias_ou_Registros", d.get("Dias/Registros", ""))
    out["Fee do Mês"] = d["Fee_Mes"].map(fmt_brl)
    out["Período"] = pd.to_datetime(d["Data_Inicial"]).dt.strftime("%d/%m/%Y") + " a " + pd.to_datetime(d["Data_Final"]).dt.strftime("%d/%m/%Y")
    return out.reset_index(drop=True)


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


def _segmentacao_export(registry: pd.DataFrame) -> pd.DataFrame:
    reg = registry.copy()
    out = pd.DataFrame()
    out["Grupo Familiar"] = _group_label(reg)
    out["Cliente"] = reg["Cliente"]
    out["Corretora"] = reg["Corretora"]
    out["Conta"] = reg["Conta_norm"]
    out["Tabela Fee"] = reg["Tabela Fee"]
    out["Regra do Fee"] = reg["Regra_Fee"]
    out["PL Referência USD"] = reg["PL_ref_original"].where(reg["Moeda_ref"].eq("USD"))
    out["PTAX Referência"] = reg["PTAX_ref"].where(reg["Moeda_ref"].eq("USD"))
    out["PL Referência BRL"] = reg["PL_ref_BRL"]
    if "PL_Ativore_Grupo_BRL" in reg.columns:
        out["PL Atívore Grupo BRL"] = reg["PL_Ativore_Grupo_BRL"]
    out["PL Grupo Familiar BRL"] = reg["PL_Grupo_Ref_BRL"]
    out["Participação no Grupo"] = reg["Participacao_no_Grupo"]
    out["Fee a.a."] = reg["Fee_aa"]
    return out


def _groups_export(groups: pd.DataFrame) -> pd.DataFrame:
    return groups.copy().rename(columns={
        "Grupo_chave": "Grupo Familiar",
        "PL_Cobrado_Ref_BRL": "PL Contas Cobradas BRL",
        "PL_Ativore_Grupo_BRL": "PL Atívore BRL",
        "PL_Grupo_Ref_BRL": "PL Grupo Familiar BRL",
        "Fee_Projetado_Anual": "Fee Projetado Anual",
        "ROA_Contratado": "ROA Contratado",
        "Data_Referencia": "Data de Referência",
        "PTAX_Referencia": "PTAX Referência",
        "Data_PTAX_Utilizada": "Data PTAX Utilizada",
        "Fonte_PTAX": "Fonte PTAX",
    })


def _billing_export(billing: pd.DataFrame) -> pd.DataFrame:
    if billing.empty:
        return billing.copy()
    return billing.copy().rename(columns={
        "Grupo_chave": "Grupo Familiar",
        "Corretora_Cobranca": "Corretora Cobrança",
        "Conta_Cobranca": "Conta para Cobrança",
        "Fee_aa": "Fee a.a.",
        "PL_Local_BRL": "PL Base Local BRL",
        "PL_XP_Global_USD": "PL XP Global USD",
        "PL_XP_Global_BRL": "PL XP Global BRL",
        "Fee_Local": "Fee Local",
        "Fee_XP_Global": "Fee XP Global",
        "Fee_Total": "Fee Total Cobrança",
        "Contas_XP_Global": "Conta(s) XP Global",
        "Observacao": "Observação",
    })


def _detail_export(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return detail.copy()
    det = detail.copy().rename(columns={
        "Conta_norm": "Conta Origem",
        "PL_original": "PL Original",
        "PL_BRL": "PL BRL",
        "Metodo": "Metodologia",
        "Grupo_chave": "Grupo Familiar",
        "Tabela Fee": "Tabela Fee",
        "PL_ref_BRL": "PL Referência BRL",
        "PL_Grupo_Ref_BRL": "PL Grupo Familiar BRL",
        "Participacao_no_Grupo": "Participação no Grupo",
        "Fee_aa": "Fee a.a.",
        "Regra_Fee": "Regra do Fee",
        "Fee_Calculado": "Fee Calculado",
        "Corretora_Cobranca": "Corretora Cobrança",
        "Conta_Cobranca": "Conta para Cobrança",
        "XP_Global_Incorporado": "XP Global Incorporado",
        "Observacao_Cobranca": "Observação Cobrança",
    })
    if "Moeda" in det.columns:
        det["PL USD"] = det["PL Original"].where(det["Moeda"].eq("USD"))
    preferred = [
        "Data", "Grupo Familiar", "Cliente", "Corretora", "Conta Origem", "Moeda",
        "PL Original", "PL USD", "PTAX", "PL BRL", "Tabela Fee", "Regra do Fee",
        "Fee a.a.", "Fee Calculado", "Corretora Cobrança", "Conta para Cobrança",
        "XP Global Incorporado", "Observação Cobrança", "Metodologia"
    ]
    cols = [c for c in preferred if c in det.columns]
    return det[cols]


def export_excel(registry, groups, detail, billing, year, month, ref_date) -> bytes:
    output = io.BytesIO()
    seg = _segmentacao_export(registry)
    grp = _groups_export(groups)
    bill = _billing_export(billing)
    det = _detail_export(detail)

    ptax_ref = None
    if not groups.empty and "PTAX_Referencia" in groups.columns:
        vals = pd.to_numeric(groups["PTAX_Referencia"], errors="coerce").dropna()
        if not vals.empty:
            ptax_ref = float(vals.iloc[0])

    readme = pd.DataFrame([
        ["Período calculado", f"{month:02d}/{year}"],
        ["Fechamento usado na segmentação", ref_date.strftime("%d/%m/%Y")],
        ["PTAX de referência", f"{ptax_ref:.4f}" if ptax_ref is not None else "—"],
        ["", ""],
        ["Resumo Cobranca", "Resumo operacional para cobrança. A XP Global é incorporada à conta XP Brasil do mesmo cliente e aparece destacada em colunas próprias."],
        ["Segmentacao", "Fotografia usada para definir a faixa de fee. Inclui custódias auxiliares como Atívore no PL do grupo, sem gerar cobrança própria."],
        ["Grupos", "Consolidado por grupo familiar, separando PL das contas cobradas e PL Atívore usado somente para segmentação."],
        ["Memoria Calculo", "Trilha completa do cálculo por data/conta. Mantém a XP Global como origem separada e informa em qual conta XP Brasil o fee foi incorporado."],
    ], columns=["Item / Aba", "Descrição"])

    datasets = {
        "LEIA-ME": readme,
        "Resumo Cobranca": bill,
        "Segmentacao": seg,
        "Grupos": grp,
        "Memoria Calculo": det,
    }

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        for sname, data in datasets.items():
            data.to_excel(writer, sheet_name=sname, index=False)

        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": WEALTH_NAVY, "font_color": "#FFFFFF", "border": 0, "align": "center", "valign": "vcenter"})
        fmt_brl_x = wb.add_format({"num_format": 'R$ #,##0.00;[Red](R$ #,##0.00);-'})
        fmt_usd_x = wb.add_format({"num_format": '$ #,##0.00;[Red]($ #,##0.00);-'})
        fmt_pct_x = wb.add_format({"num_format": '0.000%;[Red](0.000%);-'})
        fmt_ptax = wb.add_format({"num_format": '0.0000'})
        fmt_int = wb.add_format({"num_format": '#,##0'})
        fmt_wrap = wb.add_format({"text_wrap": True, "valign": "top"})

        for sname, data in datasets.items():
            ws = writer.sheets[sname]
            tab_colors = {"LEIA-ME": WEALTH_SAND, "Resumo Cobranca": WEALTH_BLUE, "Segmentacao": WEALTH_BLUE, "Grupos": WEALTH_SAND, "Memoria Calculo": WEALTH_NAVY}
            ws.set_tab_color(tab_colors.get(sname, WEALTH_BLUE))
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_row(0, 24, fmt_header)
            if not data.empty:
                ws.autofilter(0, 0, len(data), len(data.columns)-1)
            for i, col in enumerate(data.columns):
                name = str(col)
                width = 16
                fmt = None
                if sname == "LEIA-ME":
                    width = 30 if i == 0 else 95
                    fmt = fmt_wrap
                elif "Cliente" in name or "Grupo Familiar" in name or "Observação" in name:
                    width = 30
                elif "Corretora" in name or "Metodologia" in name or "Regra" in name or "Tabela" in name:
                    width = 20
                elif "Data" in name:
                    width = 15
                elif "Conta" in name:
                    width = 18
                if "USD" in name:
                    fmt = fmt_usd_x
                if "BRL" in name or name.startswith("Fee ") or name in {"Fee Total Cobrança", "Fee Calculado", "Fee Projetado Anual"}:
                    fmt = fmt_brl_x
                if name in {"Fee a.a.", "ROA Contratado", "Participação no Grupo"}:
                    fmt = fmt_pct_x
                if "PTAX" in name:
                    fmt = fmt_ptax
                if name in {"Contas"}:
                    fmt = fmt_int
                ws.set_column(i, i, width, fmt)

        # Destaque visual na aba operacional.
        if "Resumo Cobranca" in writer.sheets and not bill.empty:
            ws = writer.sheets["Resumo Cobranca"]
            xp_col = bill.columns.get_loc("Fee XP Global") if "Fee XP Global" in bill.columns else None
            if xp_col is not None:
                highlight = wb.add_format({"bg_color": "#FFF2CC", "num_format": 'R$ #,##0.00;[Red](R$ #,##0.00);-'})
                ws.conditional_format(1, xp_col, len(bill), xp_col, {"type": "cell", "criteria": ">", "value": 0, "format": highlight})
    return output.getvalue()


# ----------------------------- UI -----------------------------
icon_path = find_icon_path()
page_icon = Image.open(icon_path) if icon_path else "M"
st.set_page_config(page_title="Calculadora Fee | M Wealth", page_icon=page_icon, layout="wide")

st.markdown(f"""
<style>
    :root {{
        --mw-navy: {WEALTH_NAVY};
        --mw-navy-2: {WEALTH_NAVY_2};
        --mw-sand: {WEALTH_SAND};
        --mw-blue: {WEALTH_BLUE};
        --mw-text: #F7F8FA;
        --mw-muted: #9EA8B3;
        --mw-border: rgba(220,201,177,.16);
    }}
    .stApp {{ background: #0C1218; color: var(--mw-text); }}
    [data-testid="stAppViewContainer"] {{
        background:
          radial-gradient(circle at 86% 0%, rgba(93,115,175,.16), transparent 29rem),
          linear-gradient(180deg, #0C1218 0%, #0F161E 100%);
    }}
    [data-testid="stHeader"] {{
        background: rgba(12,18,24,.96) !important;
        border-bottom: 1px solid rgba(220,201,177,.08);
    }}
    /* Reserva a altura real do toolbar fixo do Streamlit.
       Evita que o cabeçalho da M Wealth fique por baixo da barra superior. */
    [data-testid="stAppViewContainer"] .main .block-container,
    section.main > div.block-container {{
        padding-top: 4.75rem !important;
        padding-bottom: 4rem !important;
        max-width: 1480px !important;
    }}

    .mw-header {{
        display:flex; align-items:flex-end; justify-content:space-between; gap:2rem;
        padding: .15rem 0 1.15rem 0; margin-bottom:.55rem;
        border-bottom:1px solid var(--mw-border);
    }}
    .mw-brand-row {{display:flex; align-items:center; gap:.65rem; margin-bottom:.45rem;}}
    .mw-mark {{
        width:31px; height:31px; border-radius:9px; display:inline-flex; align-items:center; justify-content:center;
        background:var(--mw-blue); color:white; font-weight:800; font-size:1rem; box-shadow:0 7px 20px rgba(93,115,175,.22);
    }}
    .mw-brand {{font-size:.72rem; letter-spacing:.17em; font-weight:650; color:var(--mw-sand); text-transform:uppercase;}}
    .mw-title {{font-size:2.2rem; line-height:1.05; font-weight:720; letter-spacing:-.045em; color:white; margin:0;}}
    .mw-sub {{font-size:.88rem; color:var(--mw-muted); margin:.45rem 0 0 0; max-width:760px;}}
    .mw-version {{font-size:.71rem; color:var(--mw-muted); white-space:nowrap; padding-bottom:.1rem;}}

    .section-label {{font-size:.68rem; letter-spacing:.15em; text-transform:uppercase; color:var(--mw-sand); margin-top:1.25rem; font-weight:650;}}
    .section-title {{font-size:1.25rem; color:#F7F8FA; font-weight:680; letter-spacing:-.025em; margin:.12rem 0 .65rem 0;}}
    .small-note {{font-size:.78rem; color:var(--mw-muted); margin-top:.4rem;}}
    .status-row {{display:flex; flex-wrap:wrap; gap:.45rem; margin:.25rem 0 1rem 0;}}
    .status-pill {{
        display:inline-flex; align-items:center; gap:.35rem; padding:.34rem .62rem; border-radius:999px;
        border:1px solid rgba(220,201,177,.13); background:rgba(255,255,255,.025);
        color:#C8D0D8; font-size:.73rem;
    }}
    .status-dot {{width:6px; height:6px; border-radius:50%; background:var(--mw-blue); display:inline-block;}}

    div[data-testid="stMetric"] {{
        background: linear-gradient(180deg, rgba(20,28,36,.96), rgba(19,25,37,.90));
        border:1px solid rgba(220,201,177,.13);
        border-top:2px solid rgba(93,115,175,.78);
        padding:15px 17px; border-radius:13px; min-height:104px;
        box-shadow:0 10px 28px rgba(0,0,0,.14);
    }}
    div[data-testid="stMetricLabel"] {{font-size:.74rem; color:#AEB7C0; letter-spacing:.01em;}}
    div[data-testid="stMetricValue"] {{font-size:1.34rem; color:#FFFFFF; font-weight:680;}}

    [data-testid="stDataFrame"] {{border:1px solid rgba(220,201,177,.12); border-radius:12px; overflow:hidden;}}
    [data-testid="stFileUploader"] {{
        border:1px dashed rgba(220,201,177,.24); border-radius:12px; background:rgba(20,28,36,.44);
    }}
    [data-testid="stFileUploader"] section {{padding-top:.7rem; padding-bottom:.7rem;}}
    div[data-testid="stExpander"] {{border-radius:12px; border:1px solid rgba(220,201,177,.12); background:rgba(20,28,36,.32);}}

    div.stButton > button[kind="primary"] {{
        background:var(--mw-blue) !important; color:white !important; border:1px solid rgba(255,255,255,.08) !important;
        border-radius:10px !important; font-weight:680 !important; min-height:2.8rem;
        box-shadow:0 9px 24px rgba(93,115,175,.22);
    }}
    div.stButton > button[kind="primary"]:hover {{background:#6B82C2 !important; border-color:var(--mw-sand) !important;}}
    div[data-testid="stDownloadButton"] > button {{
        background:var(--mw-sand) !important; color:var(--mw-navy) !important;
        border:1px solid var(--mw-sand) !important; border-radius:10px !important;
        font-weight:760 !important; min-height:3rem; box-shadow:0 10px 28px rgba(220,201,177,.16);
    }}
    div[data-testid="stDownloadButton"] > button:hover {{
        background:#E8D9C5 !important; color:#0C1218 !important; border-color:#FFFFFF !important; transform:translateY(-1px);
    }}
    .stSelectbox label, .stTextInput label, .stFileUploader label {{font-size:.78rem !important; color:#B7C0C8 !important;}}
    hr {{border-color:rgba(220,201,177,.10) !important;}}

    @media (max-width: 900px) {{
        [data-testid="stAppViewContainer"] .main .block-container,
        section.main > div.block-container {{ padding-top: 4.9rem !important; }}
    }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="mw-header">
  <div>
    <div class="mw-brand-row"><span class="mw-mark">M</span><span class="mw-brand">M Wealth · Operações</span></div>
    <div class="mw-title">Calculadora de Fee</div>
  </div>
  <div class="mw-version">Motor de cobrança · base 250</div>
</div>
""", unsafe_allow_html=True)

today = date.today()
month_names = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
sel1, sel2, sel3, spacer = st.columns([1.0, 1.15, 2.7, 5.15])
with sel1:
    year = st.selectbox("Ano", list(range(today.year-2, today.year+3)), index=2)
with sel2:
    month_name = st.selectbox("Mês", list(month_names.values()), index=today.month-1)
    month = next(k for k,v in month_names.items() if v == month_name)
with sel3:
    ptax_text = st.text_input("PTAX manual · opcional", placeholder="30/09/2026;5,1809")
manual_ptax = parse_manual_ptax(ptax_text)

control_path = find_fixed_control_path()
if control_path is None:
    st.error("Base fixa não encontrada. Mantenha 'Controle de Clientes MWealth 2026.xlsx' na raiz do projeto.")
    st.stop()

try:
    ctrl, fee_tables, pl_cols, segmentation_aux = read_control(control_path)
    ref_col, ref_date = choose_reference_pl_column(pl_cols, int(year), int(month))
    registry, groups = build_fee_registry(ctrl, fee_tables, ref_col, ref_date, manual_ptax, segmentation_aux)
except Exception as e:
    st.error(f"Erro ao carregar a base fixa: {e}")
    st.stop()

latest_pl_date = max(datetime.strptime(c[3:], "%d/%m/%Y").date() for c in pl_cols)
accounts_count = int(ctrl["Conta_norm"].ne("").sum())
groups_count = int(groups["Grupo_chave"].nunique())
total_pl = float(groups["PL_Grupo_Ref_BRL"].sum()) if not groups.empty else 0.0
ativore_pl = float(groups["PL_Ativore_Grupo_BRL"].sum()) if "PL_Ativore_Grupo_BRL" in groups.columns else 0.0
fee_ann = float(registry["Fee_Projetado_Mensal_Ref"].sum() * 12)
roa = fee_ann / total_pl if total_pl else 0
ptax_ref = float(groups["PTAX_Referencia"].dropna().iloc[0]) if not groups.empty else 0
ptax_source = str(groups["Fonte_PTAX"].dropna().iloc[0]) if not groups.empty else "—"

st.markdown(
    f'<div class="status-row">'
    f'<span class="status-pill"><span class="status-dot"></span>Fechamento {ref_date:%d/%m/%Y}</span>'
    f'<span class="status-pill"><span class="status-dot"></span>PTAX venda {ptax_ref:.4f} · {ptax_source}</span>'
    f'<span class="status-pill"><span class="status-dot"></span>{accounts_count} contas cobradas</span>'
    f'</div>', unsafe_allow_html=True
)

st.markdown('<div class="section-label">Visão geral</div><div class="section-title">Fotografia da segmentação</div>', unsafe_allow_html=True)
r1, r2, r3, r4, r5 = st.columns(5)
r1.metric("PL para segmentação", fmt_brl(total_pl))
r2.metric("Fee mensal teórico", fmt_brl(fee_ann/12))
r3.metric("ROA efetivo", fmt_pct(roa, 3))
r4.metric("Grupos familiares", f"{groups_count:,}".replace(",", "."))
r5.metric("PL Atívore incluído", fmt_brl(ativore_pl))
st.markdown(f'<div class="small-note">Último PL disponível no Controle: {latest_pl_date:%d/%m/%Y} · Atívore entra somente na segmentação e não gera cobrança própria.</div>', unsafe_allow_html=True)

with st.expander("Detalhar segmentação e taxas", expanded=False):
    st.dataframe(registry_display(registry), use_container_width=True, hide_index=True, height=420)

st.markdown('<div class="section-label">Entradas</div><div class="section-title">Arquivos operacionais do mês</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1:
    xp_file = st.file_uploader("XP · relatório com todos os dias", type=["xlsx"], key="xp")
with c2:
    btg_files = st.file_uploader("BTG · um arquivo por dia", type=["xlsx"], accept_multiple_files=True, key="btg")
with c3:
    cs_files = st.file_uploader("Charles Schwab · um CSV por dia", type=["csv"], accept_multiple_files=True, key="cs")

monthly_preview = None
try:
    monthly_preview, monthly_meta = build_monthly_from_control(
        registry, ctrl, pl_cols, int(year), int(month), manual_ptax, ref_col, ref_date
    )
except Exception:
    monthly_preview = None

st.markdown('<div style="height:.25rem"></div>', unsafe_allow_html=True)
if st.button("Calcular cobrança do mês", type="primary", use_container_width=True):
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
            parts.append(read_cs_daily(cs_files, int(year), int(month), manual_ptax, ref_date))
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
        coverage = build_coverage(mov[mov["Metodo"].eq("Diário / 250")].copy(), int(year), int(month))
        detail, summary = calculate_fees(mov, registry)
        detail_billing, billing = build_billing_summary(detail, registry)

        st.markdown('<div class="section-label">Fechamento</div><div class="section-title">Resultado da cobrança</div>', unsafe_allow_html=True)
        daily_detail = detail_billing[detail_billing["Metodo"].eq("Diário / 250")].copy() if not detail_billing.empty else pd.DataFrame()
        monthly_detail = detail_billing[detail_billing["Metodo"].str.startswith("Mensal / 12", na=False)].copy() if not detail_billing.empty else pd.DataFrame()
        xp_global_fee = detail_billing.loc[detail_billing["Corretora"].eq("XP GLOBAL"), "Fee_Calculado"].sum() if not detail_billing.empty else 0
        total_fee = detail_billing["Fee_Calculado"].sum() if not detail_billing.empty else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Fee total", fmt_brl(total_fee))
        m2.metric("Fee diário", fmt_brl(daily_detail["Fee_Calculado"].sum() if not daily_detail.empty else 0))
        m3.metric("Fee mensal", fmt_brl(monthly_detail["Fee_Calculado"].sum() if not monthly_detail.empty else 0))
        m4.metric("XP Global incorporado", fmt_brl(xp_global_fee))

        st.markdown("#### Cobrança consolidada")
        if billing.empty:
            st.info("Sem contas calculadas.")
        else:
            st.dataframe(billing_summary_display(billing), use_container_width=True, hide_index=True, height=440)

        missing_count = 0 if coverage.empty else int(coverage["Datas Ausentes"].ne("Nenhuma").sum())
        with st.expander(f"Validação operacional · {len(diagnostics)} pendência(s) de conta · {missing_count} fonte(s) com data ausente", expanded=False):
            st.markdown("**Cobertura dos arquivos diários**")
            if coverage.empty:
                st.info("Nenhuma fonte diária carregada.")
            else:
                st.dataframe(coverage_display(coverage), use_container_width=True, hide_index=True)
                if missing_count:
                    st.warning("Há datas úteis sem arquivo/posição. Confira a coluna 'Datas Ausentes' antes do fechamento.")

            st.markdown("**Conferência de contas**")
            st.caption("XP e BTG são confrontadas com o Controle. Contas extras da Charles Schwab são ignoradas automaticamente.")
            if diagnostics.empty:
                st.success("XP e BTG sem divergências entre contas carregadas e Controle.")
            else:
                st.dataframe(diagnostics_display(diagnostics), use_container_width=True, hide_index=True, height=330)

        excel = export_excel(registry, groups, detail_billing, billing, int(year), int(month), ref_date)
        st.markdown('<div class="section-label">Saída</div><div class="section-title">Relatório operacional</div>', unsafe_allow_html=True)
        st.download_button(
            "Baixar relatório completo em Excel", data=excel,
            file_name=f"Fee_MWealth_{int(year)}_{int(month):02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
