
import re
import unicodedata
from io import BytesIO
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st

from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# =========================================================
# Helpers de normalização e conversão
# =========================================================

def normalize_colname(name: str) -> str:
    if not isinstance(name, str):
        return ""
    s = name.strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def name_2p(s: str) -> str:
    """Primeiras 2 palavras (igual seu código do endereço)."""
    s = "" if s is None else str(s)
    s = re.sub(r"[\n\r\t]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    parts = s.split()
    if len(parts) >= 2:
        return " ".join(parts[:2]).strip()
    return s.strip()


CANONICAL_MAP = {
    "colaborador": "Colaborador",
    "cpf": "CPF",
    "data": "Data",
    "previstas": "Previstas",
    "trabalhadas": "Trabalhadas",
    "abonadas": "Abonadas",
    "atrasos": "Atrasos",
    "noturnas": "Noturnas",
    "hr. ficta": "Horas Fictas",
    "hr ficta": "Horas Fictas",
    "horas fictas": "Horas Fictas",
    "faltas": "Faltas",
}


def parse_time_to_timedelta(value):
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timedelta):
        return value

    s = str(value).strip()
    if s == "" or s in {"-", "—", "–", "--", "nan", "None"}:
        return pd.NaT

    if re.match(r"^\d{1,2}:\d{2}$", s):
        try:
            h, m = map(int, s.split(":"))
            return pd.Timedelta(hours=h, minutes=m)
        except Exception:
            return pd.NaT

    if re.match(r"^\d{1,2}:\d{2}:\d{2}$", s):
        try:
            h, m, sec = map(int, s.split(":"))
            return pd.Timedelta(hours=h, minutes=m, seconds=sec)
        except Exception:
            return pd.NaT

    td = pd.to_timedelta(s, errors="coerce")
    if not pd.isna(td):
        return td

    # float Excel (dias)
    try:
        f = float(s.replace(",", "."))
        return pd.to_timedelta(f, unit="D")
    except Exception:
        return pd.NaT


def fmt_td_hms(td: pd.Timedelta) -> str:
    """HH:MM:SS (sempre positivo)"""
    if pd.isna(td):
        return "00:00:00"
    total = int(td.total_seconds())
    if total < 0:
        total = 0
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def fmt_td_hms_signed(td: pd.Timedelta) -> str:
    """(+/-)HH:MM:SS"""
    if pd.isna(td):
        return "00:00:00"
    total = int(td.total_seconds())
    sign = "-" if total < 0 else ""
    total = abs(total)
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"


def fmt_dt_time(x) -> str:
    """HH:MM:SS ou vazio"""
    if x is None or pd.isna(x):
        return ""
    try:
        ts = pd.Timestamp(x)
        return ts.strftime("%H:%M:%S")
    except Exception:
        return ""


def month_name_pt(m: int) -> str:
    meses = [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ]
    return meses[m - 1] if 1 <= m <= 12 else ""


# =========================================================
# Etapa 1 — Detectar cabeçalhos e montar tabela longa (HE)
# =========================================================

def _build_long_from_folha_ponto(raw: pd.DataFrame) -> pd.DataFrame:
    mask_blocks = raw.apply(
        lambda r: r.astype(str).str.contains("DADOS DO COLABORADOR", na=False).any(),
        axis=1,
    )
    block_starts = [i for i in raw.index if mask_blocks[i]]
    if not block_starts:
        return pd.DataFrame()

    n_rows, _ = raw.shape
    rows = []

    def find_header_row(start, end):
        for i in range(start, end):
            if raw.iloc[i].astype(str).str.contains("DIA / MÊS", na=False).any():
                return i
        return None

    def find_col_idx(header_row, label):
        row = raw.iloc[header_row].astype(str)
        hits = np.where(row.str.contains(label, case=False, na=False))[0]
        return int(hits[0]) if hits.size else None

    def extract_value_after_label(block_start, block_end, label):
        for i in range(block_start, block_end):
            row = raw.iloc[i]
            srow = row.astype(str)
            for j, val in enumerate(srow):
                if label in str(val):
                    for k in range(j + 1, len(row)):
                        v = row[k]
                        if (not pd.isna(v)) and str(v).strip() != "":
                            return v
        return None

    for idx, start in enumerate(block_starts):
        end = block_starts[idx + 1] if idx + 1 < len(block_starts) else n_rows
        header_row = find_header_row(start, end)
        if header_row is None:
            continue

        col_data = find_col_idx(header_row, "DIA / MÊS")
        col_pontos = find_col_idx(header_row, "PONTOS")
        col_trab = find_col_idx(header_row, "TRABALHADAS")
        col_abono = find_col_idx(header_row, "ABONO")
        col_prev = find_col_idx(header_row, "PREVISTAS")

        nome_val = extract_value_after_label(start, header_row, "Nome:")
        cpf_val = extract_value_after_label(start, header_row, "CPF:")

        cpf_str = ""
        if cpf_val is not None:
            try:
                if isinstance(cpf_val, (int, float)):
                    cpf_str = str(int(cpf_val))
                else:
                    cpf_str = str(cpf_val).strip()
            except Exception:
                cpf_str = str(cpf_val).strip()

        r = header_row + 1
        while r < end:
            v_data = raw.iloc[r, col_data] if col_data is not None else None
            if isinstance(v_data, str) and v_data.strip().startswith("Total"):
                break
            if pd.isna(v_data):
                r += 1
                continue

            pontos_val = raw.iloc[r, col_pontos] if col_pontos is not None else None
            trab_val = raw.iloc[r, col_trab] if col_trab is not None else None
            abono_val = raw.iloc[r, col_abono] if col_abono is not None else None
            prev_val = raw.iloc[r, col_prev] if col_prev is not None else None

            faltas_val = None
            if isinstance(pontos_val, str) and "FALTA" in pontos_val.upper():
                faltas_val = prev_val

            rows.append(
                {
                    "Colaborador": str(nome_val).strip() if nome_val is not None else "",
                    "CPF": cpf_str,
                    "Data": v_data,
                    "Previstas": prev_val,
                    "Trabalhadas": trab_val,
                    "Abonadas": abono_val,
                    "Atrasos": None,
                    "Noturnas": None,
                    "Horas Fictas": None,
                    "Faltas": faltas_val,
                }
            )
            r += 1

    return pd.DataFrame(rows)


def _build_long_from_solides(raw: pd.DataFrame) -> pd.DataFrame:
    mask_previstas = raw.apply(
        lambda row: row.astype(str).str.contains("Previstas", case=False, na=False).any(),
        axis=1,
    )
    header_idx = [i for i in raw.index if mask_previstas[i]]
    if not header_idx:
        return pd.DataFrame()

    blocks = []
    for pos, h in enumerate(header_idx):
        next_h = header_idx[pos + 1] if pos + 1 < len(header_idx) else len(raw)
        header_row = raw.iloc[h]
        block = raw.iloc[h + 1: next_h].copy()
        block = block.dropna(how="all")
        if block.empty:
            continue
        block.columns = header_row.values
        blocks.append(block)

    if not blocks:
        return pd.DataFrame()

    return pd.concat(blocks, ignore_index=True)


def detect_blocks_and_build_long(raw: pd.DataFrame) -> pd.DataFrame:
    long_df = _build_long_from_folha_ponto(raw)
    if not long_df.empty:
        return long_df

    long_df = _build_long_from_solides(raw)
    if not long_df.empty:
        return long_df

    raise ValueError(
        "Não consegui identificar o layout do relatório.\n"
        "- Para Folha de Ponto, preciso da área 'DADOS DO COLABORADOR' com o quadro DIA / MÊS.\n"
        "- Para o modelo antigo, preciso das colunas com cabeçalho 'Previstas', 'Trabalhadas' etc."
    )


# =========================================================
# Etapa 2 — Semanas 26→25
# =========================================================

def classify_weeks(dates: pd.Series) -> pd.Series:
    if dates.empty:
        return pd.Series(dtype="object")

    dts = pd.to_datetime(dates, errors="coerce")
    dts_norm = dts.dt.normalize()

    ym_list = sorted({(d.year, d.month) for d in dts_norm.dropna()})
    if not ym_list:
        return pd.Series(["Semana 1"] * len(dts), index=dates.index)

    prev_year, prev_month = ym_list[0]
    if len(ym_list) > 1:
        curr_year, curr_month = ym_list[1]
    else:
        curr_year, curr_month = prev_year, prev_month

    def _classify(d):
        if pd.isna(d):
            return None
        y, m, day = d.year, d.month, d.day

        if (y == prev_year) and (m == prev_month):
            return "Semana 1" if day >= 26 else "Semana 1"

        if (y == curr_year) and (m == curr_month):
            if 1 <= day <= 7:
                return "Semana 2"
            elif 8 <= day <= 14:
                return "Semana 3"
            elif 15 <= day <= 21:
                return "Semana 4"
            else:
                return "Semana 5"

        return "Semana 5"

    return dts_norm.apply(_classify)


# =========================================================
# Etapa 2.1 — Feriados nacionais e domingos (HE 100%)
# =========================================================

NATIONAL_HOLIDAYS = {
    (1, 1),
    (4, 21),
    (5, 1),
    (9, 7),
    (10, 12),
    (11, 2),
    (11, 15),
    (11, 20),
    (12, 25),
}

def is_national_holiday(date_value) -> bool:
    if pd.isna(date_value):
        return False
    d = pd.Timestamp(date_value)
    return (d.month, d.day) in NATIONAL_HOLIDAYS


# =========================================================
# Etapa 3 — Limpeza, tipos, Atrasos e HE 50/70/100/Abono
# =========================================================

def clean_and_enrich(long_df: pd.DataFrame) -> pd.DataFrame:
    df = long_df.dropna(axis=1, how="all").copy()

    rename_map = {}
    for col in df.columns:
        canon = CANONICAL_MAP.get(normalize_colname(col))
        if canon:
            rename_map[col] = canon
    df = df.rename(columns=rename_map)

    required = ["Colaborador", "CPF", "Data", "Previstas", "Trabalhadas"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(missing))

    for opt in ["Abonadas", "Atrasos", "Noturnas", "Horas Fictas", "Faltas"]:
        if opt not in df.columns:
            df[opt] = pd.NA

    df = df[
        [
            "Colaborador",
            "CPF",
            "Data",
            "Previstas",
            "Trabalhadas",
            "Abonadas",
            "Atrasos",
            "Noturnas",
            "Horas Fictas",
            "Faltas",
        ]
    ].copy()

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[~df["Data"].isna()].copy()

    df["Colaborador"] = df["Colaborador"].astype(str).str.strip()
    df["CPF"] = df["CPF"].astype(str).str.strip()

    # >>> CHAVE 2P (pra cruzar com endereço)
    df["Nome_2p"] = df["Colaborador"].apply(name_2p)

    time_cols = [
        "Previstas",
        "Trabalhadas",
        "Abonadas",
        "Atrasos",
        "Noturnas",
        "Horas Fictas",
        "Faltas",
    ]
    for col in time_cols:
        df[f"{col}_td"] = df[col].apply(parse_time_to_timedelta)
        df[f"{col}_td"] = df[f"{col}_td"].fillna(pd.Timedelta(0))

    zero = pd.Timedelta(0)

    # Atrasos: Trabalhadas - Previstas (se negativo)
    df["Atrasos_td"] = df["Trabalhadas_td"] - df["Previstas_td"]
    df["Atrasos_td"] = df["Atrasos_td"].where(df["Atrasos_td"] < zero, zero)

    # Excedente: positivo
    df["Excedente_td"] = df["Trabalhadas_td"] - df["Previstas_td"]
    df["Excedente_td"] = df["Excedente_td"].where(df["Excedente_td"] > zero, zero)

    df["is_domingo"] = df["Data"].dt.dayofweek == 6
    df["is_feriado"] = df["Data"].apply(is_national_holiday)
    mask_especial = df["is_domingo"] | df["is_feriado"]
    mask_normal = ~mask_especial

    df["HE50_td"] = zero
    df["HE70_td"] = zero
    df["HE100_td"] = zero

    two_hours = pd.to_timedelta(2, unit="h")

    # Normal: 50% até 2h, resto 70%
    df.loc[mask_normal, "HE50_td"] = df.loc[mask_normal, "Excedente_td"].clip(
        lower=zero, upper=two_hours
    )
    df.loc[mask_normal, "HE70_td"] = (
        df.loc[mask_normal, "Excedente_td"] - df.loc[mask_normal, "HE50_td"]
    ).clip(lower=zero)

    # Domingo/feriado: 100%
    df.loc[mask_especial, ["HE50_td", "HE70_td"]] = zero
    df.loc[mask_especial, "HE100_td"] = df.loc[mask_especial, "Excedente_td"]

    # Sem previstas e trabalhou (escala 0): 100%
    mask_sem_previstas = (
        (df["Previstas_td"] == zero)
        & (df["Trabalhadas_td"] > zero)
        & (~mask_especial)
    )
    df.loc[mask_sem_previstas, ["HE50_td", "HE70_td"]] = zero
    df.loc[mask_sem_previstas, "HE100_td"] = df.loc[mask_sem_previstas, "Excedente_td"]

    # Abono sem trabalho: zera tudo
    mask_abono_sem_trabalho = (df["Abonadas_td"] > zero) & (df["Trabalhadas_td"] == zero)
    df.loc[
        mask_abono_sem_trabalho,
        ["Atrasos_td", "Excedente_td", "HE50_td", "HE70_td", "HE100_td"],
    ] = zero

    # Abono com trabalho: tudo trabalhado vira 100%
    mask_abono_com_trabalho = (df["Abonadas_td"] > zero) & (df["Trabalhadas_td"] > zero)
    df.loc[mask_abono_com_trabalho, "Atrasos_td"] = zero
    df.loc[mask_abono_com_trabalho, "Excedente_td"] = df.loc[
        mask_abono_com_trabalho, "Trabalhadas_td"
    ]
    df.loc[mask_abono_com_trabalho, ["HE50_td", "HE70_td"]] = zero
    df.loc[mask_abono_com_trabalho, "HE100_td"] = df.loc[
        mask_abono_com_trabalho, "Excedente_td"
    ]

    df["Semana"] = classify_weeks(df["Data"])
    df = df.sort_values(["Colaborador", "Data"]).reset_index(drop=True)
    return df


# =========================================================
# Etapa 4 — Abas Semanais e Resumo Geral (HE)
# =========================================================

def build_weekly_sheets(df: pd.DataFrame) -> dict:
    weeks = {}
    zero = pd.Timedelta(0)

    for i in range(1, 6):
        wname = f"Semana {i}"
        sub = df[df["Semana"] == wname].copy()

        cols_out = [
            "Semana","Colaborador","CPF","Data",
            "Previstas","Trabalhadas","Atrasos",
            "HE 50%","HE 70%","HE 100%","Noturnas"
        ]

        if sub.empty:
            weeks[wname] = pd.DataFrame(columns=cols_out)
            continue

        sub = sub.sort_values(["Colaborador", "CPF", "Data"]).reset_index(drop=True)
        linhas = []

        for (colab, cpf), g in sub.groupby(["Colaborador", "CPF"], sort=False):
            for _, r in g.iterrows():
                worked_display = r["Trabalhadas_td"]
                if (r["Abonadas_td"] > zero) and (r["Trabalhadas_td"] == zero):
                    worked_display = r["Abonadas_td"]

                linhas.append({
                    "Semana": r["Semana"],
                    "Colaborador": r["Colaborador"],
                    "CPF": r["CPF"],
                    "Data": r["Data"].date(),
                    "Previstas": fmt_td_hms(r["Previstas_td"]),
                    "Trabalhadas": fmt_td_hms(worked_display),
                    "Atrasos": fmt_td_hms_signed(r["Atrasos_td"]),
                    "HE 50%": fmt_td_hms(r["HE50_td"] + r["Atrasos_td"]),  # 50 líquido no dia
                    "HE 70%": fmt_td_hms(r["HE70_td"]),
                    "HE 100%": fmt_td_hms(r["HE100_td"]),
                    "Noturnas": fmt_td_hms(r["Noturnas_td"]),
                })

            soma_prev = g["Previstas_td"].sum()
            soma_atra = g["Atrasos_td"].sum()
            soma_he50 = g["HE50_td"].sum()
            soma_he70 = g["HE70_td"].sum()
            soma_he100 = g["HE100_td"].sum()
            soma_not = g["Noturnas_td"].sum()

            worked_display_series = g["Trabalhadas_td"].copy()
            mask_abono_sem_trab = (g["Abonadas_td"] > zero) & (g["Trabalhadas_td"] == zero)
            worked_display_series = worked_display_series.where(~mask_abono_sem_trab, g["Abonadas_td"])
            soma_trab_display = worked_display_series.sum()

            total50_semana_td = soma_he50 + soma_atra

            linhas.append({
                "Semana": wname,
                "Colaborador": f"{colab} - TOTAL",
                "CPF": cpf,
                "Data": "",
                "Previstas": fmt_td_hms(soma_prev),
                "Trabalhadas": fmt_td_hms(soma_trab_display),
                "Atrasos": fmt_td_hms_signed(soma_atra),
                "HE 50%": fmt_td_hms_signed(total50_semana_td),
                "HE 70%": fmt_td_hms(soma_he70),
                "HE 100%": fmt_td_hms(soma_he100),
                "Noturnas": fmt_td_hms(soma_not),
            })

        weeks[wname] = pd.DataFrame(linhas, columns=cols_out)

    return weeks


def build_resumo_geral(df: pd.DataFrame) -> pd.DataFrame:
    grp = df.groupby(["Colaborador", "CPF"], as_index=False)[
        ["Previstas_td","Trabalhadas_td","Abonadas_td","HE50_td","HE70_td","HE100_td","Atrasos_td","Faltas_td","Noturnas_td"]
    ].sum()

    grp["Total50_td"] = grp["HE50_td"] + grp["Atrasos_td"]
    grp["Total70_td"] = grp["HE70_td"]
    grp["Total100_td"] = grp["HE100_td"]
    grp["Geral_td"] = grp["Total50_td"] + grp["Total70_td"] + grp["Total100_td"]

    out = pd.DataFrame()
    out["Colaborador"] = grp["Colaborador"]
    out["CPF"] = grp["CPF"]
    out["Previstas"] = grp["Previstas_td"].apply(fmt_td_hms)
    out["Trabalhadas"] = (grp["Trabalhadas_td"] + grp["Abonadas_td"]).apply(fmt_td_hms)
    out["Atrasos"] = grp["Atrasos_td"].apply(fmt_td_hms_signed)
    out["50% (líq)"] = grp["Total50_td"].apply(fmt_td_hms_signed)
    out["70%"] = grp["Total70_td"].apply(fmt_td_hms)
    out["100%"] = grp["Total100_td"].apply(fmt_td_hms)
    out["GERAL"] = grp["Geral_td"].apply(fmt_td_hms)
    out["Noturnas"] = grp["Noturnas_td"].apply(fmt_td_hms)

    return out.sort_values("Colaborador").reset_index(drop=True)


# =========================================================
# Relatório de Endereço (PONTO/ENDEREÇO)
# =========================================================

def read_pontos_endereco(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=6)
    df.columns = [str(c).strip() for c in df.columns]

    col_colab = "Colaborador"
    col_data_inicio = "Data"
    col_data_fim = [c for c in df.columns if c.startswith("Data.") or "Saída" in c][0]
    col_endereco = "Endereço"

    df[col_colab] = (
        df[col_colab]
        .astype(str)
        .replace("nan", "", regex=False)
        .str.replace(r"[\n\r\t]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    df = df[df[col_colab].str.split().str.len() >= 2].copy()
    df["Nome_2p"] = df[col_colab].apply(name_2p)

    df[col_data_inicio] = pd.to_datetime(df[col_data_inicio], errors="coerce")
    df[col_data_fim] = pd.to_datetime(df[col_data_fim], errors="coerce")

    # ✅ mantém linhas com pelo menos UMA marcação
    df = df[df[col_data_inicio].notna() | df[col_data_fim].notna()].copy()

    # ✅ se não tiver INÍCIO, usa FIM como início
    df[col_data_inicio] = df[col_data_inicio].fillna(df[col_data_fim])

    df["DATA"] = df[col_data_inicio].dt.date

    resumo = (
        df.groupby(["Nome_2p", "DATA"], as_index=False)
        .agg(
            INICIO=(col_data_inicio, "min"),
            FIM=(col_data_fim, "max"),
            ENDEREÇO=(col_endereco, "first")
        )
    )

    resumo.rename(columns={"Nome_2p": "Colaborador"}, inplace=True)
    resumo = resumo[["Colaborador", "DATA", "INICIO", "FIM", "ENDEREÇO"]]
    resumo = resumo.sort_values(by=["Colaborador", "DATA"]).reset_index(drop=True)
    return resumo


# =========================================================
# Matinal (VALORES PRONTOS, sem fórmula)
# =========================================================

def build_matinal_table(df_he_final: pd.DataFrame, df_end: pd.DataFrame):
    datas = sorted(df_end["DATA"].dropna().unique().tolist())
    if len(datas) < 2:
        raise ValueError("O relatório de endereço precisa ter pelo menos 2 dias para montar o Matinal.")

    dia_1, dia_2 = datas[-2], datas[-1]

    end_sub = df_end[df_end["DATA"].isin([dia_1, dia_2])].copy()
    colaboradores = sorted(end_sub["Colaborador"].unique().tolist())
    idx_end = {(r.Colaborador, r.DATA): r for r in end_sub.itertuples(index=False)}

    he = df_he_final.copy()
    he["Data"] = pd.to_datetime(he["Data"], errors="coerce").dt.date

    daily = (
        he.groupby(["Nome_2p", "Data"], as_index=False)[["HE50_td", "HE70_td", "HE100_td", "Atrasos_td"]]
          .sum()
    )
    daily["HE50_liq_td"] = daily["HE50_td"] + daily["Atrasos_td"]
    daily["HE70100_td"] = daily["HE70_td"] + daily["HE100_td"]

    acum50_before = (
        daily[daily["Data"] < dia_1]
        .groupby("Nome_2p", as_index=False)[["HE50_liq_td"]].sum()
        .rename(columns={"HE50_liq_td": "ACUM50_BEFORE"})
    )
    acum70100_to = (
        daily[daily["Data"] <= dia_1]
        .groupby("Nome_2p", as_index=False)[["HE70100_td"]].sum()
        .rename(columns={"HE70100_td": "ACUM70100"})
    )
    day1_vals = (
        daily[daily["Data"] == dia_1][["Nome_2p", "HE50_liq_td", "HE70100_td"]]
        .rename(columns={"HE50_liq_td": "D1_50", "HE70100_td": "D1_70100"})
    )

    base = pd.DataFrame({"Nome_2p": colaboradores})
    tmp = (
        base.merge(day1_vals, on="Nome_2p", how="left")
            .merge(acum50_before, on="Nome_2p", how="left")
            .merge(acum70100_to, on="Nome_2p", how="left")
    )

    zero = pd.Timedelta(0)
    for c in ["D1_50", "D1_70100", "ACUM50_BEFORE", "ACUM70100"]:
        tmp[c] = tmp[c].fillna(zero)

    # saída pronta
    rows = []
    for nome2p in colaboradores:
        rec1 = idx_end.get((nome2p, dia_1))
        rec2 = idx_end.get((nome2p, dia_2))

        t = tmp[tmp["Nome_2p"] == nome2p].iloc[0]
        acumulado_50 = t["ACUM50_BEFORE"] + t["D1_50"]

        rows.append({
            "COLABORADOR": nome2p.title(),
            "D1_INICIO": fmt_dt_time(rec1.INICIO) if rec1 else "",
            "D1_FIM": fmt_dt_time(rec1.FIM) if (rec1 and pd.notna(rec1.FIM)) else "",
            "D1_BASE": rec1.ENDEREÇO if rec1 else "",
            "D1_50": fmt_td_hms_signed(t["D1_50"]),
            "D1_70100": fmt_td_hms(t["D1_70100"]),
            "D2_INICIO": fmt_dt_time(rec2.INICIO) if rec2 else "",
            "D2_BASE": rec2.ENDEREÇO if rec2 else "",
            "ACUM_50": fmt_td_hms_signed(acumulado_50),
            "ACUM_70100": fmt_td_hms(t["ACUM70100"]),
        })

    return pd.DataFrame(rows), dia_1, dia_2


# =========================================================
# Matinal — Formatação igual ao print
# =========================================================

def create_matinal_sheet(wb, dia_1: date, dia_2: date):
    ws = wb.create_sheet("Matinal")

    purple = PatternFill("solid", fgColor="4C208E")
    white_bold = Font(color="FFFFFF", bold=True)
    title_font = Font(bold=True, size=18)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Linhas (como seu print)
    TITLE_ROW = 4
    GROUP_ROW = 6
    SUB_ROW = 7
    DATA_START_ROW = 8

    total_cols = 10
    last_col_letter = get_column_letter(total_cols)

    mes = month_name_pt(pd.Timestamp(dia_2).month)
    ws.merge_cells(f"A{TITLE_ROW}:{last_col_letter}{TITLE_ROW}")
    ws[f"A{TITLE_ROW}"] = f"Relatório Matinal - {mes}"
    ws[f"A{TITLE_ROW}"].font = title_font
    ws[f"A{TITLE_ROW}"].alignment = center

    # Cabeçalhos grupo
    ws.cell(GROUP_ROW, 1, "Nome")
    ws.cell(SUB_ROW, 1, "COLABORADOR")

    d1_txt = pd.Timestamp(dia_1).strftime("%d/%m/%Y")
    ws.merge_cells(start_row=GROUP_ROW, start_column=2, end_row=GROUP_ROW, end_column=6)
    ws.cell(GROUP_ROW, 2, d1_txt)
    subs_d1 = ["INÍCIO", "FIM", "BASE", "50%", "70% e 100%"]
    for j, s in enumerate(subs_d1):
        ws.cell(SUB_ROW, 2 + j, s)

    d2_txt = pd.Timestamp(dia_2).strftime("%d/%m/%Y")
    ws.merge_cells(start_row=GROUP_ROW, start_column=7, end_row=GROUP_ROW, end_column=8)
    ws.cell(GROUP_ROW, 7, d2_txt)
    ws.cell(SUB_ROW, 7, "INÍCIO")
    ws.cell(SUB_ROW, 8, "BASE")

    ws.merge_cells(start_row=GROUP_ROW, start_column=9, end_row=GROUP_ROW, end_column=10)
    ws.cell(GROUP_ROW, 9, "ACUMULADO")
    ws.cell(SUB_ROW, 9, "50%")
    ws.cell(SUB_ROW, 10, "70% e 100%")

    # Estilo cabeçalhos
    for rr in [GROUP_ROW, SUB_ROW]:
        for cc in range(1, total_cols + 1):
            cell = ws.cell(rr, cc)
            cell.fill = purple
            cell.font = white_bold
            cell.alignment = center
            cell.border = border

    # larguras
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 11
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 12

    ws.freeze_panes = f"B{DATA_START_ROW}"

    return ws, border, center, left, DATA_START_ROW


def write_matinal_values(ws, border, center, left, matinal_df: pd.DataFrame, start_row: int):
    r = start_row
    for _, row in matinal_df.iterrows():
        ws.cell(r, 1).value = row["COLABORADOR"]
        ws.cell(r, 1).alignment = left
        ws.cell(r, 1).border = border

        values = [
            row["D1_INICIO"],
            row["D1_FIM"],
            row["D1_BASE"],
            row["D1_50"],
            row["D1_70100"],
            row["D2_INICIO"],
            row["D2_BASE"],
            row["ACUM_50"],
            row["ACUM_70100"],
        ]

        for c, v in enumerate(values, start=2):
            cell = ws.cell(r, c)
            cell.value = v
            cell.alignment = center if c != 4 and c != 8 else center  # base também central no print
            cell.border = border

        r += 1


# =========================================================
# Excel final
# =========================================================

def generate_excel_bytes(df_he_final: pd.DataFrame, df_end: pd.DataFrame) -> bytes:
    weeks = build_weekly_sheets(df_he_final)
    resumo_geral = build_resumo_geral(df_he_final)

    matinal_df, dia_1, dia_2 = build_matinal_table(df_he_final, df_end)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Semanais
        for wname, wdf in weeks.items():
            wdf.to_excel(writer, sheet_name=wname, index=False)

        # Resumo Geral
        resumo_geral.to_excel(writer, sheet_name="Resumo Geral", index=False)

        wb = writer.book

        # Matinal formatado (igual ao print)
        ws_mat, border, center, left, DATA_START_ROW = create_matinal_sheet(wb, dia_1, dia_2)
        write_matinal_values(ws_mat, border, center, left, matinal_df, start_row=DATA_START_ROW)

        # remove "Sheet" se existir
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    output.seek(0)
    return output.getvalue()


# =========================================================
# Interface Streamlit
# =========================================================

def main():
    st.set_page_config(page_title="WS | Processador de Jornada + Matinal", layout="wide")
    st.title("📊 Processador de Jornada (HE) + Relatório Matinal — WS Transportes")

    st.markdown(
        """
**Como funciona agora (do jeito que você pediu):**
- Você envia **2 arquivos**:
  1) **Horas Extras / Folha de Ponto (HE)** → calcula 50/70/100 com suas regras.
  2) **Folha Ponto Endereço** → pega INÍCIO/FIM/BASE para os 2 últimos dias (dia atual pode vir sem FIM).
- A aba **Matinal** sai **sem fórmulas**, já com os valores preenchidos, no layout do print.
        """
    )

    colA, colB = st.columns(2)
    with colA:
        he_file = st.file_uploader("📁 1) Envie o relatório de Horas Extras / Folha de Ponto (.xlsx)", type=["xlsx"], key="he")
    with colB:
        end_file = st.file_uploader("📁 2) Envie o relatório Folha Ponto Endereço (.xlsx)", type=["xlsx"], key="end")

    if not he_file or not end_file:
        st.info("⬆️ Envie **os dois arquivos** para gerar o Excel completo.")
        return

    try:
        # ===== HE =====
        raw = pd.read_excel(he_file, sheet_name=0, header=None, dtype=str)
        long_df = detect_blocks_and_build_long(raw)
        df_he_final = clean_and_enrich(long_df)

        # ===== ENDEREÇO =====
        df_end = read_pontos_endereco(end_file)

        # previews
        st.success("✅ Arquivos processados!")
        st.subheader("Prévia HE (tratado)")
        st.dataframe(df_he_final[["Colaborador","Nome_2p","CPF","Data","Semana"]].head(30), use_container_width=True)

        st.subheader("Prévia Endereço (tratado)")
        st.dataframe(df_end.head(30), use_container_width=True)

        excel_bytes = generate_excel_bytes(df_he_final, df_end)

        st.download_button(
            label="⬇️ Baixar Excel completo (HorasExtras_WS_Completo.xlsx)",
            data=excel_bytes,
            file_name="HorasExtras_WS_Completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"⚠️ Erro ao processar: {e}")


if __name__ == "__main__":
    main()
