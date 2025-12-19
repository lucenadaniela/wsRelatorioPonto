# app_horas_extras.py
# -*- coding: utf-8 -*-

import re
import unicodedata
from io import BytesIO
from datetime import date

import pandas as pd
import numpy as np
import streamlit as st

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# Helpers
# =========================================================

def normalize_colname(name: str) -> str:
    if not isinstance(name, str):
        return ""
    s = name.strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def norm_person_key(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^A-Z ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_time_to_timedelta(value):
    if pd.isna(value):
        return pd.NaT

    s = str(value).strip()
    if s == "" or s in {"-", "—", "–", "--", "nan", "None"}:
        return pd.NaT

    if isinstance(value, pd.Timedelta):
        return value

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

    try:
        f = float(s.replace(",", "."))
        return pd.to_timedelta(f, unit="D")
    except Exception:
        return pd.NaT


def fmt_td_hms(td: pd.Timedelta) -> str:
    """Formato: H:MM:SS (igual print)."""
    if pd.isna(td):
        return "0:00:00"
    sec = int(td.total_seconds())
    if sec < 0:
        sec = 0
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h}:{m:02d}:{s:02d}"


def fmt_td_hms_signed(td: pd.Timedelta) -> str:
    if pd.isna(td):
        return "0:00:00"
    sec = int(td.total_seconds())
    sign = "-" if sec < 0 else ""
    sec = abs(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{sign}{h}:{m:02d}:{s:02d}"


def fmt_dt_time(dt) -> str:
    """INÍCIO/FIM em HH:MM:SS (ou vazio)."""
    if pd.isna(dt) or dt is None:
        return ""
    try:
        ts = pd.Timestamp(dt)
        return ts.strftime("%H:%M:%S")
    except Exception:
        return ""


def month_name_pt(m: int) -> str:
    nomes = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return nomes[m - 1] if 1 <= m <= 12 else ""


# =========================================================
# Mapeamento colunas (Horas Extras)
# =========================================================

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


# =========================================================
# Etapa 1 — Detectar cabeçalhos e montar tabela longa
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

            data_val = v_data
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
                    "Data": data_val,
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
            return "Semana 1"

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
# Etapa 2.1 — Feriados nacionais
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
# Etapa 3 — Limpeza + regras HE (ATUALIZADA p/ Nome_2p)
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

    # Data
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[~df["Data"].isna()].copy()

    # LIMPEZA NOME (igual teu outro código)
    df["Colaborador"] = (
        df["Colaborador"]
        .astype(str)
        .replace("nan", "", regex=False)
        .str.replace(r"[\n\r\t]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Mantém apenas nomes compostos
    df = df[df["Colaborador"].str.split().str.len() >= 2].copy()

    # Nome_2p igual teu código
    df["Nome_2p"] = df["Colaborador"].apply(lambda x: " ".join(x.split()[:2]))

    # CPF
    df["CPF"] = df["CPF"].astype(str).str.strip()

    # Data como date (pra casar com endereço e pra RESUMO)
    df["Data"] = df["Data"].dt.date

    # Timedeltas
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

    # Domingo / feriado (usa Timestamp pra dayofweek)
    data_ts = pd.to_datetime(df["Data"], errors="coerce")

    df["is_domingo"] = data_ts.dt.dayofweek == 6
    df["is_feriado"] = data_ts.apply(is_national_holiday)
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

    # Semana
    df["Semana"] = classify_weeks(pd.to_datetime(df["Data"], errors="coerce"))
    df = df.sort_values(["Nome_2p", "Data"]).reset_index(drop=True)

    # mantém (se você ainda usa em algum lugar)
    df["ColabKey"] = df["Colaborador"].apply(norm_person_key)

    return df


# =========================================================
# Semanais e Resumo Geral (mantidos)
# =========================================================

def _fmt_hhmm(td: pd.Timedelta) -> str:
    if pd.isna(td):
        return "00:00"
    sec = int(td.total_seconds())
    if sec < 0:
        sec = 0
    h = sec // 3600
    m = (sec % 3600) // 60
    return f"{h:02d}:{m:02d}"

def _fmt_hhmm_signed(td: pd.Timedelta) -> str:
    if pd.isna(td):
        return "00:00"
    sec = int(td.total_seconds())
    sign = "-" if sec < 0 else ""
    sec = abs(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    return f"{sign}{h:02d}:{m:02d}"

def build_weekly_sheets(df: pd.DataFrame) -> dict:
    weeks = {}
    zero = pd.Timedelta(0)

    for i in range(1, 6):
        wname = f"Semana {i}"
        sub = df[df["Semana"] == wname].copy()

        if sub.empty:
            weeks[wname] = pd.DataFrame(
                columns=[
                    "Semana", "Colaborador", "CPF", "Data",
                    "Previstas", "Trabalhadas", "Atrasos",
                    "HE 50%", "HE 70%", "HE 100%", "Noturnas",
                ]
            )
            continue

        sub = sub.sort_values(["Nome_2p", "CPF", "Data"]).reset_index(drop=True)
        linhas = []

        for (nome2p, cpf), g in sub.groupby(["Nome_2p", "CPF"], sort=False):
            for _, r in g.iterrows():
                worked_display = r["Trabalhadas_td"]
                if (r["Abonadas_td"] > zero) and (r["Trabalhadas_td"] == zero):
                    worked_display = r["Abonadas_td"]

                linhas.append(
                    {
                        "Semana": r["Semana"],
                        "Colaborador": nome2p.title(),
                        "CPF": r["CPF"],
                        "Data": r["Data"],
                        "Previstas": _fmt_hhmm(r["Previstas_td"]),
                        "Trabalhadas": _fmt_hhmm(worked_display),
                        "Atrasos": _fmt_hhmm_signed(r["Atrasos_td"]),
                        "HE 50%": _fmt_hhmm(r["HE50_td"]),
                        "HE 70%": _fmt_hhmm(r["HE70_td"]),
                        "HE 100%": _fmt_hhmm(r["HE100_td"]),
                        "Noturnas": _fmt_hhmm(r["Noturnas_td"]),
                    }
                )

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

            linhas.append(
                {
                    "Semana": wname,
                    "Colaborador": f"{nome2p.title()} - TOTAL",
                    "CPF": cpf,
                    "Data": "",
                    "Previstas": _fmt_hhmm(soma_prev),
                    "Trabalhadas": _fmt_hhmm(soma_trab_display),
                    "Atrasos": _fmt_hhmm_signed(soma_atra),
                    "HE 50%": _fmt_hhmm_signed(total50_semana_td),
                    "HE 70%": _fmt_hhmm(soma_he70),
                    "HE 100%": _fmt_hhmm(soma_he100),
                    "Noturnas": _fmt_hhmm(soma_not),
                }
            )

        weeks[wname] = pd.DataFrame(linhas)

    return weeks


def build_resumo_geral(df: pd.DataFrame) -> pd.DataFrame:
    grp = df.groupby(["Nome_2p", "CPF"], as_index=False)[
        ["Previstas_td", "Trabalhadas_td", "Abonadas_td", "HE50_td", "HE70_td", "HE100_td", "Atrasos_td", "Noturnas_td"]
    ].sum()

    grp["Total50_td"] = grp["HE50_td"] + grp["Atrasos_td"]
    grp["Total70_td"] = grp["HE70_td"]
    grp["Total100_td"] = grp["HE100_td"]
    grp["Geral_td"] = grp["Total50_td"] + grp["Total70_td"] + grp["Total100_td"]

    out = pd.DataFrame()
    out["Colaborador"] = grp["Nome_2p"].str.title()
    out["CPF"] = grp["CPF"]
    out["Previstas"] = grp["Previstas_td"].apply(_fmt_hhmm)
    out["Trabalhadas"] = (grp["Trabalhadas_td"] + grp["Abonadas_td"]).apply(_fmt_hhmm)
    out["Atrasos"] = grp["Atrasos_td"].apply(_fmt_hhmm_signed)
    out["50% (liq)"] = grp["Total50_td"].apply(_fmt_hhmm_signed)
    out["70%"] = grp["Total70_td"].apply(_fmt_hhmm)
    out["100%"] = grp["Total100_td"].apply(_fmt_hhmm)
    out["GERAL"] = grp["Geral_td"].apply(_fmt_hhmm)
    out["Noturnas"] = grp["Noturnas_td"].apply(_fmt_hhmm)

    return out.sort_values("Colaborador").reset_index(drop=True)


# =========================================================
# Endereço (EXATAMENTE teu código)
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

    # Mantém apenas nomes compostos
    df = df[df[col_colab].str.split().str.len() >= 2].copy()
    df["Nome_2p"] = df[col_colab].apply(lambda x: " ".join(x.split()[:2]))

    df[col_data_inicio] = pd.to_datetime(df[col_data_inicio], errors="coerce")
    df[col_data_fim] = pd.to_datetime(df[col_data_fim], errors="coerce")

    # somente com início e fim (igual teu código)
    df = df[df[col_data_inicio].notna() & df[col_data_fim].notna()].copy()
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
# RESUMO (aba técnica) para fórmulas do Matinal
# =========================================================

def build_resumo_ws(df_he_final: pd.DataFrame, dia_1: date, colaboradores_ordem: list[str]) -> pd.DataFrame:
    """
    Monta tabela com colunas A..S (19) para:
    - G: 50% do dia_1 (HE50 + atrasos)
    - H: 70+100 do dia_1
    - M: acumulado 70+100 até dia_1
    - S: acumulado 50% antes do dia_1
    """
    zero = pd.Timedelta(0)

    tmp = df_he_final.copy()
    tmp["Data"] = pd.to_datetime(tmp["Data"], errors="coerce").dt.date

    daily = (
        tmp.groupby(["Nome_2p", "Data"], as_index=False)[["HE50_td", "HE70_td", "HE100_td", "Atrasos_td"]]
        .sum()
    )
    daily["HE50_liq_td"] = daily["HE50_td"] + daily["Atrasos_td"]
    daily["HE70100_td"] = daily["HE70_td"] + daily["HE100_td"]

    d1 = daily[daily["Data"] == dia_1][["Nome_2p", "HE50_liq_td", "HE70100_td"]].copy()
    d1 = d1.rename(columns={"HE50_liq_td": "DIA1_50", "HE70100_td": "DIA1_70100"})

    acum50_antes = (
        daily[daily["Data"] < dia_1]
        .groupby("Nome_2p", as_index=False)[["HE50_liq_td"]].sum()
        .rename(columns={"HE50_liq_td": "ACUM50_ANTES"})
    )

    acum70100_ate = (
        daily[daily["Data"] <= dia_1]
        .groupby("Nome_2p", as_index=False)[["HE70100_td"]].sum()
        .rename(columns={"HE70100_td": "ACUM70100_ATE"})
    )

    base = pd.DataFrame({"Nome_2p": colaboradores_ordem})
    out = (
        base.merge(d1, on="Nome_2p", how="left")
            .merge(acum70100_ate, on="Nome_2p", how="left")
            .merge(acum50_antes, on="Nome_2p", how="left")
    )

    for c in ["DIA1_50", "DIA1_70100", "ACUM70100_ATE", "ACUM50_ANTES"]:
        out[c] = out[c].fillna(zero)

    # cria colunas A..S (19)
    cols = [f"COL{i}" for i in range(1, 20)]
    resumo = pd.DataFrame({c: [""] * len(out) for c in cols})

    resumo["COL1"] = out["Nome_2p"].str.title()                # A
    resumo["COL7"] = out["DIA1_50"].apply(fmt_td_hms_signed)    # G
    resumo["COL8"] = out["DIA1_70100"].apply(fmt_td_hms)        # H
    resumo["COL13"] = out["ACUM70100_ATE"].apply(fmt_td_hms)    # M
    resumo["COL19"] = out["ACUM50_ANTES"].apply(fmt_td_hms_signed)  # S

    return resumo


# =========================================================
# Matinal (sheet formatada) + fórmulas
# =========================================================

def create_matinal_sheet(wb, dia_1: date, dia_2: date):
    purple = PatternFill("solid", fgColor="4C208E")
    white_bold = Font(color="FFFFFF", bold=True)
    title_font = Font(bold=True, size=18)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws = wb.create_sheet("Matinal")

    # Layout fixo: 10 colunas
    # A Nome
    # Dia1: B INÍCIO, C FIM, D BASE, E 50%, F 70+100
    # Dia2: G INÍCIO, H BASE
    # ACUM: I 50%, J 70+100
    total_cols = 10
    last_col_letter = get_column_letter(total_cols)

    mes = month_name_pt(dia_2.month)
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = f"Relatório Matinal - {mes}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    # Cabeçalhos grupo
    r_group = 2
    r_sub = 3

    # Nome
    ws.cell(r_group, 1, "Nome")
    ws.cell(r_sub, 1, "COLABORADOR")

    # Dia 1
    d1_txt = pd.Timestamp(dia_1).strftime("%d/%m/%Y")
    ws.merge_cells(start_row=r_group, start_column=2, end_row=r_group, end_column=6)
    ws.cell(r_group, 2, d1_txt)
    subs_d1 = ["INÍCIO", "FIM", "BASE", "50%", "70% e 100%"]
    for j, s in enumerate(subs_d1):
        ws.cell(r_sub, 2 + j, s)

    # Dia 2
    d2_txt = pd.Timestamp(dia_2).strftime("%d/%m/%Y")
    ws.merge_cells(start_row=r_group, start_column=7, end_row=r_group, end_column=8)
    ws.cell(r_group, 7, d2_txt)
    ws.cell(r_sub, 7, "INÍCIO")
    ws.cell(r_sub, 8, "BASE")

    # Acumulado
    ws.merge_cells(start_row=r_group, start_column=9, end_row=r_group, end_column=10)
    ws.cell(r_group, 9, "ACUMULADO")
    ws.cell(r_sub, 9, "50%")
    ws.cell(r_sub, 10, "70% e 100%")

    # Estilos dos cabeçalhos (2 e 3)
    for rr in [r_group, r_sub]:
        for cc in range(1, total_cols + 1):
            cell = ws.cell(rr, cc)
            cell.fill = purple
            cell.font = white_bold
            cell.alignment = center
            cell.border = border

    # colunas
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 34
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 11
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 12

    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 20

    ws.freeze_panes = "B4"

    return ws, border, center, left


def write_matinal_rows(ws, border, center, left, df_end: pd.DataFrame, dia_1: date, dia_2: date, start_row: int = 6):
    """
    Preenche Matinal com:
    - dados do df_end (INÍCIO/FIM/BASE)
    - fórmulas para horas puxando da aba RESUMO
    """
    sub = df_end[df_end["DATA"].isin([dia_1, dia_2])].copy()
    colaboradores = sorted(sub["Colaborador"].unique().tolist())

    idx = {(r.Colaborador, r.DATA): r for r in sub.itertuples(index=False)}

    r = start_row
    for colab in colaboradores:
        # A - nome
        ws.cell(r, 1).value = colab.title()
        ws.cell(r, 1).alignment = left
        ws.cell(r, 1).border = border

        # Dia 1
        rec1 = idx.get((colab, dia_1))
        ws.cell(r, 2).value = fmt_dt_time(rec1.INICIO) if rec1 else ""
        ws.cell(r, 3).value = fmt_dt_time(rec1.FIM) if rec1 else ""
        ws.cell(r, 4).value = rec1.ENDEREÇO if rec1 else ""

        # horas por fórmula (dia 1)
        ws.cell(r, 5).value = f"=RESUMO!G{r}"             # 50%
        ws.cell(r, 6).value = f"=RESUMO!H{r}"             # 70+100

        # Dia 2
        rec2 = idx.get((colab, dia_2))
        ws.cell(r, 7).value = fmt_dt_time(rec2.INICIO) if rec2 else ""
        ws.cell(r, 8).value = rec2.ENDEREÇO if rec2 else ""

        # acumulado por fórmula
        ws.cell(r, 9).value = f"=RESUMO!S{r}+RESUMO!G{r}" # ACUM 50
        ws.cell(r, 10).value = f"=RESUMO!M{r}"            # ACUM 70+100

        # estilo borda/centro nas demais
        for c in range(2, 11):
            ws.cell(r, c).alignment = center
            ws.cell(r, c).border = border
        # base alinhada ao centro fica ok; se quiser left nas bases:
        ws.cell(r, 4).alignment = center
        ws.cell(r, 8).alignment = center

        r += 1

    return colaboradores


# =========================================================
# Excel final
# =========================================================

def generate_excel_bytes(df_he_final: pd.DataFrame, df_end: pd.DataFrame) -> bytes:
    weeks = build_weekly_sheets(df_he_final)
    resumo_geral = build_resumo_geral(df_he_final)

    datas = sorted(df_end["DATA"].unique().tolist())
    if len(datas) < 2:
        raise ValueError("O relatório de endereço precisa ter pelo menos 2 dias para montar o Matinal.")
    dia_1, dia_2 = datas[-2], datas[-1]

    # Ordem do Matinal manda a ordem do RESUMO (pra fórmula bater)
    sub = df_end[df_end["DATA"].isin([dia_1, dia_2])]
    colaboradores_ordem = sorted(sub["Colaborador"].unique().tolist())

    resumo_ws = build_resumo_ws(df_he_final, dia_1, colaboradores_ordem)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Semanais
        for wname, wdf in weeks.items():
            wdf.to_excel(writer, sheet_name=wname, index=False)

        # Resumo Geral
        resumo_geral.to_excel(writer, sheet_name="Resumo Geral", index=False)

        # RESUMO (aba técnica) começa na linha 6 (startrow=5) p/ bater G6 etc
        resumo_ws.to_excel(writer, sheet_name="RESUMO", index=False, startrow=5)

        wb = writer.book

        # Matinal (formatado)
        ws_mat, border, center, left = create_matinal_sheet(wb, dia_1, dia_2)
        write_matinal_rows(ws_mat, border, center, left, df_end, dia_1, dia_2, start_row=6)

        # remove "Sheet" se existir
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    output.seek(0)
    return output.getvalue()


# =========================================================
# Streamlit UI
# =========================================================

def main():
    st.set_page_config(page_title="WS | Jornada (HE + Matinal)", layout="wide")
    st.title("📊 Processador de Jornada — HE + Matinal (igual ao print)")

    st.markdown(
        """
**Como funciona:**
- **Arquivo 1:** Horas Extras / Folha de Ponto → calcula HE 50/70/100 (regras).
- **Arquivo 2:** Relatório de Pontos com Endereço → monta Matinal (INÍCIO/FIM/BASE) com **Nome_2p**.
- O Matinal puxa horas por **fórmulas** da aba **RESUMO** (G/H/M/S).
        """
    )

    col1, col2 = st.columns(2)
    with col1:
        horas_file = st.file_uploader("1) Envie Horas Extras / Folha de Ponto (.xlsx)", type=["xlsx"], key="he")
    with col2:
        endereco_file = st.file_uploader("2) Envie Relatório de Pontos com Endereço (.xlsx)", type=["xlsx"], key="end")

    if not horas_file or not endereco_file:
        st.info("⬆️ Envie os dois arquivos para gerar o Excel completo.")
        return

    try:
        # Horas Extras
        raw = pd.read_excel(horas_file, sheet_name=0, header=None, dtype=str)
        long_df = detect_blocks_and_build_long(raw)
        df_he_final = clean_and_enrich(long_df)

        # Endereço (matinal base)
        df_end = read_pontos_endereco(endereco_file)

        st.success("✅ Processado com sucesso!")

        st.subheader("Prévia — Endereço (base do Matinal)")
        st.dataframe(df_end.head(40), use_container_width=True)

        st.subheader("Prévia — HE (tratado)")
        st.dataframe(df_he_final[["Nome_2p", "Data", "Semana"]].head(40), use_container_width=True)

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
