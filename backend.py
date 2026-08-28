# -*- coding: utf-8 -*-
"""
backend.py
-------------------------------------------------
Contém toda a lógica de negócio, acesso ao Google Sheets e geração de PDF
para a aplicação Gestor do Bolsão. Este módulo é independente da interface gráfica.
"""
# --- Importações de Módulos ---
import re
import uuid
from datetime import date, timedelta, datetime
from functools import lru_cache
from pathlib import Path
import sys
import os
import base64 
import json
import requests 
import pytz

import gspread
import pandas as pd
import weasyprint
from google.oauth2.service_account import Credentials

# --------------------------------------------------
# UTILITÁRIOS DE ACESSO AO GOOGLE SHEETS (OTIMIZADOS)
# --------------------------------------------------
SPREAD_URL = "https://docs.google.com/spreadsheets/d/1qBV70qrPswnAUDxnHfBgKEU4FYAISpL7iVP0IM9zU2Q/edit#gid=0"

GCP_CREDS_B64 = '''ewogICJ0eXBlIjogInNlcnZpY2VfYWNjb3VudCIsCiAgInByb2plY3RfaWQiOiAiZ2VuLWxhbmctY2xpZW50LTA2Mzk5NTMwMzciLAogICJwcml2YXRlX2tleV9pZCI6ICJkMmE5ZjJkZTE3NTUwN
mVmYmIyNDRhMmM4NzlhOTk3ZjM0NmQ2ODA3IiwKICAicHJpdmF0ZV9rZXkiOiAiLS0tLS1CRUdJTiBQUklWQVRFIEtFWS0tLS0tXG5NSUlFdlFJQkFEQU5CZ2txaGtpRzl3MEJBUUVGQUFTQ0JLY3
dnZ1NqQWdFQUFvSUJBUURGZjhkNVVCZWM0bmJQXG5VY0Z3MzFmRVlRTUdFYzlUamNLZDdaUytqaGk5MndZNDQ3TzNLeDM2M1dZWFJBcFI0TlNzbDIzQnl4UGhjSTJHXG5xZS8rNS96dUdGRm1VYlB
vWHIwMEFzZEF3YXo1Q2ZnZXFsK2g3aGQ5RkF6VWFFbkF6NkR5eGtQc25rU2o1d2JRXG53Z3NaQ0llSTR6bjArWVlrdGo5enllQUwvZkppM2RXVWtMbmdtUldyTEw4R1hid1Z1TXlUdHFseXdoeEN4
WXlzXG5mN3RrcGZmQ2V1WXhGL0JaYmg1NnpTYkx0OXowNGxXTk1qNzlkTWZQN2U1eVVCN21CK2NPRDNNcUQwb2x5VDUvXG5DR054dzFpemo5bnFxdXE1RUcyRFBhRksyOU5uOUpCY3Q2WUZwTGlDU
WNleGNRN0pNWm93dGhMaGUxMzN1eTNOXG5xckpxNzdoM0FnTUJBQUVDZ2dFQUd6U1JMTDZCL05seFFQa09hdGJ6M0cya29LY1VFU1ZtUS9lZ2dCTkxpSWp0XG5pdFAxbUpZTVBpeXJUZFI3TzNEUX
V2N3RDVDlKSmgvVCthaUZFbThGekxmSzFyWmdldkk1OW8vVXYxc3VmRUVKXG5veEp0cy96NzBxaWxaUjVYTW9QU0NaU1N4Mlc4Zk5EQkxwYUc0dFE2OUtTZ1Y0VGJ4cDR0a0JkUE8rSlNZTzNXXG4
zd1ZwSnZyY05WVlprRUx2TDFrZC9rcU5uMVBLT0ZqSnVhVGhoWkZaNzZzeVFxeS83R29iMksxT255ME4yNzAxXG5NVjNDL1Fia1F2MEN0bEIxdW90RjBhaUptc0k5amtpWkFZaXRwMWxRNjdOeUxL
cDViaEJYT2U1VlZyRnVubUU5XG5qcjcrMDd5NmxuSi9pdWYwSHg3UkJCNS94QkRiMENDc1dZRWhiZEJLK1FLQmdRRDhweGJ3U3V1L1NKUGlNYmxnXG5RNkNlMkJXbDR0TTFQcTlrR0I2OHBMang5U
GY1NTAxVEk4L1gvaE8rT3hxTVh1VWNNbUFBYS9uY2wvTkRTa095XG54T01sczNCMlNlK09tU1pPY3p5VFYxL25pMzlUNmdiRkZiK3lzSWlrWHByN0pQQ0d6bGw1V1RRbkZWeVI3Z2ZUXG5rTWxxNk
RPdVdZZWxDUGo1QytvQ3lNOTRvd0tCZ1FESUhhQzUvSTZTN0NSMWZTVk9NUlV1bDBtdERDN2V6ckNJXG5rUENldUNHWG14a096K2xmelVXa1d4aitDcEhGdkZPSk92b3kvcndUWGRRU3lKbyt6aE5
jcXpGZXVtQVlQbTZsXG42cHV4MnNQWlhuWlpOcmdSN0lDbE5RWlJoWFNLYXIwbFd5aHBhYnBIUDBqenlJeHViblVySUVGNkNjaERlaXNIXG5oWjhrTEgrYUhRS0JnRElMRXBNVDgvQVdleHhCaHh0
WEtkaHNxcFVLQVpXNVRkcEFvTTN5dnFOR1IxdmJnY2htXG4vb01rMDFwSnNLOS9HRmhtYmZlSGE1MVRCNThiZFo5U09qKzhkQmtwK2VZLzFZT1NkYndsZ1Z5R2wvalNFUGMyXG5vSnJQTnJGLy82a
WhIM2RFSmhyUUhYRmtYTXFXNER2Nm9McXBOWnRpSm1qOUJ2dWxiWHVSdE82TEFvR0JBS29LXG53dkJHOHI5VmVRVjhlWjZmaG1nNGxabzlwL1libTFYd0V3WkxYLy83QXFmYmMycThlbGpXTDc1dU
xML2c3dnJFXG42cVU2WDRYMVRLZEpYZ2ZRQkJQU1EvbENWVFZFbDdEVVVjZ25KQlFvZUZnR2J4S0w3Q1Ixa2hEalhjdmdOeWp1XG45TC9kM3pON1N2b2JoM2l1MDFENTc2bExkcWdmdjBCOUFtVXl
pcHJoQW9HQVlXTTNKYmQ3bE1rMUk3R2NMT1I5XG5kQmkxZU02YzJ6cGlGOGNEK0VVdHNKTHdqTCt2RklBeUU2N0E3R0drVEdHNFB5RGp2T0h4aSsrVTQ1RjFDTGxJXG5pRmtld1EwRWtkYjJJYlpN
NzFCUnYvQndVcG9mY1BMMWxzNm4vdGNxS3lOTlBxYjNXV2Q0YzRBL0RlcWlIVWhiXG5QWFJOOEJYU3Vwd2tVZGdITW1DWVk1UT1cbi0tLS0tRU5EIFBSSVZBVEUgS0VZLS0tLS1cbiIsCiAgImNsa
WVudF9lbWFpbCI6ICJib2xzLW9AZ2VuLWxhbmctY2xpZW50LTA2Mzk5NTMwMzcuaWFtLmdzZXJ2aWNlYWNjb3VudC5jb20iLAogICJjbGllbnRfaWQiOiAiMTEyNzYzNjc5NDcxMzk1MzcxMDg2Ii
wKICAiYXV0aF91cmkiOiAiaHR0cHM6Ly9hY2NvdW50cy5nb29nbGUuY29tL28vb2F1dGgyL2F1dGgiLAogICJ0b2tlbl91cmkiOiAiaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW4
iLAogICJhdXRoX3Byb3ZpZGVyX3g1MDlfY2VydF91cmwiOiAiaHR0cHM6Ly93d3cuZ29vZ2xlYXBpcy5jb20vb2F1dGgyL3YxL2NlcnRzIiwKICAiY2xpZW50X3g1MDlfY2VydF91cmwiOiAiaHR0
cHM6Ly93d3cuZ29vZ2xlYXBpcy5jb20vcm9ib3QvdjEvbWV0YWRhdGEveDUwOS9ib2xzLW8lNDBnZW4tbGFuZy1jbGllbnQtMDYzOTk1MzAzNy5pYW0uZ3NlcnZpY2VhY2NvdW50LmNvbSIsCiAgI
nVuaXZlcnNlX2RvbWFpbiI6ICJnb29nbGVhcGlzLmNvbSIKfQo=='''

def resource_path(relative_path):
    """ Obtém o caminho absoluto para um recurso, funcionando para dev e para PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_gspread_client():
    """Conecta ao Google Sheets usando credenciais embutidas no código."""
    try:
        decoded_creds_json = base64.b64decode(GCP_CREDS_B64)
        creds_dict = json.loads(decoded_creds_json)
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        raise Exception(f"❌ Erro de autenticação com o Google Sheets a partir das credenciais embutidas: {e}")

client_cache = None
workbook_cache = None

def get_cached_client():
    """Retorna o cliente gspread em cache ou cria um novo."""
    global client_cache
    if client_cache is None:
        client_cache = get_gspread_client()
    return client_cache

def get_cached_workbook():
    """Retorna o workbook (planilha) em cache ou abre um novo."""
    global workbook_cache
    client = get_cached_client()
    if workbook_cache is None and client:
        workbook_cache = client.open_by_url(SPREAD_URL)
    return workbook_cache

@lru_cache(maxsize=32)
def get_ws(title: str):
    """Obtém uma aba (worksheet) pelo título e faz cache."""
    wb = get_cached_workbook()
    if wb:
        try:
            return wb.worksheet(title)
        except gspread.WorksheetNotFound:
            raise gspread.WorksheetNotFound(f"Aba da planilha com o nome '{title}' não foi encontrada.")
    return None

@lru_cache(maxsize=32)
def header_map(ws_title: str):
    """Cria um mapa de 'nome_da_coluna': indice para uma dada aba."""
    ws = get_ws(ws_title)
    if ws:
        headers = ws.row_values(1)
        return {h.strip(): i + 1 for i, h in enumerate(headers) if h and h.strip()}
    return {}

def get_values(ws, a1_range: str):
    """Função auxiliar para leitura de um range específico."""
    return ws.get(a1_range, value_render_option="UNFORMATTED_VALUE")

def find_row_by_id(ws, id_col_idx: int, target_id: str):
    """Encontra o número da linha de um registro pelo seu ID."""
    try:
        col_values = ws.col_values(id_col_idx)[1:]
        for i, value in enumerate(col_values, start=2):
            if str(value) == str(target_id):
                return i
    except Exception:
        return None
    return None

def batch_update_cells(ws, updates):
    """Executa múltiplas atualizações de células em uma única requisição à API."""
    if not updates:
        return
    fixed = []
    sheet_title_safe = ws.title.replace("'", "''")
    for u in updates:
        rng = u.get("range", "")
        if not rng:
            continue
        if "!" not in rng:
            rng = f"'{sheet_title_safe}'!{rng}"
        fixed.append({"range": rng, "values": u.get("values", [[]])})
    body = {"valueInputOption": "USER_ENTERED", "data": fixed}
    ws.spreadsheet.values_batch_update(body)

def ensure_size(ws, min_rows=2000, min_cols=40):
    """Garante que a planilha tenha um tamanho mínimo para evitar erros."""
    try:
        if ws and (ws.row_count < min_rows or ws.col_count < min_cols):
            ws.resize(rows=max(ws.row_count, min_rows), cols=max(ws.col_count, min_cols))
    except Exception:
        pass

def new_uuid():
    """Gera um ID único e curto (12 caracteres) para cada registro."""
    return uuid.uuid4().hex[:12]

def a1_col_letter(col_idx: int) -> str:
    """Converte um índice numérico de coluna (1, 2, 3) para sua letra A1 ('A', 'B', 'C')."""
    return re.sub(r"\d", "", gspread.utils.rowcol_to_a1(1, col_idx))

def batch_get_values_prefixed(ws, ranges, value_render_option="UNFORMATTED_VALUE"):
    """Faz uma leitura em lote de múltiplos ranges em uma única requisição."""
    if not ranges:
        return []
    title_safe = ws.title.replace("'", "''")
    prefixed = [f"'{title_safe}'!{r}" if "!" not in r else r for r in ranges]
    params = {'valueRenderOption': value_render_option}
    resp = ws.spreadsheet.values_batch_get(prefixed, params=params)
    return resp.get("valueRanges", [])

def load_resultados_snapshot():
    """
    Função otimizada para carregar os dados da aba 'Resultados_Bolsao'.
    """
    ws = get_ws("Resultados_Bolsao")
    if not ws:
        return {"rows": [], "id_to_rownum": {}}

    hmap = header_map("Resultados_Bolsao")
    
    columns_needed = [
        "REGISTRO_ID", "Nome do Aluno", "Unidade", "Bolsão", "% Bolsa", 
        "Valor da Mensalidade com Bolsa", "Escola de Origem", "Valor Negociado",
        "Responsável Financeiro", "Telefone", "Aluno Matriculou?", 
        "Observações (Form)", "Data/Hora"
    ]

    col_expectativa = "Expectativa de mensalidade"
    col_expectativa_fallback = "Valor Limite (PIA)"
    if col_expectativa in hmap:
        columns_needed.append(col_expectativa)
    elif col_expectativa_fallback in hmap:
        columns_needed.append(col_expectativa_fallback)

    missing = [c for c in columns_needed if c not in hmap]
    if missing:
        if not (len(missing) == 1 and missing[0] in [col_expectativa, col_expectativa_fallback]):
                 raise RuntimeError(f"Faltam colunas em 'Resultados_Bolsao': {', '.join(missing)}")

    letters = {c: a1_col_letter(hmap[c]) for c in columns_needed if c in hmap}
    ranges = [f"{letters[c]}2:{letters[c]}" for c in columns_needed if c in hmap]

    vranges = batch_get_values_prefixed(ws, ranges)
    series = {}
    valid_columns_from_fetch = [c for c in columns_needed if c in hmap]
    for c, vr in zip(valid_columns_from_fetch, vranges):
        vals = vr.get("values", [])
        series[c] = [row[0] if row else "" for row in vals]

    max_len = max((len(v) for v in series.values()), default=0)
    for c in valid_columns_from_fetch:
        col = series[c]
        if len(col) < max_len:
            col.extend([""] * (max_len - len(col)))

    rows = [{c: series.get(c, [""] * max_len)[i] for c in columns_needed} for i in range(max_len)]

    id_to_rownum = {}
    for i, rid in enumerate(series.get("REGISTRO_ID", []), start=2):
        if rid:
            id_to_rownum[str(rid)] = i

    return {"rows": rows, "id_to_rownum": id_to_rownum}

# --------------------------------------------------
# DADOS DE REFERÊNCIA E CONFIGURAÇÕES (CONSTANTES)
# --------------------------------------------------
TUITION = {
    "1ª e 2ª Série EM Militar": {"anuidade": 40263.66, "parcela13": 3097.20},
    "1ª e 2ª Série EM Vestibular": {"anuidade": 40263.66, "parcela13": 3097.20},
    "1º ao 5º Ano": {"anuidade": 29266.09, "parcela13": 2251.24},
    "3ª Série (PV/PM)": {"anuidade": 40419.58, "parcela13": 3109.20},
    "3ª Série EM Medicina": {"anuidade": 40419.58, "parcela13": 3109.20},
    "6º ao 8º Ano": {"anuidade": 34426.69, "parcela13": 2648.21},
    "9º Ano EF II Militar": {"anuidade": 37492.31, "parcela13": 2884.02},
    "9º Ano EF II Vestibular": {"anuidade": 37492.31, "parcela13": 2884.02},
    "AFA/EN/EFOMM": {"anuidade": 16252.60, "parcela13": 1250.20},
    "CN/EPCAr": {"anuidade": 9731.57, "parcela13": 748.58},
    "ESA": {"anuidade": 7845.21, "parcela13": 603.48},
    "EsPCEx": {"anuidade": 16252.60, "parcela13": 1250.20},
    "IME/ITA": {"anuidade": 16252.60, "parcela13": 1250.20},
    "Medicina (Pré)": {"anuidade": 16252.60, "parcela13": 1250.20},
    "Pré-Vestibular": {"anuidade": 16252.60, "parcela13": 1250.20},
}
TURMA_DE_INTERESSE_MAP = {
    "1ª série IME ITA Jr": "1ª e 2ª Série EM Militar", "1ª série do EM - Militar": "1ª e 2ª Série EM Militar",
    "1ª série do EM - Pré-Vestibular": "1ª e 2ª Série EM Vestibular", "1º ano do EF1": "1º ao 5º Ano",
    "2ª série IME ITA Jr": "1ª e 2ª Série EM Militar", "2ª série do EM - Militar": "1ª e 2ª Série EM Militar",
    "2ª série do EM - Pré-Vestibular": "1ª e 2ª Série EM Vestibular", "2º ano do EF1": "1º ao 5º Ano",
    "3ª série do EM - AFA EN EFOMM": "3ª Série (PV/PM)", "3ª série do EM - ESA": "3ª Série (PV/PM)",
    "3ª série do EM - EsPCEx": "3ª Série (PV/PM)", "3ª série do EM - IME ITA": "3ª Série (PV/PM)",
    "3ª série do EM - Medicina": "3ª Série EM Medicina", "3ª série do EM - Pré-Vestibular": "3ª Série (PV/PM)",
    "3º ano do EF1": "1º ao 5º Ano", "4º ano do EF1": "1º ao 5º Ano", "5º ano do EF1": "1º ao 5º Ano",
    "6º ano do EF2": "6º ao 8º Ano", "7º ano do EF2": "6º ao 8º Ano", "8º ano do EF2": "6º ao 8º Ano",
    "9º ano do EF2 - Militar": "9º Ano EF II Militar", "9º ano do EF2 - Vestibular": "9º Ano EF II Vestibular",
    "Pré-Militar AFA EN EFOMM": "AFA/EN/EFOMM", "Pré-Militar CN EPCAr": "CN/EPCAr", "Pré-Militar ESA": "ESA",
    "Pré-Militar EsPCEx": "EsPCEx", "Pré-Militar IME ITA": "IME/ITA", "Pré-Vestibular": "Pré-Vestibular",
    "Pré-Vestibular - Medicina": "Medicina (Pré)",
}
UNIDADES_COMPLETAS = [
    "COLEGIO E CURSO MATRIZ EDUCACAO CAMPO GRANDE", "COLEGIO E CURSO MATRIZ EDUCAÇÃO TAQUARA",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO BANGU", "COLEGIO E CURSO MATRIZ EDUCACAO NOVA IGUACU",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO DUQUE DE CAXIAS", "COLEGIO E CURSO MATRIZ EDUCAÇÃO SÃO JOÃO DE MERITI",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO ROCHA MIRANDA", "COLEGIO E CURSO MATRIZ EDUCAÇÃO MADUREIRA",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO RETIRO DOS ARTISTAS", "COLEGIO E CURSO MATRIZ EDUCACAO TIJUCA",
    "COLEGIO E CURSO MATRIZ EDUCACAO RECREIO"
]
UNIDADES_MAP = {name.replace("COLEGIO E CURSO MATRIZ EDUCACAO", "").replace("COLEGIO E CURSO MATRIZ EDUCAÇÃO", "").strip(): name for name in UNIDADES_COMPLETAS}
UNIDADES_LIMPAS = sorted(list(UNIDADES_MAP.keys()))
DESCONTOS_MAXIMOS_POR_UNIDADE = {
    "RETIRO DOS ARTISTAS": 0.50, "CAMPO GRANDE": 0.6320, "ROCHA MIRANDA": 0.6606,
    "TAQUARA": 0.6755, "NOVA IGUACU": 0.6700, "DUQUE DE CAXIAS": 0.6823,
    "BANGU": 0.6806, "MADUREIRA": 0.7032, "TIJUCA": 0.6800, "SÃO JOÃO DE MERITI": 0.7197,
}

# --- NOVA ESTRUTURA PARA REGRAS DE BOLSA ---
REGRAS_BOLSA_POR_UNIDADE_SERIE = {
    "BANGU": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "CAMPO GRANDE": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.54, (3, 5): 0.57, (6, 8): 0.60, (9, 11): 0.63, (12, 14): 0.66, (15, 17): 0.69, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.56, (3, 5): 0.59, (6, 8): 0.62, (9, 11): 0.65, (12, 14): 0.68, (15, 17): 0.71, (18, 20): 0.76, (21, 23): 0.81, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.56, (3, 5): 0.59, (6, 8): 0.62, (9, 11): 0.65, (12, 14): 0.68, (15, 17): 0.71, (18, 20): 0.76, (21, 23): 0.81, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.56, (3, 5): 0.59, (6, 8): 0.62, (9, 11): 0.65, (12, 14): 0.68, (15, 17): 0.71, (18, 20): 0.76, (21, 23): 0.81, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
    },
    "DUQUE DE CAXIAS": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "MADUREIRA": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "NOVA IGUACU": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "RETIRO DOS ARTISTAS": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.49, (3, 5): 0.52, (6, 8): 0.55, (9, 11): 0.58, (12, 14): 0.61, (15, 17): 0.64, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.54, (3, 5): 0.57, (6, 8): 0.60, (9, 11): 0.63, (12, 14): 0.66, (15, 17): 0.69, (18, 20): 0.74, (21, 23): 0.79, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.51, (3, 5): 0.54, (6, 8): 0.57, (9, 11): 0.60, (12, 14): 0.63, (15, 17): 0.66, (18, 20): 0.71, (21, 23): 0.76, (24, 24): 1.00 },
    },
    "ROCHA MIRANDA": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.77, (21, 23): 0.82, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "SÃO JOÃO DE MERITI": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.62, (3, 5): 0.65, (6, 8): 0.68, (9, 11): 0.71, (12, 14): 0.74, (15, 17): 0.77, (18, 20): 0.82, (21, 23): 0.87, (24, 24): 1.00 },
    },
    "TAQUARA": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1º ao 5º Ano": { (0, 2): 0.57, (3, 5): 0.60, (6, 8): 0.63, (9, 11): 0.66, (12, 14): 0.69, (15, 17): 0.72, (18, 20): 0.75 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
    "TIJUCA": {
        "1ª e 2ª Série EM Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "1ª e 2ª Série EM Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série (PV/PM)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "3ª Série EM Medicina": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "6º ao 8º Ano": { (0, 2): 0.59, (3, 5): 0.62, (6, 8): 0.65, (9, 11): 0.68, (12, 14): 0.71, (15, 17): 0.74, (18, 20): 0.79, (21, 23): 0.84, (24, 24): 1.00 },
        "9º Ano EF II Militar": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "9º Ano EF II Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "AFA/EN/EFOMM": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "CN/EPCAr": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "ESA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "EsPCEx": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "IME/ITA": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Medicina (Pré)": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
        "Pré-Vestibular": { (0, 2): 0.61, (3, 5): 0.64, (6, 8): 0.67, (9, 11): 0.70, (12, 14): 0.73, (15, 17): 0.76, (18, 20): 0.81, (21, 23): 0.86, (24, 24): 1.00 },
    },
}


# --------------------------------------------------
# FUNÇÕES DE LÓGICA E UTILITÁRIOS
# --------------------------------------------------
def get_current_brasilia_datetime() -> datetime:
    """Obtém a data e hora atuais de Brasília usando o horário do sistema."""
    # Removemos a consulta à API externa instável (worldtimeapi)
    # e usamos direto o pytz, que é robusto e rápido.
    try:
        br_tz = pytz.timezone("America/Sao_Paulo")
        return datetime.now(br_tz)
    except Exception as e:
        # Fallback extremo caso o pytz falhe, usa o horário local da máquina sem timezone
        print(f"Aviso: Erro ao definir timezone. Usando hora local do sistema. Erro: {e}")
        return datetime.now()

def get_current_brasilia_date() -> date:
    """Função auxiliar que retorna apenas a data de Brasília."""
    return get_current_brasilia_datetime().date()

@lru_cache(maxsize=1)
def get_bolsao_name_for_date(target_date=None):
    """Verifica a data e retorna o nome do bolsão ou 'Bolsão Avulso'."""
    if target_date is None:
        target_date = get_current_brasilia_date()
    try:
        ws_bolsao = get_ws("Bolsão")
        if not ws_bolsao: return "Bolsão Avulso" # Proteção caso a aba não carregue
        
        dates_cells = ws_bolsao.get('A2:A', value_render_option='FORMATTED_STRING')
        names_cells = ws_bolsao.get('C2:C')
        dates_col = [cell[0] for cell in dates_cells if cell]
        names_col = [cell[0] for cell in names_cells if cell]
        
        for i, date_str in enumerate(dates_col):
            if i < len(names_col) and names_col[i]:
                try:
                    bolsao_date = datetime.strptime(date_str, "%d/%m/%Y").date()
                    if bolsao_date == target_date:
                        return names_col[i]
                except ValueError: continue
        return "Bolsão Avulso"
    except Exception: return "Bolsão Avulso"

def precos_2027(serie_modalidade: str) -> dict:
    """Busca os preços corretos no dicionário TUITION para 2027."""
    base = TUITION.get(serie_modalidade, {})
    if not base:
        return {"primeira_cota": 0.0, "parcela_mensal": 0.0, "anuidade": 0.0}
    
    valor_mensal = float(base.get("parcela13", 0.0))
    primeira_cota = valor_mensal
    anuidade_total = float(base.get("anuidade", primeira_cota * 13))
    
    return {
        "primeira_cota": primeira_cota, 
        "parcela_mensal": valor_mensal, 
        "anuidade": anuidade_total
    }

# --- FUNÇÃO DE CÁLCULO DE BOLSA ATUALIZADA ---
def calcula_bolsa(acertos: int, serie_modalidade: str, unidade: str) -> float:
    """
    Calcula o percentual de bolsa com base na unidade, série/modalidade e número de acertos.
    """
    regras_unidade = REGRAS_BOLSA_POR_UNIDADE_SERIE.get(unidade)
    if not regras_unidade:
        print(f"Aviso: Regras de bolsa não encontradas para a unidade '{unidade}'. Usando 0% de bolsa.")
        return 0.0

    tabela_bolsa = regras_unidade.get(serie_modalidade)
    if not tabela_bolsa:
        print(f"Aviso: Série '{serie_modalidade}' não possui regras de bolsa para a unidade '{unidade}'. Usando 0% de bolsa.")
        return 0.0

    # Procura a faixa de acertos correta na tabela
    for (min_acertos, max_acertos), percentual in tabela_bolsa.items():
        if min_acertos <= acertos <= max_acertos:
            return percentual
    
    print(f"Aviso: Nenhum percentual encontrado para {acertos} acertos na série {serie_modalidade} da unidade {unidade}. Usando 0%.")
    return 0.0

    regras_unidade = REGRAS_BOLSA_POR_UNIDADE.get(unidade)
    if not regras_unidade:
        print(f"Aviso: Regras de bolsa não encontradas para a unidade '{unidade}'. Usando 0% de bolsa.")
        return 0.0

    tabela_bolsa = regras_unidade.get(segmento)
    if not tabela_bolsa:
        print(f"Aviso: Segmento '{segmento}' não possui regras de bolsa para a unidade '{unidade}'. Usando 0% de bolsa.")
        return 0.0

    # Procura a faixa de acertos correta na tabela
    for (min_acertos, max_acertos), percentual in tabela_bolsa.items():
        if min_acertos <= acertos <= max_acertos:
            return percentual
    
    # Caso o número de acertos não se encaixe em nenhuma faixa (ex: acertos negativos)
    print(f"Aviso: Nenhum percentual encontrado para {acertos} acertos no segmento {segmento} da unidade {unidade}. Usando 0%.")
    return 0.0

def format_currency(v: float) -> str:
    """Formata um número float para uma string de moeda brasileira (ex: R$ 1.234,56)."""
    try:
        v_float = float(v)
        return f"R$ {v_float:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except (ValueError, TypeError):
        return str(v)

def parse_brl_to_float(x) -> float:
    """Converte uma string de moeda brasileira (ex: 'R$ 1.234,56') para um float (1234.56)."""
    if isinstance(x, (int, float)):
        return float(x)
    if not x:
        return 0.0
    s = str(x).strip().replace("R$", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def format_phone_mask(raw: str) -> str:
    """Aplica uma máscara de telefone (##) #####-#### a uma string de dígitos."""
    if raw is None:
        return ""
    digits = re.sub(r"\D", "", str(raw))
    digits = digits[:11]
    if len(digits) >= 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:11]}"
    elif len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:10]}"
    return digits

def gera_pdf_html(ctx: dict) -> bytes:
    """
    Gera um arquivo PDF a partir de um template HTML e um dicionário de dados.
    Retorna os bytes do PDF gerado.
    """
    base_dir = Path(__file__).parent
    html_path = base_dir / "carta.html"
    try:
        with open(html_path, encoding="utf-8") as f:
            html_template = f.read()
        html_renderizado = html_template
        for k, v in ctx.items():
            html_renderizado = html_renderizado.replace(f"{{{{{k}}}}}", str(v))
        html_obj = weasyprint.HTML(string=html_renderizado, base_url=str(base_dir))
        return html_obj.write_pdf()
    except FileNotFoundError:
        raise Exception("Arquivo 'carta.html' ou 'style.css' não encontrado no diretório.")
    except Exception as e:
        raise Exception(f"Erro ao gerar PDF: {e}")

def get_hubspot_data_for_activation():
    """Obtém dados da aba 'Hubspot' para a funcionalidade de carregar candidato."""
    try:
        ws_hub = get_ws("Hubspot")
        if not ws_hub:
            return pd.DataFrame()

        hmap_h = header_map("Hubspot")
        cols_needed = ["Unidade", "Nome do Candidato", "Contato ID", "Status do Contato",
                       "Contato Realizado", "Observações", "Celular Tratado", "Nome",
                       "E-mail", "Turma de Interesse - Geral", "Fonte original"]
        missing_cols = [c for c in cols_needed if c not in hmap_h]
        if missing_cols:
            raise Exception(f"As seguintes colunas necessárias não foram encontradas na aba 'Hubspot': {', '.join(missing_cols)}")

        data = ws_hub.get_all_records(head=1)
        df = pd.DataFrame(data)
        if "Contato Realizado" in df.columns:
            df.rename(columns={"Contato Realizado": "Contato realizado"}, inplace=True)
        return df

    except Exception as e:
        raise Exception(f"❌ Falha ao carregar dados do Hubspot: {e}")

def calcula_valor_minimo(unidade, serie_modalidade):
    """Calcula o valor mínimo de parcela negociável para uma unidade e série."""
    try:
        desconto_maximo = DESCONTOS_MAXIMOS_POR_UNIDADE.get(unidade, 0)
        precos = precos_2027(serie_modalidade)
        valor_anuidade_integral = precos.get("anuidade", 0.0)
        if valor_anuidade_integral > 0 and desconto_maximo > 0:
            valor_minimo_anual = valor_anuidade_integral * (1 - desconto_maximo)
            return valor_minimo_anual / 12
        else:
            return 0.0
    except Exception as e:
        raise Exception(f"❌ Erro ao calcular valor mínimo: {e}")

def gerar_html_material_didatico(unidade: str) -> str:
    """
    Gera o código HTML para as tabelas de material didático
    com base na unidade selecionada.
    """
    precos_gerais = {
        "Medicina": ("R$ 4.512,73", "12x de R$ 376,57"),
        "Pré-Vestibular": ("R$ 4.512,73", "12x de R$ 376,57"),
    }
    
    precos_militares = {
        "AFA/EN/EFOMM": ("R$ 2.562,44", "12x de R$ 213,54"),
        "EPCAR": ("R$ 2.746,49", "12x de R$ 228,88"),
        "ESA": ("R$ 1.220,95", "12x de R$ 101,75"),
        "EsPCEx": ("R$ 2.930,53", "12x de R$ 244,21"),
        "IME/ITA": ("R$ 2.562,44", "12x de R$ 213,54"),
    }

    precos_didatico_padrao = {
        "1ª ao 5ª ano": ("R$ 2.802,97", "12x de R$ 233,58"),
        "6ª ao 8ª ano": ("R$ 3.036,82", "12x de R$ 253,07"),
        "9ª ano Vestibular": ("R$ 3.154,21", "12x de R$ 262,85"),
        "1ª e 2ª série Vestibular": ("R$ 3.842,64", "12x de R$ 320,22"),
        "3ª série": ("R$ 4.512,73", "12x de R$ 376,57"),
    }

    precos_sao_joao = {
        "1ª ao 5ª ano": ("R$ 2.123,05", "12x de R$ 176,92"),
        "6ª ao 8ª ano": ("R$ 2.218,97", "12x de R$ 184,91"),
        "9ª ano Vestibular": ("R$ 2.217,78", "12x de R$ 184,82"),
        "1ª e 2ª série Vestibular": ("R$ 2.826,47", "12x de R$ 235,54"),
        "3ª série": ("R$ 3.329,37", "12x de R$ 277,44"),
    }
    
    precos_retiro = {
        "1ª ao 5ª ano": ("R$ 2.802,97", "12x de R$ 233,58"),
    }
    
    dados_didatico = {}
    
    if unidade == "SÃO JOÃO DE MERITI":
        titulo_didatico = "Material Didático (exclusivo São João de Meriti)"
        dados_didatico = precos_sao_joao
    elif unidade == "RETIRO DOS ARTISTAS":
        titulo_didatico = "Material Didático"
        dados_didatico = precos_didatico_padrao.copy()
        dados_didatico.update(precos_retiro)
    else:
        titulo_didatico = "Material Didático"
        dados_didatico = precos_didatico_padrao

    tabela_didatico_html = f'<table class="pag2"><tr><th colspan="3">{titulo_didatico}</th></tr>'
    for curso, valores in dados_didatico.items():
        tabela_didatico_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_didatico_html += '</table><br>'

    tabela_geral_html = '<table class="pag2"><tr><th colspan="3">Material Didático (geral)</th></tr>'
    for curso, valores in precos_gerais.items():
        tabela_geral_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_geral_html += '</table><br>'
    
    tabela_militares_html = '<table class="pag2"><tr><th colspan="3">Material Militares</th></tr>'
    for curso, valores in precos_militares.items():
        tabela_militares_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_militares_html += '</table><br>'

    return tabela_didatico_html + tabela_geral_html + tabela_militares_html

