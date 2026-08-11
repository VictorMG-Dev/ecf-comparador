"""
generate_test_data.py — Gera dados de teste realistas para o comparador ECF.

Cria dois arquivos Excel (ecf_anterior.xlsx e ecf_atual.xlsx) com múltiplos
blocos e cenários de comparação: linhas incluídas, excluídas, alteradas e
sem alteração. Ideal para validar o comparador sem precisar de arquivos reais.

Uso:
    python generate_test_data.py
"""

from __future__ import annotations

import random
import string
from datetime import date

import pandas as pd

SEED = 42
random.seed(SEED)


def _cnpj() -> str:
    """Gera um CNPJ fictício formatado."""
    nums = [random.randint(0, 9) for _ in range(14)]
    return f"{''.join(map(str, nums[:2]))}.{''.join(map(str, nums[2:5]))}.{''.join(map(str, nums[5:8]))}/{''.join(map(str, nums[8:12]))}-{''.join(map(str, nums[12:]))}"


def _cod_cta(prefixo: str = "1", digitos: int = 6) -> str:
    sufixo = "".join([str(random.randint(0, 9)) for _ in range(digitos)])
    return f"{prefixo}.{sufixo[:2]}.{sufixo[2:4]}.{sufixo[4:]}"


def _valor() -> str:
    return str(round(random.uniform(1_000, 9_999_999), 2)).replace(".", ",")


def _data_str() -> str:
    return date(2024, random.randint(1, 12), random.randint(1, 28)).strftime("%d/%m/%Y")


# -----------------------------------------------------------------------
# Bloco 0000 — Identificação (1 linha por arquivo)
# -----------------------------------------------------------------------
def _bloco_0000_anterior() -> pd.DataFrame:
    return pd.DataFrame([{
        "DT_INI": "01/01/2024",
        "DT_FIN": "31/12/2024",
        "CNPJ": "12.345.678/0001-90",
        "NOME": "EMPRESA EXEMPLO LTDA",
        "NIRE": "3330012345",
        "IND_SIT_ESP": "0",
        "IND_SIT_INI_PER": "0",
        "IND_NIRE": "1",
        "IND_FIN_ESC": "0",
        "COD_HASH_ESC": "ABCD1234",
    }])


def _bloco_0000_atual() -> pd.DataFrame:
    df = _bloco_0000_anterior().copy()
    df["DT_INI"] = "01/01/2025"
    df["DT_FIN"] = "31/12/2025"
    df["COD_HASH_ESC"] = "WXYZ9999"   # campo alterado
    return df


# -----------------------------------------------------------------------
# Bloco M300 — Adições ao Lucro Líquido
# -----------------------------------------------------------------------
def _bloco_m300(n: int = 20) -> list[dict]:
    contas = [f"3.{i:02d}.001" for i in range(1, n + 1)]
    cnpjs = [_cnpj() for _ in range(n)]
    rows = []
    for cta, cnpj in zip(contas, cnpjs):
        rows.append({
            "COD_CTA": cta,
            "CNPJ_PARTE": cnpj,
            "DESCR_CTA": f"CONTA ADICAO {cta}",
            "VL_LANCTO": _valor(),
            "VL_AJUSTE": _valor(),
            "PERC_PART": str(round(random.uniform(1, 100), 2)).replace(".", ","),
        })
    return rows


def _blocos_m300_anterior_atual():
    base = _bloco_m300(20)
    # Anterior: linhas 0-17 (18 linhas)
    ant = base[:18]
    # Atual: linhas 0-14 preservadas, 15-16 alteradas, 17 removido, 18-19 adicionados
    atu = []
    for i, row in enumerate(base[:15]):
        atu.append(row.copy())
    for row in base[15:17]:
        r = row.copy()
        r["VL_LANCTO"] = _valor()   # valor alterado
        r["VL_AJUSTE"] = _valor()
        atu.append(r)
    # linha 17 (índice) removida do atual
    # linhas novas
    novas = _bloco_m300(2)
    for r in novas:
        r["COD_CTA"] = f"3.99.{random.randint(100,999)}"
        r["CNPJ_PARTE"] = _cnpj()
        atu.append(r)

    return pd.DataFrame(ant), pd.DataFrame(atu)


# -----------------------------------------------------------------------
# Bloco I051 — Plano de Contas Referencial
# -----------------------------------------------------------------------
def _bloco_i051(n: int = 30) -> list[dict]:
    rows = []
    for i in range(1, n + 1):
        rows.append({
            "COD_CTA_REF": f"REF.{i:04d}",
            "COD_CTA": _cod_cta(),
            "DESCR_CTA": f"CONTA REFERENCIAL {i:04d}",
            "COD_CTA_SUP": f"REF.{max(1, i-1):04d}",
            "IND_CTA": random.choice(["A", "S"]),
            "NIVEL": str(random.randint(1, 5)),
        })
    return rows


def _blocos_i051_anterior_atual():
    base = _bloco_i051(30)
    ant = base[:]
    atu = []
    for i, row in enumerate(base):
        if i < 25:
            atu.append(row.copy())
        # linhas 25-29 removidas no atual
    # Novas linhas no atual
    for j in range(31, 36):
        atu.append({
            "COD_CTA_REF": f"REF.{j:04d}",
            "COD_CTA": _cod_cta(),
            "DESCR_CTA": f"CONTA NOVA {j:04d}",
            "COD_CTA_SUP": "REF.0001",
            "IND_CTA": "A",
            "NIVEL": "3",
        })
    # Altera descrição de algumas
    atu[5]["DESCR_CTA"] = "CONTA ALTERADA 0005"
    atu[10]["IND_CTA"] = "S"
    return pd.DataFrame(ant), pd.DataFrame(atu)


# -----------------------------------------------------------------------
# Bloco J100 — Balanço Patrimonial
# -----------------------------------------------------------------------
def _bloco_j100(n: int = 25) -> list[dict]:
    rows = []
    for i in range(1, n + 1):
        rows.append({
            "COD_CTA": _cod_cta(str(random.randint(1, 5))),
            "COD_CCUS": f"CC{i:03d}",
            "DESCR_CTA": f"CONTA BP {i:03d}",
            "VL_INI_DEBITO": _valor(),
            "VL_INI_CREDITO": _valor(),
            "VL_FIN_DEBITO": _valor(),
            "VL_FIN_CREDITO": _valor(),
        })
    return rows


def _blocos_j100_anterior_atual():
    base = _bloco_j100(25)
    ant = base[:]
    atu = [row.copy() for row in base]
    # Altera valores de 8 linhas
    for idx in random.sample(range(len(atu)), 8):
        atu[idx]["VL_FIN_DEBITO"] = _valor()
        atu[idx]["VL_FIN_CREDITO"] = _valor()
    return pd.DataFrame(ant), pd.DataFrame(atu)


# -----------------------------------------------------------------------
# Bloco Y600 — Responsável (somente no anterior → bloco removido)
# -----------------------------------------------------------------------
def _bloco_y600_anterior() -> pd.DataFrame:
    return pd.DataFrame([{
        "CPF_RESP": "123.456.789-00",
        "NOME": "FULANO DE TAL",
        "CARGO": "CONTADOR",
        "FONE": "(11)99999-9999",
        "EMAIL": "fulano@empresa.com",
    }])


# -----------------------------------------------------------------------
# Bloco Q100 — Pagamentos (somente no atual → bloco novo)
# -----------------------------------------------------------------------
def _bloco_q100_atual() -> pd.DataFrame:
    rows = []
    for i in range(5):
        rows.append({
            "CNPJ_CPF_BENEF": _cnpj(),
            "COD_REC": f"REC{i+1:04d}",
            "NOME_BENEF": f"BENEFICIARIO {i+1}",
            "VL_PAGTO": _valor(),
            "IR_FONTE": _valor(),
        })
    return pd.DataFrame(rows)


# -----------------------------------------------------------------------
# Montagem dos arquivos
# -----------------------------------------------------------------------
def gerar_arquivos(
    caminho_anterior: str = "ecf_anterior.xlsx",
    caminho_atual: str = "ecf_atual.xlsx",
) -> None:
    m300_ant, m300_atu = _blocos_m300_anterior_atual()
    i051_ant, i051_atu = _blocos_i051_anterior_atual()
    j100_ant, j100_atu = _blocos_j100_anterior_atual()

    abas_anterior = {
        "0000": _bloco_0000_anterior(),
        "I051": i051_ant,
        "J100": j100_ant,
        "M300": m300_ant,
        "Y600": _bloco_y600_anterior(),
    }

    abas_atual = {
        "0000": _bloco_0000_atual(),
        "I051": i051_atu,
        "J100": j100_atu,
        "M300": m300_atu,
        "Q100": _bloco_q100_atual(),   # bloco novo
        # Y600 ausente no atual → bloco removido
    }

    with pd.ExcelWriter(caminho_anterior, engine="openpyxl") as writer:
        for aba, df in abas_anterior.items():
            df.to_excel(writer, sheet_name=aba, index=False)
    print(f"[OK] Arquivo anterior gerado: {caminho_anterior}")

    with pd.ExcelWriter(caminho_atual, engine="openpyxl") as writer:
        for aba, df in abas_atual.items():
            df.to_excel(writer, sheet_name=aba, index=False)
    print(f"[OK] Arquivo atual gerado: {caminho_atual}")


if __name__ == "__main__":
    gerar_arquivos()
