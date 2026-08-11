"""
reader.py — Leitura dos arquivos Excel de ECF.

Responsável por:
  - Ler todos os arquivos Excel de entrada.
  - Retornar um dicionário { nome_aba: DataFrame } para cada arquivo.
  - Validar e reportar abas vazias ou ilegíveis sem interromper a execução.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict

import pandas as pd

logger = logging.getLogger(__name__)


def ler_excel(caminho: str | Path) -> Dict[str, pd.DataFrame]:
    """
    Lê todas as abas de um arquivo Excel e retorna um dicionário
    { nome_aba: DataFrame }.

    Parâmetros
    ----------
    caminho : str ou Path
        Caminho completo para o arquivo .xlsx/.xls.

    Retorna
    -------
    Dict[str, pd.DataFrame]
        Chaves = nome da aba; Valores = DataFrame com os dados brutos.
        Abas vazias são incluídas como DataFrames vazios com aviso.

    Lança
    -----
    FileNotFoundError
        Se o arquivo não existir no caminho fornecido.
    ValueError
        Se o arquivo não for um Excel válido.
    """
    caminho = Path(caminho)

    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    if caminho.suffix.lower() not in (".xlsx", ".xls", ".xlsm", ".xlsb"):
        raise ValueError(
            f"Extensão não suportada: '{caminho.suffix}'. "
            "Use arquivos .xlsx, .xls, .xlsm ou .xlsb."
        )

    logger.info("Lendo arquivo: %s", caminho)

    try:
        xls = pd.ExcelFile(caminho, engine=_detectar_engine(caminho))
    except Exception as exc:
        raise ValueError(f"Não foi possível abrir o arquivo '{caminho}': {exc}") from exc

    abas: Dict[str, pd.DataFrame] = {}

    for nome_aba in xls.sheet_names:
        try:
            df = xls.parse(
                nome_aba,
                dtype=str,          # lê tudo como texto; normalização posterior
                keep_default_na=False,
                na_values=[""],
            )
            df = df.fillna("")      # células em branco → string vazia

            if df.empty or df.columns.empty:
                logger.warning(
                    "Aba '%s' está vazia ou sem colunas em '%s'. "
                    "Será tratada como bloco vazio.",
                    nome_aba, caminho.name,
                )
                abas[nome_aba] = pd.DataFrame()
            else:
                abas[nome_aba] = df
                logger.debug(
                    "Aba '%s' lida: %d linhas × %d colunas.",
                    nome_aba, len(df), len(df.columns),
                )

        except Exception as exc:  # noqa: BLE001
            logger.error(
                "Erro ao ler aba '%s' em '%s': %s. Aba ignorada.",
                nome_aba, caminho.name, exc,
            )
            abas[nome_aba] = pd.DataFrame()

    logger.info(
        "Arquivo '%s' carregado: %d abas encontradas.",
        caminho.name, len(abas),
    )
    return abas


# ---------------------------------------------------------------------------
# Utilitários internos
# ---------------------------------------------------------------------------

def _detectar_engine(caminho: Path) -> str:
    """Retorna o engine pandas adequado para a extensão do arquivo."""
    ext = caminho.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return "openpyxl"
    if ext == ".xlsb":
        return "pyxlsb"
    # .xls (legado)
    return "xlrd"
