"""
reader.py — Leitura dos arquivos ECF (Excel .xlsx/.xls ou Texto SPED .txt).

Responsável por:
  - Ler arquivos Excel ou TXT de entrada.
  - Se for TXT (SPED ECF oficial): parsear registros em blocos/registros (ex: 0000, M300, etc).
  - Retornar um dicionário { nome_bloco: DataFrame } para cada arquivo.
  - Validar e reportar blocos vazios sem interromper a execução.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict

import pandas as pd

logger = logging.getLogger(__name__)


def ler_arquivo_ecf(caminho: str | Path) -> Dict[str, pd.DataFrame]:
    """
    Lê um arquivo ECF (seja Excel .xlsx/.xls ou arquivo de texto SPED .txt)
    e retorna um dicionário { nome_bloco: DataFrame }.
    """
    caminho = Path(caminho)

    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    ext = caminho.suffix.lower()
    if ext == ".txt":
        return ler_txt_sped(caminho)
    elif ext in (".xlsx", ".xls", ".xlsm", ".xlsb"):
        return ler_excel(caminho)
    else:
        raise ValueError(
            f"Extensão não suportada: '{ext}'. "
            "Use arquivos Excel (.xlsx, .xls) ou arquivo texto SPED ECF (.txt)."
        )


def ler_txt_sped(caminho: Path) -> Dict[str, pd.DataFrame]:
    """
    Lê arquivo .txt do SPED ECF (formato pipe '|REG|CAMPO1|CAMPO2|...|')
    e agrupa as linhas por tipo de Registro/Bloco.
    """
    logger.info("Lendo arquivo TXT SPED: %s", caminho)
    registros: Dict[str, list[list[str]]] = {}

    # Tenta UTF-8 ou Latin-1 / ISO-8859-1 (comum em arquivos SPED da RFB)
    conteudo = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with open(caminho, "r", encoding=enc) as f:
                conteudo = f.readlines()
            logger.debug("Arquivo TXT lido com encoding %s", enc)
            break
        except UnicodeDecodeError:
            continue

    if conteudo is None:
        raise ValueError(f"Não foi possível ler o arquivo texto '{caminho}' com encondings padrão (UTF-8, Latin-1).")

    for linha in conteudo:
        linha_str = linha.strip()
        if not linha_str or not linha_str.startswith("|"):
            continue

        partes = linha_str.split("|")
        # Formato SPED: |REGISTRO|CAMPO1|CAMPO2|...|
        if len(partes) >= 3:
            nome_reg = partes[1].upper().strip()
            if not nome_reg:
                continue
            valores = [p.strip() for p in partes[2:-1]] if partes[-1] == "" else [p.strip() for p in partes[2:]]
            
            if nome_reg not in registros:
                registros[nome_reg] = []
            registros[nome_reg].append(valores)

    abas: Dict[str, pd.DataFrame] = {}
    for nome_reg, linhas in registros.items():
        if not linhas:
            continue
        max_cols = max(len(l) for l in linhas)
        # Gera nomes de colunas C1, C2, C3...
        col_names = [f"CAMPO_{i+1}" for i in range(max_cols)]
        
        # Ajusta linhas mais curtas se houver
        linhas_normalizadas = [l + [""] * (max_cols - len(l)) for l in linhas]
        
        df = pd.DataFrame(linhas_normalizadas, columns=col_names, dtype=str)
        df = df.fillna("")
        abas[nome_reg] = df
        logger.debug("Bloco TXT '%s' lido: %d linhas × %d colunas.", nome_reg, len(df), len(df.columns))

    logger.info("Arquivo TXT '%s' carregado: %d registros/blocos encontrados.", caminho.name, len(abas))
    return abas


def ler_excel(caminho: Path) -> Dict[str, pd.DataFrame]:
    """
    Lê todas as abas de um arquivo Excel e retorna um dicionário
    { nome_aba: DataFrame }.
    """
    logger.info("Lendo arquivo Excel: %s", caminho)

    try:
        xls = pd.ExcelFile(caminho, engine=_detectar_engine(caminho))
    except Exception as exc:
        raise ValueError(f"Não foi possível abrir o arquivo '{caminho}': {exc}") from exc

    abas: Dict[str, pd.DataFrame] = {}

    for nome_aba in xls.sheet_names:
        try:
            df = xls.parse(
                nome_aba,
                dtype=str,          # lê tudo como texto
                keep_default_na=False,
                na_values=[""],
            )
            df = df.fillna("")

            if df.empty or df.columns.empty:
                logger.warning(
                    "Aba '%s' está vazia ou sem colunas em '%s'.",
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
                "Erro ao ler aba '%s' em '%s': %s.",
                nome_aba, caminho.name, exc,
            )
            abas[nome_aba] = pd.DataFrame()

    logger.info(
        "Arquivo Excel '%s' carregado: %d abas encontradas.",
        caminho.name, len(abas),
    )
    return abas


def _detectar_engine(caminho: Path) -> str:
    """Retorna o engine pandas adequado para a extensão do arquivo."""
    ext = caminho.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return "openpyxl"
    if ext == ".xlsb":
        return "pyxlsb"
    return "xlrd"

