"""
normalizer.py — Normalização dos dados de cada DataFrame de bloco ECF.

Responsável por:
  - Strip de espaços em campos texto.
  - Uniformização de maiúsculas/minúsculas.
  - Detecção e conversão de campos numéricos (tratando vírgula e ponto).
  - Arredondamento de floats para evitar falso-positivo por precisão.
  - Retorno do DataFrame normalizado sem modificar o original.
"""

from __future__ import annotations

import logging
import re
from typing import List

import pandas as pd

from config import DECIMAL_PLACES, NUMERIC_COLUMN_HINTS

logger = logging.getLogger(__name__)

# Regex: detecta strings que representam números (inteiro ou decimal com . ou ,)
_RE_NUMERO = re.compile(
    r"^\s*-?\s*\d{1,3}(?:[.\d{3}]*)?(?:[,\.]\d+)?\s*$"
)

# Regex mais simples: apenas dígitos opcionalmente com separadores
_RE_NUMERO_SIMPLES = re.compile(r"^\s*-?\d+([,\.]\d+)?\s*$")


def normalizar_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza todos os campos de um DataFrame:
      1. Converte tudo para string (caso já não seja).
      2. Aplica strip e uppercase em campos texto.
      3. Tenta converter colunas numéricas (detectadas automaticamente).
      4. Arredonda floats para DECIMAL_PLACES casas decimais.

    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame bruto lido do Excel.

    Retorna
    -------
    pd.DataFrame
        Cópia normalizada do DataFrame original.
    """
    if df.empty:
        return df.copy()

    df = df.copy()

    for col in df.columns:
        # --- passo 1: garantir string e strip ---
        df[col] = df[col].astype(str).str.strip()

        # --- passo 2: uppercase apenas em campos não-numéricos ---
        if not _e_coluna_numerica(df[col]):
            df[col] = df[col].str.upper()
        else:
            # --- passo 3: conversão numérica ---
            df[col] = _converter_coluna_numerica(df[col])

    return df


def _e_coluna_numerica(serie: pd.Series) -> bool:
    """
    Heurística: considera numérica se >= 70 % dos valores não-vazios
    correspondem ao padrão de número (inteiro ou decimal).
    """
    nao_vazios = serie[serie != ""]
    if nao_vazios.empty:
        return False

    qtd_num = nao_vazios.apply(lambda v: bool(_RE_NUMERO_SIMPLES.match(str(v)))).sum()
    return (qtd_num / len(nao_vazios)) >= 0.70


def _converter_coluna_numerica(serie: pd.Series) -> pd.Series:
    """
    Converte uma série de strings numéricas para float arredondado,
    lidando com separadores de milhar (ponto) e decimal (vírgula ou ponto).
    Valores não convertíveis permanecem como string.
    """
    def _parse(valor: str) -> str:
        valor = valor.strip()
        if valor in ("", "NAN", "NONE", "NULL"):
            return ""
        try:
            # Formato brasileiro: 1.234,56 → remove ponto, substitui vírgula
            if "," in valor and "." in valor:
                # ambos presentes: ponto = milhar, vírgula = decimal
                normalizado = valor.replace(".", "").replace(",", ".")
            elif "," in valor:
                # só vírgula: decimal brasileiro
                normalizado = valor.replace(",", ".")
            else:
                normalizado = valor

            numero = float(normalizado)
            arredondado = round(numero, DECIMAL_PLACES)
            # Remove zero decimal desnecessário: 10.00 → "10.0" (consistente)
            return str(arredondado)
        except (ValueError, TypeError):
            return valor

    return serie.apply(_parse)


def detectar_colunas_chave_automatica(
    df: pd.DataFrame,
    nome_bloco: str,
) -> List[str]:
    """
    Estratégia de fallback quando o bloco não possui chave configurada.

    Regras:
      1. Remove colunas cujo NOME contém algum padrão de NUMERIC_COLUMN_HINTS.
      2. Remove colunas identificadas como numéricas pelo conteúdo (_e_coluna_numerica).
      3. Retorna as colunas restantes como chave.
      4. Se sobrar apenas 1 coluna, usa ela mesmo que numérica.

    Emite aviso no log indicando chave automática detectada.
    """
    candidatas = []
    for col in df.columns:
        nome_upper = str(col).upper()
        # Exclui por dica de nome
        if any(hint.upper() in nome_upper for hint in NUMERIC_COLUMN_HINTS):
            continue
        # Exclui por conteúdo numérico
        if _e_coluna_numerica(df[col]):
            continue
        candidatas.append(col)

    if not candidatas:
        # Último recurso: usa todas as colunas
        candidatas = list(df.columns)

    logger.warning(
        "[CHAVE AUTOMÁTICA] Bloco '%s' não está na configuração. "
        "Usando chave automática: %s. "
        "Considere adicionar este bloco ao config.py para resultados precisos.",
        nome_bloco, candidatas,
    )
    return candidatas
