"""
comparator.py — Lógica de comparação entre dois DataFrames de um bloco ECF.

Responsável por:
  - Receber dois DataFrames (anterior e atual) já normalizados.
  - Fazer o merge (outer join) pela chave configurada (ou automática).
  - Classificar cada linha com o Status correto.
  - Preencher a coluna "Campos Alterados".
  - Retornar o DataFrame resultante pronto para ser escrito no Excel.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

from config import CHAVES_BLOCO, LABEL_ANTERIOR, LABEL_ATUAL
from normalizer import detectar_colunas_chave_automatica

logger = logging.getLogger(__name__)

# Sufixos internos usados durante o merge (não expostos ao usuário)
_SFX_ANT = f"_{LABEL_ANTERIOR}"
_SFX_ATU = f"_{LABEL_ATUAL}"

# Nome das colunas especiais no resultado
COL_STATUS = "Status"
COL_CAMPOS = "Campos Alterados"

STATUS_INCLUIDO = "Incluído"
STATUS_EXCLUIDO = "Excluído"
STATUS_ALTERADO = "Alterado"
STATUS_SEM_ALT = "Sem alteração"


@dataclass
class ResultadoBloco:
    """Contêiner com o DataFrame comparado e estatísticas do bloco."""

    nome_bloco: str
    df_resultado: pd.DataFrame
    chaves_usadas: List[str]
    chave_automatica: bool = False
    qtd_incluidos: int = 0
    qtd_excluidos: int = 0
    qtd_alterados: int = 0
    qtd_sem_alteracao: int = 0
    aviso: str = ""
    erros: List[str] = field(default_factory=list)


def comparar_bloco(
    nome_bloco: str,
    df_anterior: Optional[pd.DataFrame],
    df_atual: Optional[pd.DataFrame],
) -> ResultadoBloco:
    """
    Compara dois DataFrames de um mesmo bloco ECF e retorna o ResultadoBloco.

    Parâmetros
    ----------
    nome_bloco : str
        Nome do bloco (ex: "M300", "I051").
    df_anterior : pd.DataFrame | None
        DataFrame do ano anterior (já normalizado). None se bloco ausente.
    df_atual : pd.DataFrame | None
        DataFrame do ano atual (já normalizado). None se bloco ausente.

    Retorna
    -------
    ResultadoBloco
        Objeto com DataFrame de resultado e estatísticas.
    """
    resultado = ResultadoBloco(
        nome_bloco=nome_bloco,
        df_resultado=pd.DataFrame(),
        chaves_usadas=[],
    )

    # -----------------------------------------------------------------------
    # Caso especial: bloco novo (só existe no atual)
    # -----------------------------------------------------------------------
    if df_anterior is None or df_anterior.empty:
        if df_atual is None or df_atual.empty:
            resultado.aviso = "Bloco vazio nos dois arquivos."
            return resultado

        resultado.aviso = "Bloco NOVO — presente apenas no arquivo atual."
        resultado.qtd_incluidos = len(df_atual)
        df_saida = _construir_bloco_apenas_um_lado(
            df_atual, STATUS_INCLUIDO, lado="atual"
        )
        resultado.df_resultado = df_saida
        resultado.chaves_usadas = []
        return resultado

    # -----------------------------------------------------------------------
    # Caso especial: bloco removido (só existe no anterior)
    # -----------------------------------------------------------------------
    if df_atual is None or df_atual.empty:
        resultado.aviso = "Bloco REMOVIDO — presente apenas no arquivo anterior."
        resultado.qtd_excluidos = len(df_anterior)
        df_saida = _construir_bloco_apenas_um_lado(
            df_anterior, STATUS_EXCLUIDO, lado="anterior"
        )
        resultado.df_resultado = df_saida
        resultado.chaves_usadas = []
        return resultado

    # -----------------------------------------------------------------------
    # Resolução da chave
    # -----------------------------------------------------------------------
    chave_automatica = False
    if nome_bloco in CHAVES_BLOCO:
        chaves = CHAVES_BLOCO[nome_bloco]
        chaves = _validar_chaves(chaves, df_anterior, df_atual, nome_bloco, resultado)
    else:
        # Usa a heurística de chave automática sobre o DataFrame maior
        df_ref = df_atual if len(df_atual) >= len(df_anterior) else df_anterior
        chaves = detectar_colunas_chave_automatica(df_ref, nome_bloco)
        chave_automatica = True
        resultado.aviso = (
            f"Chave automática detectada: {chaves}. "
            "Adicione este bloco ao config.py para resultado preciso."
        )

    if not chaves:
        resultado.erros.append(
            f"Bloco '{nome_bloco}': nenhuma coluna de chave válida encontrada. "
            "Bloco ignorado."
        )
        return resultado

    resultado.chaves_usadas = chaves
    resultado.chave_automatica = chave_automatica

    # -----------------------------------------------------------------------
    # Determinação das colunas de valor (não-chave)
    # -----------------------------------------------------------------------
    todas_colunas = _uniao_ordenada(
        list(df_anterior.columns), list(df_atual.columns)
    )
    colunas_valor = [c for c in todas_colunas if c not in chaves]

    # -----------------------------------------------------------------------
    # Merge outer join
    # -----------------------------------------------------------------------
    try:
        df_merged = pd.merge(
            df_anterior,
            df_atual,
            on=chaves,
            how="outer",
            suffixes=(_SFX_ANT, _SFX_ATU),
            indicator=True,
        )
    except Exception as exc:  # noqa: BLE001
        resultado.erros.append(
            f"Erro ao fazer merge do bloco '{nome_bloco}': {exc}"
        )
        return resultado

    # -----------------------------------------------------------------------
    # Classificação de Status e preenchimento de Campos Alterados
    # -----------------------------------------------------------------------
    linhas_status: List[str] = []
    linhas_campos: List[str] = []

    for _, row in df_merged.iterrows():
        indicador = row["_merge"]

        if indicador == "right_only":
            linhas_status.append(STATUS_INCLUIDO)
            linhas_campos.append("")

        elif indicador == "left_only":
            linhas_status.append(STATUS_EXCLUIDO)
            linhas_campos.append("")

        else:  # both
            campos_divergentes = _detectar_divergencias(row, colunas_valor)
            if campos_divergentes:
                linhas_status.append(STATUS_ALTERADO)
                linhas_campos.append("; ".join(campos_divergentes))
            else:
                linhas_status.append(STATUS_SEM_ALT)
                linhas_campos.append("")

    df_merged[COL_STATUS] = linhas_status
    df_merged[COL_CAMPOS] = linhas_campos
    df_merged.drop(columns=["_merge"], inplace=True)

    # -----------------------------------------------------------------------
    # Reorganização das colunas: chaves | valor_ANT | valor_ATU | Status | Campos
    # -----------------------------------------------------------------------
    colunas_ordenadas = list(chaves)
    for col in colunas_valor:
        col_ant = f"{col}{_SFX_ANT}"
        col_atu = f"{col}{_SFX_ATU}"
        if col_ant in df_merged.columns:
            colunas_ordenadas.append(col_ant)
        if col_atu in df_merged.columns:
            colunas_ordenadas.append(col_atu)
        # Coluna não duplicada (estava só em um dos lados e manteve nome original)
        if col in df_merged.columns and col not in colunas_ordenadas:
            colunas_ordenadas.append(col)

    colunas_ordenadas += [COL_STATUS, COL_CAMPOS]
    # Garante que não há duplicatas nem colunas ausentes
    colunas_finais = [c for c in colunas_ordenadas if c in df_merged.columns]
    df_resultado = df_merged[colunas_finais].copy()

    # -----------------------------------------------------------------------
    # Estatísticas
    # -----------------------------------------------------------------------
    contagem = df_resultado[COL_STATUS].value_counts()
    resultado.qtd_incluidos = int(contagem.get(STATUS_INCLUIDO, 0))
    resultado.qtd_excluidos = int(contagem.get(STATUS_EXCLUIDO, 0))
    resultado.qtd_alterados = int(contagem.get(STATUS_ALTERADO, 0))
    resultado.qtd_sem_alteracao = int(contagem.get(STATUS_SEM_ALT, 0))
    resultado.df_resultado = df_resultado

    return resultado


# ---------------------------------------------------------------------------
# Funções auxiliares privadas
# ---------------------------------------------------------------------------

def _validar_chaves(
    chaves: List[str],
    df_ant: pd.DataFrame,
    df_atu: pd.DataFrame,
    nome_bloco: str,
    resultado: ResultadoBloco,
) -> List[str]:
    """
    Filtra as chaves configuradas removendo as que não existem em nenhum dos
    DataFrames. Emite aviso para cada coluna ausente.
    """
    todas_colunas_disponiveis = set(df_ant.columns) | set(df_atu.columns)
    chaves_validas = []
    for chave in chaves:
        if chave in todas_colunas_disponiveis:
            chaves_validas.append(chave)
        else:
            aviso = (
                f"Bloco '{nome_bloco}': coluna de chave '{chave}' não encontrada "
                "em nenhum dos arquivos. Coluna ignorada da chave."
            )
            logger.warning(aviso)
            resultado.erros.append(aviso)

    # Garante que chaves existam em ambos (merge exige isso)
    chaves_ant = [c for c in chaves_validas if c in df_ant.columns]
    chaves_atu = [c for c in chaves_validas if c in df_atu.columns]
    comuns = [c for c in chaves_validas if c in set(chaves_ant) & set(chaves_atu)]

    if len(comuns) < len(chaves_validas):
        faltantes = set(chaves_validas) - set(comuns)
        for f in faltantes:
            aviso = (
                f"Bloco '{nome_bloco}': chave '{f}' não está presente em ambos os "
                "arquivos. Removida da chave de comparação."
            )
            logger.warning(aviso)
            resultado.erros.append(aviso)

    return comuns


def _detectar_divergencias(row: pd.Series, colunas_valor: List[str]) -> List[str]:
    """
    Compara os valores _ANTERIOR e _ATUAL de cada coluna de valor.
    Retorna lista dos nomes originais das colunas que divergem.
    """
    divergentes = []
    for col in colunas_valor:
        val_ant = str(row.get(f"{col}{_SFX_ANT}", "")).strip()
        val_atu = str(row.get(f"{col}{_SFX_ATU}", "")).strip()
        # Trata NaN / nan como string vazia
        if val_ant.lower() in ("nan", "none", "null"):
            val_ant = ""
        if val_atu.lower() in ("nan", "none", "null"):
            val_atu = ""
        if val_ant != val_atu:
            divergentes.append(col)
    return divergentes


def _uniao_ordenada(lista_a: List[str], lista_b: List[str]) -> List[str]:
    """
    Une duas listas preservando a ordem de aparição (lista_a primeiro,
    depois elementos exclusivos de lista_b).
    """
    vistos = set()
    resultado = []
    for item in lista_a + lista_b:
        if item not in vistos:
            vistos.add(item)
            resultado.append(item)
    return resultado


def _construir_bloco_apenas_um_lado(
    df: pd.DataFrame,
    status: str,
    lado: str,  # "anterior" ou "atual"
) -> pd.DataFrame:
    """
    Constrói o DataFrame de resultado quando o bloco existe em apenas um
    dos arquivos. Duplica as colunas com sufixo _ANTERIOR ou _ATUAL e
    deixa o outro lado vazio.
    """
    df = df.copy()
    df_saida = pd.DataFrame()

    for col in df.columns:
        if lado == "anterior":
            df_saida[f"{col}{_SFX_ANT}"] = df[col]
            df_saida[f"{col}{_SFX_ATU}"] = ""
        else:
            df_saida[f"{col}{_SFX_ANT}"] = ""
            df_saida[f"{col}{_SFX_ATU}"] = df[col]

    df_saida[COL_STATUS] = status
    df_saida[COL_CAMPOS] = ""
    return df_saida
