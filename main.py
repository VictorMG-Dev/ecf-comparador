"""
main.py — Orquestrador da comparação de arquivos Excel de ECF.

Uso via linha de comando:
    python main.py --anterior ecf_2024.xlsx --atual ecf_2025.xlsx --saida revisao.xlsx
    python main.py --anterior ecf_2024.xlsx --atual ecf_2025.xlsx  (saída padrão)

Ou com variáveis configuradas diretamente no topo do arquivo.
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO RÁPIDA — altere aqui se não quiser usar argumentos de linha de cmd
# ---------------------------------------------------------------------------
ARQUIVO_ANTERIOR: str = "ecf_anterior.xlsx"    # caminho do Excel do ano anterior
ARQUIVO_ATUAL: str = "ecf_atual.xlsx"          # caminho do Excel do ano atual
ARQUIVO_SAIDA: str = "ecf_revisao.xlsx"        # caminho do Excel de saída
LABEL_ANTERIOR: str = "2024"                   # rótulo do período anterior
LABEL_ATUAL: str = "2025"                      # rótulo do período atual
# ---------------------------------------------------------------------------

from comparator import ResultadoBloco, comparar_bloco
from normalizer import normalizar_df
from reader import ler_excel
from writer import gerar_excel


def configurar_logging(verbose: bool = False) -> None:
    """Configura o sistema de log com formato legível e suporte a arquivo."""
    nivel = logging.DEBUG if verbose else logging.INFO
    formato = (
        "%(asctime)s  [%(levelname)-8s]  %(name)s — %(message)s"
    )
    logging.basicConfig(
        level=nivel,
        format=formato,
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("ecf_comparador.log", encoding="utf-8", mode="w"),
        ],
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="ecf_comparador",
        description=(
            "Compara dois arquivos Excel de ECF (Escrituração Contábil Fiscal) "
            "e gera um terceiro Excel com o resultado da revisão por bloco."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python main.py --anterior ecf_2024.xlsx --atual ecf_2025.xlsx
  python main.py --anterior ecf_2024.xlsx --atual ecf_2025.xlsx --saida revisao_final.xlsx --verbose
        """,
    )
    parser.add_argument(
        "--anterior",
        metavar="ARQUIVO",
        default=ARQUIVO_ANTERIOR,
        help=f"Caminho do Excel do ano anterior (padrão: {ARQUIVO_ANTERIOR})",
    )
    parser.add_argument(
        "--atual",
        metavar="ARQUIVO",
        default=ARQUIVO_ATUAL,
        help=f"Caminho do Excel do ano atual (padrão: {ARQUIVO_ATUAL})",
    )
    parser.add_argument(
        "--saida",
        metavar="ARQUIVO",
        default=ARQUIVO_SAIDA,
        help=f"Caminho do Excel de saída (padrão: {ARQUIVO_SAIDA})",
    )
    parser.add_argument(
        "--label-anterior",
        metavar="LABEL",
        default=LABEL_ANTERIOR,
        help=f"Rótulo do período anterior, ex: 2024 (padrão: {LABEL_ANTERIOR})",
    )
    parser.add_argument(
        "--label-atual",
        metavar="LABEL",
        default=LABEL_ATUAL,
        help=f"Rótulo do período atual, ex: 2025 (padrão: {LABEL_ATUAL})",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Ativa log detalhado (DEBUG)",
    )
    return parser.parse_args()


def _obter_todos_blocos(
    abas_anterior: Dict[str, pd.DataFrame],
    abas_atual: Dict[str, pd.DataFrame],
) -> List[str]:
    """Retorna a união ordenada dos blocos presentes nos dois arquivos."""
    blocos_ant = set(abas_anterior.keys())
    blocos_atu = set(abas_atual.keys())
    todos = sorted(blocos_ant | blocos_atu)
    return todos


def processar_blocos(
    abas_anterior: Dict[str, pd.DataFrame],
    abas_atual: Dict[str, pd.DataFrame],
    logger: logging.Logger,
) -> List[ResultadoBloco]:
    """
    Para cada bloco presente em pelo menos um dos arquivos:
      1. Normaliza os DataFrames.
      2. Compara os dois lados.
      3. Acumula os resultados.

    Erros em blocos individuais não interrompem os demais.
    """
    todos_blocos = _obter_todos_blocos(abas_anterior, abas_atual)
    resultados: List[ResultadoBloco] = []

    total = len(todos_blocos)
    logger.info("Total de blocos a processar: %d", total)

    for i, bloco in enumerate(todos_blocos, start=1):
        logger.info("[%d/%d] Processando bloco: %s", i, total, bloco)

        try:
            df_ant = abas_anterior.get(bloco)
            df_atu = abas_atual.get(bloco)

            # Normalização (pula se vazio)
            if df_ant is not None and not df_ant.empty:
                df_ant = normalizar_df(df_ant)

            if df_atu is not None and not df_atu.empty:
                df_atu = normalizar_df(df_atu)

            resultado = comparar_bloco(bloco, df_ant, df_atu)

            # Loga avisos e erros do bloco
            if resultado.aviso:
                logger.warning("[%s] %s", bloco, resultado.aviso)
            for erro in resultado.erros:
                logger.error("[%s] %s", bloco, erro)

            resultados.append(resultado)

        except Exception as exc:  # noqa: BLE001
            logger.exception(
                "Erro inesperado ao processar bloco '%s': %s. Bloco ignorado.",
                bloco, exc,
            )
            resultados.append(
                ResultadoBloco(
                    nome_bloco=bloco,
                    df_resultado=pd.DataFrame(),
                    chaves_usadas=[],
                    erros=[f"Erro inesperado: {exc}"],
                )
            )

    return resultados


def imprimir_relatorio_final(resultados: List[ResultadoBloco], logger: logging.Logger) -> None:
    """Imprime um relatório consolidado no console ao final da execução."""
    logger.info("")
    logger.info("=" * 70)
    logger.info("  RELATÓRIO FINAL DA REVISÃO ECF")
    logger.info("=" * 70)

    total_inc = total_exc = total_alt = total_sem = 0

    for r in resultados:
        total_inc += r.qtd_incluidos
        total_exc += r.qtd_excluidos
        total_alt += r.qtd_alterados
        total_sem += r.qtd_sem_alteracao

        tem_problema = r.qtd_incluidos + r.qtd_excluidos + r.qtd_alterados > 0 or r.erros

        if tem_problema or r.aviso:
            logger.info(
                "  %-10s | +%-5d | -%-5d | ~%-5d | =%-5d | %s",
                r.nome_bloco,
                r.qtd_incluidos,
                r.qtd_excluidos,
                r.qtd_alterados,
                r.qtd_sem_alteracao,
                r.aviso or "",
            )

    logger.info("-" * 70)
    logger.info(
        "  TOTAL      | +%-5d | -%-5d | ~%-5d | =%-5d",
        total_inc, total_exc, total_alt, total_sem,
    )
    logger.info("=" * 70)


def main() -> int:
    """
    Ponto de entrada principal.

    Retorna
    -------
    int
        0 em caso de sucesso, 1 em caso de erro crítico.
    """
    args = _parse_args()
    configurar_logging(verbose=args.verbose)
    logger = logging.getLogger("main")

    inicio = time.perf_counter()

    logger.info("=" * 70)
    logger.info("  COMPARADOR DE ECF — INICIANDO")
    logger.info("=" * 70)
    logger.info("  Arquivo anterior : %s", args.anterior)
    logger.info("  Arquivo atual    : %s", args.atual)
    logger.info("  Arquivo de saída : %s", args.saida)
    logger.info("  Período anterior : %s", args.label_anterior)
    logger.info("  Período atual    : %s", args.label_atual)
    logger.info("=" * 70)

    # --- Leitura ---
    try:
        logger.info("Lendo arquivo anterior...")
        abas_anterior = ler_excel(args.anterior)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("ERRO ao ler arquivo anterior: %s", exc)
        return 1

    try:
        logger.info("Lendo arquivo atual...")
        abas_atual = ler_excel(args.atual)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("ERRO ao ler arquivo atual: %s", exc)
        return 1

    # --- Processamento ---
    resultados = processar_blocos(abas_anterior, abas_atual, logger)

    if not resultados:
        logger.error("Nenhum bloco foi processado. Verifique os arquivos de entrada.")
        return 1

    # --- Geração do Excel ---
    try:
        logger.info("Gerando arquivo de saída: %s", args.saida)
        gerar_excel(
            resultados=resultados,
            caminho_saida=args.saida,
            label_anterior=args.label_anterior,
            label_atual=args.label_atual,
        )
    except PermissionError as exc:
        logger.error("ERRO de permissão ao salvar: %s", exc)
        return 1
    except Exception as exc:  # noqa: BLE001
        logger.exception("ERRO inesperado ao gerar Excel: %s", exc)
        return 1

    # --- Relatório final ---
    imprimir_relatorio_final(resultados, logger)

    duracao = time.perf_counter() - inicio
    logger.info("")
    logger.info("[CONCLUIDO] Concluido em %.2f segundos.", duracao)
    logger.info("[RESULTADO] Resultado salvo em: %s", Path(args.saida).resolve())
    logger.info("[LOG] Log completo salvo em: %s", Path("ecf_comparador.log").resolve())

    return 0


if __name__ == "__main__":
    sys.exit(main())
