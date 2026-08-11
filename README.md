# ECF Comparador

Automação Python profissional para comparar dois arquivos Excel de **ECF — Escrituração Contábil Fiscal**, gerando um terceiro Excel com o resultado detalhado da revisão por bloco.

---

## Estrutura do Projeto

```
ecf_comparador/
├── config.py               # Chaves de identificação por bloco + parâmetros de formatação
├── reader.py               # Leitura dos arquivos Excel de entrada
├── normalizer.py           # Normalização dos dados (strip, uppercase, conversão numérica)
├── comparator.py           # Lógica de comparação (merge outer join, classificação de Status)
├── writer.py               # Geração do Excel final com formatação openpyxl
├── main.py                 # Orquestrador: CLI + log + relatório final
├── generate_test_data.py   # Gerador de dados de teste realistas
├── requirements.txt        # Dependências Python
└── ecf_comparador.log      # Log gerado automaticamente a cada execução
```

---

## Instalação

### Pré-requisitos

- [uv](https://github.com/astral-sh/uv) (instalador rápido de ambientes Python)
- Windows: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`

### Criar ambiente e instalar dependências

```powershell
# No diretório do projeto
uv venv .venv
uv pip install -r requirements.txt --python .venv/Scripts/python.exe
```

---

## Uso

### Modo padrão (variáveis no topo de main.py)

Edite as variáveis no topo de `main.py`:

```python
ARQUIVO_ANTERIOR = "ecf_2024.xlsx"
ARQUIVO_ATUAL    = "ecf_2025.xlsx"
ARQUIVO_SAIDA    = "ecf_revisao.xlsx"
LABEL_ANTERIOR   = "2024"
LABEL_ATUAL      = "2025"
```

Depois execute:

```powershell
.venv\Scripts\python.exe main.py
```

### Modo linha de comando (CLI)

```powershell
.venv\Scripts\python.exe main.py `
    --anterior ecf_2024.xlsx `
    --atual    ecf_2025.xlsx `
    --saida    ecf_revisao.xlsx `
    --label-anterior 2024 `
    --label-atual    2025 `
    --verbose
```

### Ajuda

```powershell
.venv\Scripts\python.exe main.py --help
```

---

## Dados de Teste

Para gerar arquivos de teste realistas com 6 blocos e todos os cenários (linhas incluídas, excluídas, alteradas, sem alteração, bloco novo e bloco removido):

```powershell
.venv\Scripts\python.exe generate_test_data.py
# Gera: ecf_anterior.xlsx  e  ecf_atual.xlsx
```

---

## Excel de Saída

O arquivo gerado contém:

| Aba | Conteúdo |
|-----|----------|
| **Resumo** | Tabela consolidada: bloco, chaves, contagens por status, avisos |
| **`<BLOCO> Revisado`** | Uma aba por bloco, com todas as colunas `_ANTERIOR` e `_ATUAL` |

### Colunas em cada aba Revisado

| Coluna | Descrição |
|--------|-----------|
| *Colunas-chave* | Campos que identificam a linha (sem sufixo) |
| `CAMPO_2024` / `CAMPO_2025` | Valor de cada campo nos dois anos |
| **Status** | `Incluído` / `Excluído` / `Alterado` / `Sem alteração` |
| **Campos Alterados** | Lista dos campos que mudaram (quando Status = Alterado) |

### Código de cores

| Status | Cor |
|--------|-----|
| Incluído | 🟢 Verde claro |
| Excluído | 🔴 Vermelho claro |
| Alterado | 🟡 Amarelo |
| Sem alteração | Branco |

---

## Configuração das Chaves por Bloco (`config.py`)

Cada bloco ECF tem sua própria chave de identificação de linha. O arquivo `config.py` já vem com **todos os blocos comuns** mapeados (0000, I051, J100, M300, Y600, Q100, e dezenas de outros).

Para adicionar um bloco customizado:

```python
# config.py
CHAVES_BLOCO["MEU_BLOCO"] = ["COD_CTA", "DT_INI"]
```

### Fallback automático

Se um bloco **não estiver** no `config.py`, o sistema:
1. Detecta automaticamente as colunas-chave (exclui colunas numéricas/valor).
2. Emite um **aviso no log** e na aba Resumo com a chave usada.
3. Marca a aba com fundo laranja para chamar atenção.

---

## Normalização

Antes de comparar, os dados são normalizados automaticamente:

- **Texto**: `strip()` + `UPPERCASE`
- **Números**: converte separador decimal (vírgula → ponto), arredonda para 2 casas decimais
- **Formato brasileiro**: `1.234,56` → `1234.56`

Isso evita falso-positivos por diferença de formatação.

---

## Tratamento de Erros

| Situação | Comportamento |
|----------|---------------|
| Arquivo não encontrado | Erro imediato com mensagem clara |
| Aba vazia | Aviso no log, bloco ignorado |
| Coluna de chave ausente | Aviso, coluna removida da chave |
| Bloco somente em um arquivo | Sinalizado como "NOVO" ou "REMOVIDO" |
| Erro inesperado em um bloco | Log de exceção, execução continua nos demais |
| Arquivo de saída aberto | Erro de permissão com instrução clara |

---

## Resultado de Execução (exemplo)

```
======================================================================
  COMPARADOR DE ECF — INICIANDO
======================================================================
  Arquivo anterior : ecf_anterior.xlsx
  Arquivo atual    : ecf_atual.xlsx
  Arquivo de saída : ecf_revisao.xlsx
  Período anterior : 2024
  Período atual    : 2025
======================================================================
...
  0000       | +1     | -1     | ~0     | =0     |
  I051       | +5     | -5     | ~2     | =23    |
  J100       | +0     | -0     | ~8     | =17    |
  M300       | +2     | -1     | ~2     | =15    |
  Q100       | +5     | -0     | ~0     | =0     | Bloco NOVO — presente apenas no arquivo atual.
  Y600       | +0     | -1     | ~0     | =0     | Bloco REMOVIDO — presente apenas no arquivo anterior.
----------------------------------------------------------------------
  TOTAL      | +13    | -8     | ~12    | =55
======================================================================
[CONCLUIDO] Concluido em 0.26 segundos.
```

---

## Dependências

| Pacote | Versão mínima | Uso |
|--------|--------------|-----|
| pandas | 2.0.0 | Leitura e manipulação de DataFrames |
| openpyxl | 3.1.0 | Escrita e formatação do Excel final |
| xlrd | 2.0.0 | Leitura de arquivos .xls legados |

---

## Notas

- **Limite de nome de aba**: Excel aceita no máximo 31 caracteres. Nomes longos são truncados automaticamente com aviso.
- **Log completo**: Cada execução gera `ecf_comparador.log` com todos os eventos (modo `--verbose` inclui DEBUG).
- **Extensões suportadas na entrada**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`
