"""
config.py — Configuração das chaves de identificação por bloco da ECF.

Para cada bloco (aba do Excel), define quais colunas formam a chave primária
que identifica unicamente uma linha no registro. Esses nomes devem corresponder
exatamente aos cabeçalhos das colunas no arquivo Excel de entrada.

Se um bloco não estiver mapeado aqui, o sistema usará a estratégia de chave
automática (colunas não-numéricas), emitindo um aviso no log.

Referência de campos: SPED ECF - Leiaute versão 10 (RFB 2024).
"""

# ---------------------------------------------------------------------------
# CHAVES POR BLOCO
# Formato: { "NOME_BLOCO": ["col_chave_1", "col_chave_2", ...] }
# ---------------------------------------------------------------------------
CHAVES_BLOCO: dict[str, list[str]] = {

    # -----------------------------------------------------------------------
    # BLOCO 0 — Abertura e Identificação
    # -----------------------------------------------------------------------
    "0000": [
        "DT_INI",
        "DT_FIN",
        "CNPJ",
    ],
    "0010": [
        "IDENT_MED",
    ],
    "0020": [
        "FORMA_TRIB_LUCRO",
        "ENT_IMUNE_ISENTA",
    ],
    "0030": [
        "CNPJ",
        "DT_INI",
    ],
    "0035": [
        "CNPJ_PART",
    ],

    # -----------------------------------------------------------------------
    # BLOCO C — Recuperação de Informações da ECD
    # -----------------------------------------------------------------------
    "C001": [
        "IND_MOV",
    ],
    "C010": [
        "CNPJ",
        "COD_SCP",
    ],
    "C020": [
        "COD_CONT",
        "CNPJ_CONT",
    ],

    # -----------------------------------------------------------------------
    # BLOCO E — Informações Econômicas
    # -----------------------------------------------------------------------
    "E001": [
        "IND_MOV",
    ],
    "E010": [
        "CNPJ",
    ],
    "E015": [
        "CNPJ_PART",
        "IND_REL",
    ],
    "E020": [
        "CNPJ",
        "IND_CNPJ",
    ],
    "E030": [
        "DT_INI",
        "DT_FIN",
    ],
    "E050": [
        "COD_CTA",
    ],
    "E155": [
        "COD_CTA",
        "DT_INI",
    ],
    "E200": [
        "UF",
    ],
    "E210": [
        "IND_MOV_ICMS",
    ],
    "E240": [
        "COD_CTA",
        "COD_MUN",
    ],
    "E250": [
        "COD_OR",
        "DT_VCTO",
    ],
    "E300": [
        "COD_MUN",
        "DT_INI",
    ],
    "E310": [
        "IND_MOV_ISS",
    ],
    "E312": [
        "COD_MUN",
    ],
    "E313": [
        "COD_CTA",
        "COD_MUN",
    ],
    "E316": [
        "COD_OR",
        "DT_VCTO",
    ],

    # -----------------------------------------------------------------------
    # BLOCO I — Informações Contábeis (FCONT / ECD)
    # -----------------------------------------------------------------------
    "I001": [
        "IND_MOV",
    ],
    "I010": [
        "IND_ECD",
    ],
    "I012": [
        "NUM_ORD",
        "ANO",
    ],
    "I015": [
        "COD_PLAN_REF",
    ],
    "I020": [
        "IND_CTA",
        "COD_CTA",
    ],
    "I030": [
        "COD_CTA",
        "DT_INI",
        "DT_FIN",
    ],
    "I050": [
        "COD_CTA",
    ],
    "I051": [
        "COD_CTA_REF",
        "COD_CTA",
    ],
    "I052": [
        "COD_CTA_REF",
    ],
    "I053": [
        "COD_CTA_REF",
        "CNPJ_ECD",
    ],
    "I075": [
        "COD_IND",
    ],
    "I100": [
        "COD_CTA",
        "DT_INI",
    ],
    "I150": [
        "DT_INI",
        "DT_FIN",
    ],
    "I155": [
        "COD_CTA",
        "DT_INI",
        "DT_FIN",
    ],
    "I200": [
        "COD_CTA",
    ],
    "I250": [
        "COD_CTA",
        "DT_INI",
    ],
    "I300": [
        "COD_CTA",
        "DT_INI",
    ],
    "I310": [
        "NUM_LANCTO",
    ],
    "I350": [
        "COD_CTA",
        "DT_INI",
    ],
    "I355": [
        "COD_CTA",
        "DT_INI",
        "COD_HIST",
    ],

    # -----------------------------------------------------------------------
    # BLOCO J — Demonstrações Contábeis
    # -----------------------------------------------------------------------
    "J001": [
        "IND_MOV",
    ],
    "J005": [
        "DT_INI",
        "DT_FIN",
        "COD_SIT",
    ],
    "J050": [
        "COD_CTA",
        "COD_CCUS",
    ],
    "J051": [
        "COD_CTA",
        "COD_CCUS",
    ],
    "J100": [
        "COD_CTA",
        "COD_CCUS",
    ],
    "J101": [
        "COD_AGL",
        "COD_CCUS",
    ],
    "J110": [
        "COD_CTA",
        "COD_CCUS",
    ],
    "J111": [
        "COD_CTA",
        "COD_CCUS",
    ],
    "J150": [
        "COD_CTA",
    ],
    "J200": [
        "COD_CTA",
    ],
    "J210": [
        "COD_CTA",
    ],
    "J215": [
        "COD_CTA",
    ],

    # -----------------------------------------------------------------------
    # BLOCO K — Ativos Intangíveis e Imobilizado
    # -----------------------------------------------------------------------
    "K001": [
        "IND_MOV",
    ],
    "K030": [
        "DT_INI",
        "DT_FIN",
    ],
    "K155": [
        "COD_CTA",
        "IND_TP_ATIVO",
    ],
    "K156": [
        "COD_CTA",
        "IND_TP_ATIVO",
        "CNPJ_EST",
    ],
    "K200": [
        "COD_CTA",
        "IND_TP_ATIVO",
    ],
    "K210": [
        "COD_CTA",
        "IND_TP_ATIVO",
    ],
    "K215": [
        "COD_CTA",
        "IND_TP_ATIVO",
        "CNPJ_EST",
    ],
    "K300": [
        "COD_CTA",
        "IND_TP_ATIVO",
    ],
    "K310": [
        "COD_CTA",
        "IND_TP_ATIVO",
        "CNPJ_EST",
    ],
    "K350": [
        "COD_CTA",
        "IND_TP_ATIVO",
    ],

    # -----------------------------------------------------------------------
    # BLOCO L — Cálculo do IRPJ/CSLL pelo Lucro Real
    # -----------------------------------------------------------------------
    "L001": [
        "IND_MOV",
    ],
    "L010": [
        "IND_MOV",
    ],
    "L020": [
        "COD_ORG_JULGADOR",
    ],
    "L030": [
        "COD_CTA",
    ],
    "L050": [
        "COD_CTA",
    ],
    "L100": [
        "COD_CTA",
    ],
    "L110": [
        "COD_CTA",
    ],
    "L111": [
        "COD_CTA",
    ],
    "L120": [
        "COD_CTA",
    ],
    "L200": [
        "COD_CTA",
    ],

    # -----------------------------------------------------------------------
    # BLOCO M — Cálculo da Base de Cálculo do IRPJ pelo Lucro Real
    # -----------------------------------------------------------------------
    "M001": [
        "IND_MOV",
    ],
    "M010": [
        "IND_FED",
    ],
    "M012": [
        "COD_IND_DEP",
        "CNPJ_DEP",
    ],
    "M015": [
        "COD_SCP",
    ],
    "M020": [
        "COD_PART",
        "CNPJ_PART",
    ],
    "M025": [
        "CNPJ_FONTE",
    ],
    "M026": [
        "COD_REC",
        "DT_VCTO",
    ],
    "M030": [
        "DT_INI",
        "DT_FIN",
    ],
    "M300": [
        "COD_CTA",
        "CNPJ_PARTE",
    ],
    "M305": [
        "COD_CTA",
    ],
    "M310": [
        "COD_CTA",
    ],
    "M350": [
        "COD_CTA",
    ],
    "M355": [
        "COD_CTA",
    ],
    "M360": [
        "COD_CTA",
    ],
    "M400": [
        "COD_IND_CONTRIB",
    ],
    "M410": [
        "COD_IND_JCP",
    ],
    "M500": [
        "COD_CTA",
    ],
    "M505": [
        "COD_CTA",
    ],

    # -----------------------------------------------------------------------
    # BLOCO N — Cálculo do CSLL (Regime de Apuração – Lucro Real)
    # -----------------------------------------------------------------------
    "N001": [
        "IND_MOV",
    ],
    "N010": [
        "IND_FED",
    ],
    "N020": [
        "COD_PART",
        "CNPJ_PART",
    ],
    "N025": [
        "CNPJ_FONTE",
    ],
    "N026": [
        "COD_REC",
        "DT_VCTO",
    ],
    "N030": [
        "DT_INI",
        "DT_FIN",
    ],
    "N500": [
        "COD_CTA",
    ],
    "N505": [
        "COD_CTA",
    ],
    "N600": [
        "COD_CTA",
    ],
    "N610": [
        "COD_CTA",
    ],
    "N620": [
        "COD_CTA",
    ],
    "N630": [
        "COD_CTA",
    ],
    "N640": [
        "COD_CTA",
    ],
    "N650": [
        "COD_CTA",
    ],
    "N660": [
        "COD_CTA",
    ],
    "N670": [
        "COD_CTA",
    ],

    # -----------------------------------------------------------------------
    # BLOCO P — Cálculo do IRPJ/CSLL – Lucro Presumido
    # -----------------------------------------------------------------------
    "P001": [
        "IND_MOV",
    ],
    "P010": [
        "CNPJ",
    ],
    "P020": [
        "DT_INI",
        "DT_FIN",
    ],
    "P100": [
        "COD_CTA",
    ],
    "P110": [
        "COD_CTA",
    ],
    "P130": [
        "COD_CTA",
    ],
    "P150": [
        "DT_INI",
        "DT_FIN",
    ],
    "P200": [
        "COD_CTA",
    ],
    "P230": [
        "DT_INI",
        "DT_FIN",
    ],

    # -----------------------------------------------------------------------
    # BLOCO Q — Informações de Pagamentos/Rendimentos a Beneficiários
    # -----------------------------------------------------------------------
    "Q001": [
        "IND_MOV",
    ],
    "Q100": [
        "CNPJ_CPF_BENEF",
        "COD_REC",
    ],

    # -----------------------------------------------------------------------
    # BLOCO T — Informações sobre Transações com Partes Relacionadas
    # -----------------------------------------------------------------------
    "T001": [
        "IND_MOV",
    ],
    "T020": [
        "CNPJ_PART",
        "IND_REL",
    ],
    "T025": [
        "CNPJ_PART",
        "IND_TRANSACAO",
    ],
    "T030": [
        "CNPJ_PART",
        "COD_CTA",
    ],
    "T035": [
        "CNPJ_PART",
        "COD_CTA",
        "IND_TRANSACAO",
    ],

    # -----------------------------------------------------------------------
    # BLOCO U — Impostos e Contribuições Devidos por Período
    # -----------------------------------------------------------------------
    "U001": [
        "IND_MOV",
    ],
    "U100": [
        "DT_INI",
        "DT_FIN",
        "IND_FORMA_APUR",
    ],

    # -----------------------------------------------------------------------
    # BLOCO X — Informações Econômicas dos Grupos Multinacionais (GloBE/CBCR)
    # -----------------------------------------------------------------------
    "X001": [
        "IND_MOV",
    ],
    "X010": [
        "CNPJ_MNE_DECL",
    ],
    "X020": [
        "CNPJ_CONST_FILIAL",
        "IND_REL_JURD",
    ],
    "X030": [
        "COD_PAIS",
        "IND_JURISD",
    ],
    "X035": [
        "CNPJ_CONST",
        "COD_PAIS",
    ],

    # -----------------------------------------------------------------------
    # BLOCO Y — Informações Gerais
    # -----------------------------------------------------------------------
    "Y001": [
        "IND_MOV",
    ],
    "Y520": [
        "CNPJ_PART",
        "COD_REL",
    ],
    "Y540": [
        "CNPJ_SUCEDIDA",
    ],
    "Y550": [
        "COD_REC",
        "NUM_PROCESSO",
    ],
    "Y570": [
        "CNPJ_COLIG",
        "COD_PAIS_COLIG",
    ],
    "Y590": [
        "CNPJ_PART",
        "IND_TIPO_REL",
    ],
    "Y600": [
        "CPF_RESP",
    ],
    "Y612": [
        "CPF_DIR",
    ],
    "Y620": [
        "CNPJ_ADM",
    ],
    "Y630": [
        "COD_PAIS",
    ],
    "Y640": [
        "CNPJ_PART",
        "COD_PAIS_PART",
    ],

    # -----------------------------------------------------------------------
    # BLOCO 9 — Encerramento
    # -----------------------------------------------------------------------
    "9001": [
        "IND_MOV",
    ],
    "9900": [
        "REG",
    ],
    "9990": [
        "QTD_LIN_9",
    ],
    "9999": [
        "QTD_LIN",
    ],
}

# ---------------------------------------------------------------------------
# PARÂMETROS DE NORMALIZAÇÃO
# ---------------------------------------------------------------------------

# Casas decimais para arredondamento de campos numéricos (evita falso-positivo)
DECIMAL_PLACES: int = 2

# Padrões de sufixo ou nome de coluna que indicam campo de valor/montante.
# Colunas que casem com esses padrões NÃO serão usadas como chave automática.
NUMERIC_COLUMN_HINTS: list[str] = [
    "VL_", "VAL_", "VALOR", "QTD_", "QUANT",
    "PERC_", "PERCENT", "ALIQ_", "TX_",
    "VR_", "MONTANTE", "BASE_",
]

# ---------------------------------------------------------------------------
# PARÂMETROS GERAIS
# ---------------------------------------------------------------------------

# Rótulo dos anos (usado nos sufixos das colunas e no log)
LABEL_ANTERIOR: str = "ANTERIOR"
LABEL_ATUAL: str = "ATUAL"

# Sufixo máximo para nome de aba do Excel (Excel limita 31 caracteres por aba)
SUFIXO_ABA: str = " Revisado"

# Cores de preenchimento (ARGB) para cada status
CORES_STATUS: dict[str, str] = {
    "Incluído":       "FFC6EFCE",   # verde claro
    "Excluído":       "FFFFC7CE",   # vermelho claro
    "Alterado":       "FFFFEB9C",   # amarelo
    "Sem alteração":  "FFFFFFFF",   # branco (sem cor)
}

# Cor de cabeçalho
COR_CABECALHO: str = "FF4472C4"     # azul escuro Microsoft
COR_CABECALHO_FONTE: str = "FFFFFFFF"  # branco

# Largura padrão de coluna (em unidades Excel)
LARGURA_COLUNA_PADRAO: float = 18.0
LARGURA_COLUNA_STATUS: float = 16.0
LARGURA_COLUNA_CAMPOS: float = 50.0
