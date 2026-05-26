"""
Parâmetros extraídos da aba Suporte do Simulador ODP — Privalia
Atualizado em: 25/05/2026
"""
from datetime import date

# ── Peças por palete por categoria real ───────────────────────────────────────
PCS_PALETE = {
    'Kids Trends':          421.0,
    'Kids Shoes':           135.0,
    'Kids Brands':          405.0,
    'Babycare':             451.0,
    'Shoes Brands':          97.0,
    'Shoes Comfort':        108.0,
    'Shoes Trends':         107.0,
    'Shoes Men':            109.0,
    'Sports':               396.0,
    'Fitness':              401.0,
    'Bodywear':             498.0,
    'Sul Trends':           396.0,
    'Sul Brands':           408.0,
    'Varejo Feminino':      336.0,
    'Varejo Masculino':     377.0,
    'Atacado Feminino':     418.0,
    'Atacado Masculino':    467.0,
    'Accessories & Beauty': 921.0,
    'Bed and Bath':          20.0,
    'Clearance':            150.0,
    'default':              200.0,
}

# ── Fatores de conversão IN (Check-In) e OUT (Check-Out/picking) ──────────────
# IN:  peças brutas → peças convertidas para Check-In
# OUT: peças brutas → peças convertidas para picking/saída
FAT_IN = {
    'Kids Trends':          0.7407,
    'Kids Shoes':           0.7407,
    'Kids Brands':          0.7407,
    'Babycare':             0.7407,
    'Shoes Brands':         2.0000,
    'Shoes Comfort':        2.0000,
    'Shoes Trends':         2.0000,
    'Shoes Men':            2.0000,
    'Sports':               1.0000,
    'Fitness':              1.0000,
    'Bodywear':             0.6667,
    'Sul Trends':           1.0000,
    'Sul Brands':           1.0000,
    'Varejo Feminino':      1.0000,
    'Varejo Masculino':     1.0000,
    'Atacado Feminino':     1.0000,
    'Atacado Masculino':    1.0000,
    'Accessories & Beauty': 1.6667,
    'Bed and Bath':         2.5000,
    'Clearance':            1.0000,
    'default':              1.0000,
}

FAT_OUT = {
    'Kids Trends':          1.0000,
    'Kids Shoes':           1.0000,
    'Kids Brands':          1.0000,
    'Babycare':             1.0000,
    'Shoes Brands':         2.5000,
    'Shoes Comfort':        2.5000,
    'Shoes Trends':         2.5000,
    'Shoes Men':            2.5000,
    'Sports':               1.0000,
    'Fitness':              1.0000,
    'Bodywear':             1.0000,
    'Sul Trends':           1.0000,
    'Sul Brands':           1.0000,
    'Varejo Feminino':      1.0000,
    'Varejo Masculino':     1.0000,
    'Atacado Feminino':     1.0000,
    'Atacado Masculino':    1.0000,
    'Accessories & Beauty': 2.6316,
    'Bed and Bath':         2.6316,
    'Clearance':            1.0000,
    'default':              1.0000,
}

# ── De-para categoria → setor de gerência ────────────────────────────────────
DEPARA_SETOR = {
    'Kids Trends':          'Kids',
    'Kids Shoes':           'Kids',
    'Kids Brands':          'Kids',
    'Babycare':             'Kids',
    'Shoes Brands':         'Shoes e Co',
    'Shoes Comfort':        'Shoes e Co',
    'Shoes Trends':         'Shoes e Co',
    'Shoes Men':            'Shoes e Co',
    'Sports':               'Shoes e Co',
    'Fitness':              'Fashion',
    'Bodywear':             'Fashion',
    'Sul Trends':           'Fashion',
    'Sul Brands':           'Fashion',
    'Varejo Feminino':      'Fashion',
    'Varejo Masculino':     'Fashion',
    'Atacado Feminino':     'Fashion',
    'Atacado Masculino':    'Fashion',
    'Accessories & Beauty': 'Shoes e Co',
    'Bed and Bath':         'Home Decor',
    'Clearance':            'Clearance',
}

# ── Feriados 2026 (Nacional + Extrema) ───────────────────────────────────────
# Editável: adicionar/remover datas conforme necessário
FERIADOS_2026 = {
    date(2026,  1,  1),  # Ano Novo
    date(2026,  1,  2),  # Emenda
    date(2026,  2, 16),  # Carnaval
    date(2026,  2, 17),  # Carnaval
    date(2026,  4,  3),  # Sexta-feira Santa
    date(2026,  4, 17),  # Tiradentes (antecipado CD)
    date(2026,  5,  1),  # Dia do Trabalho
    date(2026,  5, 22),  # Aniversário Extrema
    date(2026,  9,  7),  # Independência
    date(2026, 10, 12),  # N.S. Aparecida
    date(2026, 11,  2),  # Finados
    date(2026, 11, 15),  # Proclamação da República
    date(2026, 12, 25),  # Natal
}

FERIADOS_2025 = {
    date(2025,  1,  1),
    date(2025,  3,  4),  # Carnaval
    date(2025,  4, 18),  # Sexta Santa
    date(2025,  4, 21),  # Tiradentes
    date(2025,  5,  1),
    date(2025,  6, 19),  # Corpus Christi
    date(2025,  8, 27),  # Municipal
    date(2025,  9,  7),
    date(2025,  9, 15),  # Aniversário Extrema
    date(2025, 10, 12),
    date(2025, 11,  2),
    date(2025, 11, 15),
    date(2025, 12, 24),
    date(2025, 12, 25),
    date(2025, 12, 26),
    date(2025, 12, 31),
}

FERIADOS = FERIADOS_2025 | FERIADOS_2026

def get_pcs_palete(cat: str) -> float:
    return PCS_PALETE.get(cat, PCS_PALETE['default'])

def get_fat_in(cat: str) -> float:
    return FAT_IN.get(cat, FAT_IN['default'])

def get_fat_out(cat: str) -> float:
    return FAT_OUT.get(cat, FAT_OUT['default'])

def get_setor(cat: str) -> str:
    return DEPARA_SETOR.get(cat, 'Outros')
