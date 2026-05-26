import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
import io

st.set_page_config(
    page_title="Simulador ODP — Privalia",
    page_icon="📦",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.block-container{padding:1.2rem 2rem 2rem;}
[data-baseweb="tab-list"]{gap:0;border-bottom:1px solid #e5e7eb;}
[data-baseweb="tab"]{font-size:13px;font-weight:500;padding:10px 18px;color:#6b7280;}
[aria-selected="true"]{color:#111827!important;border-bottom:2px solid #111827!important;}
[data-testid="metric-container"]{background:#f9fafb;border-radius:10px;padding:14px 18px;border:1px solid #f3f4f6;}
[data-testid="stMetricValue"]{font-size:22px!important;font-weight:600;}
[data-testid="stMetricLabel"]{font-size:11px!important;color:#6b7280;}
div[data-testid="stForm"]{border:1px solid #f3f4f6;border-radius:10px;padding:16px;}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_br(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors='coerce')
    return pd.to_numeric(
        s.astype(str).str.replace('.','',regex=False).str.replace(',','.',regex=False),
        errors='coerce')

def proximos_dias_uteis(data_base, n_dias, feriados_set):
    step = 1 if n_dias >= 0 else -1
    restante = abs(n_dias)
    d = data_base
    while restante > 0:
        d += timedelta(days=step)
        if d.weekday() < 5 and d not in feriados_set:
            restante -= 1
    return d

def calcular_datas_blocado(inicio, fim, feriados_set, dn_manual=None, do_manual=None):
    dn = dn_manual if dn_manual else proximos_dias_uteis(inicio, -2, feriados_set)
    do = do_manual if do_manual else proximos_dias_uteis(fim,    3, feriados_set)
    return dn, do

# ═══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO DE DADOS
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def carregar_calendario(f):
    df = pd.read_excel(f)
    rename = {
        'Id externo de Campanha Sacarino': 'id_campanha',
        'Nome da campanha':               'campanha',
        'Data de início':                 'data_inicio',
        'Data de término':                'data_fim',
        'Webdays':                        'webdays',
        'Status':                         'status',
        'Centro de distibuição':          'cd',
        'Categoria':                      'categoria',
        'Sector Calendar':                'setor',
        'Previsão de venda peças':        'previsao_pecas',
        'Estoque Total':                  'estoque_total',
        'Modelo de negócio':              'modelo_negocio',
        'Gerência':                       'gerencia',
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    df['data_inicio'] = pd.to_datetime(df['data_inicio'], dayfirst=True, errors='coerce')
    df['data_fim']    = pd.to_datetime(df['data_fim'],    dayfirst=True, errors='coerce')
    cols = ['id_campanha','campanha','data_inicio','data_fim','webdays','status',
            'cd','categoria','setor','previsao_pecas','estoque_total',
            'modelo_negocio','gerencia']
    df = df[[c for c in cols if c in df.columns]]
    # filtro ODP + dedup por ID (ignora múltiplos CNPJs)
    if 'modelo_negocio' in df.columns:
        df = df[df['modelo_negocio'].str.upper().str.strip() == 'ODP']
    df = df.dropna(subset=['data_inicio','data_fim'])
    df = df.drop_duplicates(subset=['id_campanha'], keep='first').reset_index(drop=True)
    df['webdays']       = pd.to_numeric(df['webdays'],       errors='coerce').fillna(1).clip(lower=1).astype(int)
    df['previsao_pecas']= pd.to_numeric(df['previsao_pecas'],errors='coerce').fillna(0)
    df['estoque_total'] = pd.to_numeric(df['estoque_total'], errors='coerce').fillna(0)
    return df

@st.cache_data(show_spinner=False)
def carregar_vendas(f):
    # suporta CSV (sep=;) e Excel
    try:
        df = pd.read_csv(f, sep=';', dtype=str, encoding='utf-8')
    except Exception:
        try:
            df = pd.read_csv(f, sep=';', dtype=str, encoding='latin-1')
        except Exception:
            df = pd.read_excel(f, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    rename = {'ID':'id_campanha','Dia_Click':'data_venda',
              'Items':'pecas_vendidas','Orders':'pedidos','Revenue':'receita'}
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    if 'id_campanha' in df.columns:
        df['id_campanha'] = pd.to_numeric(
            df['id_campanha'].astype(str).str.replace('.','',regex=False), errors='coerce')
    if 'data_venda' in df.columns:
        df['data_venda'] = pd.to_datetime(df['data_venda'], dayfirst=True, errors='coerce')
    for c in ['pecas_vendidas','pedidos','receita']:
        if c in df.columns:
            df[c] = parse_br(df[c])
    return df

@st.cache_data(show_spinner=False)
def carregar_odp_vertical(f):
    """Lê o export da aba Recebimentos e Aderência (ODP Vertical).
    Estrutura real: linha 1 = títulos de seção, linha 2 = cabeçalhos reais.
    Colunas por posição: I(9)=ID, J(10)=Nome, K(11)=Cat, L(12)=Peças,
                         M(13)=PeçasConv, N(14)=Status, O(15)=Data
    """
    # tenta ler com header na linha 2 (índice 1)
    df_raw = pd.read_excel(f, header=None)
    if len(df_raw) < 3:
        return pd.DataFrame()

    # encontra a linha de cabeçalho buscando 'ID campanha' ou 'Nome'
    header_row = 1  # padrão: linha 2
    for r in range(min(5, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[r].tolist()]
        if any('campanha' in v.lower() or 'nome' in v.lower() for v in row_vals):
            header_row = r
            break

    data_rows = df_raw.iloc[header_row + 1:].reset_index(drop=True)

    # mapeia por posição (colunas fixas da aba Recebimentos e Aderência)
    # col 0=Depara, 1=DataConsumo, 2=Setor, 3=Mês, 4=Semana, 5=Rec,
    # 6=Concat, 7=Negoc, 8=ID, 9=Nome, 10=Cat, 11=Pecas, 12=PecasConv, 13=Status, 14=Data
    col_map = {8: 'id_campanha', 9: 'campanha', 10: 'categoria',
               11: 'pecas', 12: 'pecas_convertidas', 13: 'status_entrada', 14: 'data_entrada'}

    df = pd.DataFrame()
    for col_idx, nome in col_map.items():
        if col_idx < len(data_rows.columns):
            df[nome] = data_rows.iloc[:, col_idx]

    if 'id_campanha' in df.columns:
        df['id_campanha'] = pd.to_numeric(df['id_campanha'], errors='coerce')
    if 'data_entrada' in df.columns:
        df['data_entrada'] = pd.to_datetime(df['data_entrada'], errors='coerce')
    for c in ['pecas','pecas_convertidas']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    return df.dropna(subset=['id_campanha']).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def carregar_parametros(f_params, f_curvas):
    """Lê parametros_odp.xlsx e curvas_venda_odp.xlsx."""
    params = {}

    if f_params is not None:
        xp = pd.ExcelFile(f_params)

        # Pcs/Palete
        if 'Pcs_Palete' in xp.sheet_names:
            df = xp.parse('Pcs_Palete', skiprows=1)
            df.columns = ['categoria','pcs_palete'] if len(df.columns) >= 2 else df.columns
            df = df.dropna(subset=['categoria'])
            df['pcs_palete'] = pd.to_numeric(df['pcs_palete'], errors='coerce')
            params['pcs_palete'] = df.dropna(subset=['pcs_palete']).drop_duplicates('categoria')

        # Fatores IN/OUT
        if 'Fatores_Conversao' in xp.sheet_names:
            df = xp.parse('Fatores_Conversao', skiprows=2)
            df.columns = ['grupo','fator_in','fator_out'] if len(df.columns) >= 3 else df.columns
            df = df.dropna(subset=['grupo'])
            for c in ['fator_in','fator_out']:
                df[c] = pd.to_numeric(df[c], errors='coerce')
            params['fatores'] = df.dropna(subset=['fator_in']).drop_duplicates('grupo')

        # Feriados
        if 'Feriados' in xp.sheet_names:
            df = xp.parse('Feriados', skiprows=1)
            df.columns = ['data','nome','local'] if len(df.columns) >= 3 else df.columns
            df['data'] = pd.to_datetime(df['data'], errors='coerce')
            params['feriados'] = df.dropna(subset=['data'])

        # De-para categoria → grupo fator
        if 'Depara_Categoria_Fator' in xp.sheet_names:
            df = xp.parse('Depara_Categoria_Fator', skiprows=2)
            df.columns = ['categoria','grupo'] if len(df.columns) >= 2 else df.columns
            df = df.dropna(subset=['categoria','grupo'])
            params['depara'] = df.drop_duplicates('categoria')

    # Curvas
    if f_curvas is not None:
        df = pd.read_excel(f_curvas, skiprows=1)
        # cols: Categoria, Webdays, Soma, d1, d2, ...
        if len(df.columns) >= 4:
            df.columns = (['categoria','webdays','soma'] +
                          [f'd{i}' for i in range(1, len(df.columns)-2)])
            df = df.dropna(subset=['categoria','webdays'])
            df['webdays'] = pd.to_numeric(df['webdays'], errors='coerce')
            for c in [c for c in df.columns if c.startswith('d')]:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            params['curvas'] = df.dropna(subset=['webdays'])

    return params

# ═══════════════════════════════════════════════════════════════════════════════
# PARÂMETROS — acesso com fallbacks embutidos
# ═══════════════════════════════════════════════════════════════════════════════

FATORES_DEFAULT = {
    'Fashion':           (1.0,    1.0),
    'Shoes':             (2.0,    2.5),
    'Kids':              (0.7407, 1.0),
    'Baby':              (0.7407, 1.0),
    'Sport':             (1.0,    1.0),
    'Underwear':         (0.6667, 1.0),
    'Beachwear':         (0.6667, 1.0),
    'Home & Decor':      (2.5,    2.6316),
    'Acessories':        (1.6667, 2.6316),
    'Beauty and Wellness':(2.5,   2.6316),
    'Eyewear':           (1.6667, 2.6316),
    'Bags':              (2.0,    2.6316),
    'Clearance':         (1.0,    1.0),
}

PCS_DEFAULT = {
    'Kids Trends': 421, 'Kids Shoes': 135, 'Kids Brands': 405,
    'Babycare': 451, 'Shoes Brands': 97, 'Shoes Comfort': 108,
    'Shoes Trends': 107, 'Shoes Men': 109, 'Sports': 396,
    'Fitness': 401, 'Bodywear': 498, 'Sul Trends': 396,
    'Sul Brands': 408, 'Varejo Feminino': 336, 'Varejo Masculino': 377,
    'Atacado Feminino': 418, 'Atacado Masculino': 467,
    'Accessories & Beauty': 921, 'Bed and Bath': 20, 'Clearance': 150,
}

DEPARA_DEFAULT = {
    'Kids Trends': 'Kids', 'Kids Shoes': 'Kids', 'Kids Brands': 'Kids',
    'Babycare': 'Baby', 'Shoes Brands': 'Shoes', 'Shoes Comfort': 'Shoes',
    'Shoes Trends': 'Shoes', 'Shoes Men': 'Shoes', 'Sports': 'Sport',
    'Fitness': 'Fashion', 'Bodywear': 'Underwear', 'Sul Trends': 'Fashion',
    'Sul Brands': 'Fashion', 'Varejo Feminino': 'Fashion', 'Varejo Masculino': 'Fashion',
    'Atacado Feminino': 'Fashion', 'Atacado Masculino': 'Fashion',
    'Accessories & Beauty': 'Acessories', 'Bed and Bath': 'Home & Decor',
    'Clearance': 'Clearance',
}

def get_feriados_set(params, cd='Extrema'):
    if params and 'feriados' in params:
        df = params['feriados']
        locais = ['Nacional']
        if 'extrema' in str(cd).lower(): locais += ['Extrema','SP']
        if 'jandira' in str(cd).lower():  locais += ['Jandira','SP']
        mask = df['local'].str.strip().isin(locais)
        return set(df.loc[mask,'data'].dt.date.tolist())
    # fallback embutido (2025-2026)
    from datetime import date as _date
    return {
        _date(2025,1,1), _date(2025,3,4), _date(2025,4,18), _date(2025,4,21),
        _date(2025,5,1), _date(2025,6,19), _date(2025,8,27), _date(2025,9,7),
        _date(2025,9,15), _date(2025,10,12), _date(2025,11,2), _date(2025,11,15),
        _date(2025,12,24), _date(2025,12,25), _date(2025,12,26), _date(2025,12,31),
        _date(2026,1,1), _date(2026,1,2), _date(2026,2,16), _date(2026,2,17),
        _date(2026,4,3), _date(2026,4,17), _date(2026,5,1), _date(2026,5,22),
        _date(2026,9,7), _date(2026,10,12), _date(2026,11,2), _date(2026,11,15),
        _date(2026,12,25),
    }

def get_pcs_palete(cat, params):
    if params and 'pcs_palete' in params:
        df = params['pcs_palete']
        r = df[df['categoria'].str.strip() == cat]
        if not r.empty:
            return float(r.iloc[0]['pcs_palete'])
    return float(PCS_DEFAULT.get(cat, 200))

def get_fator(cat, tipo, params):
    """Busca fator IN ou OUT para a categoria real, via depara."""
    grupo = None
    if params and 'depara' in params:
        df = params['depara']
        r = df[df['categoria'].str.strip() == cat]
        if not r.empty:
            grupo = r.iloc[0]['grupo']
    if grupo is None:
        grupo = DEPARA_DEFAULT.get(cat, 'Fashion')

    if params and 'fatores' in params:
        df = params['fatores']
        r = df[df['grupo'].str.strip() == grupo]
        if not r.empty:
            col = 'fator_in' if tipo == 'in' else 'fator_out'
            return float(r.iloc[0][col])

    fi, fo = FATORES_DEFAULT.get(grupo, (1.0, 1.0))
    return fi if tipo == 'in' else fo

def get_curva(cat, webdays, params):
    wd = int(webdays)
    if params and 'curvas' in params:
        df = params['curvas']
        r = df[(df['categoria'].str.strip() == cat) & (df['webdays'] == wd)]
        if r.empty:
            # fallback Standard
            r = df[(df['categoria'].str.strip() == 'Standard') & (df['webdays'] == wd)]
        if not r.empty:
            cols_d = [f'd{i}' for i in range(1, wd+1) if f'd{i}' in r.columns]
            vals = [float(r.iloc[0][c]) for c in cols_d]
            total = sum(vals)
            return [v/total for v in vals] if total > 0 else [1.0/wd]*wd
    return [1.0/wd]*wd

# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR DE CÁLCULO
# ═══════════════════════════════════════════════════════════════════════════════

def calcular_saidas(calendario, vendas, data_hoje, params, ajustes=None):
    if ajustes is None: ajustes = {}
    linhas = []
    for _, camp in calendario.iterrows():
        cid = camp['id_campanha']
        aj  = ajustes.get(cid, {})
        inicio   = aj.get('data_inicio', camp['data_inicio'].date())
        fim      = aj.get('data_fim',    camp['data_fim'].date())
        webdays  = max(int((fim - inicio).days + 1), 1)
        prev_tot = aj.get('previsao_pecas', camp['previsao_pecas'])
        cat      = camp['categoria']
        fator_out = get_fator(cat, 'out', params)
        curva     = get_curva(cat, webdays, params)

        vendas_camp = pd.DataFrame()
        if vendas is not None and 'id_campanha' in vendas.columns:
            vendas_camp = vendas[vendas['id_campanha'] == cid].copy()

        datas = [inicio + timedelta(days=i) for i in range(webdays)]
        for idx, d in enumerate(datas):
            peso     = curva[idx] if idx < len(curva) else (1.0/webdays)
            prev_dia = prev_tot * peso

            if d < data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas_bruto = float(vd['pecas_vendidas'].sum()) if not vd.empty else prev_dia
                else:
                    pecas_bruto = prev_dia
                tipo = 'realizado'
            elif d == data_hoje:
                if not vendas_camp.empty and 'data_venda' in vendas_camp.columns:
                    vd = vendas_camp[vendas_camp['data_venda'].dt.date == d]
                    pecas_bruto = float(vd['pecas_vendidas'].sum()) if not vd.empty else prev_dia
                else:
                    pecas_bruto = prev_dia
                tipo = 'reforecast'
            else:
                pecas_bruto = prev_dia
                tipo = 'previsao'

            linhas.append({
                'id_campanha': cid,
                'campanha':    camp['campanha'],
                'data':        d,
                'gerencia':    camp.get('gerencia',''),
                'categoria':   cat,
                'cd':          camp.get('cd',''),
                'pecas_bruto': round(pecas_bruto, 1),
                'pecas_conv':  round(pecas_bruto * fator_out, 1),
                'tipo':        tipo,
            })
    return pd.DataFrame(linhas)

def calcular_blocado(calendario, data_hoje, params, ajustes=None):
    if ajustes is None: ajustes = {}
    linhas = []
    for _, camp in calendario.iterrows():
        cid   = camp['id_campanha']
        aj    = ajustes.get(cid, {})
        cd    = camp.get('cd','Extrema')
        cat   = camp['categoria']
        feriados = get_feriados_set(params, cd)

        inicio = aj.get('data_inicio', camp['data_inicio'].date())
        fim    = aj.get('data_fim',    camp['data_fim'].date())
        est    = aj.get('estoque_total', camp['estoque_total'])

        dn_manual = aj.get('dn', None)
        do_manual = aj.get('do', None)
        dn, do = calcular_datas_blocado(inicio, fim, feriados, dn_manual, do_manual)

        fator_in = get_fator(cat, 'in', params)
        ppp      = get_pcs_palete(cat, params)
        est_conv = est * fator_in
        pallets  = est_conv / ppp if ppp > 0 else 0

        if dn <= data_hoje <= do:     status = 'blocado'
        elif data_hoje < dn:          status = 'aguardando'
        else:                         status = 'liberado'

        linhas.append({
            'id_campanha':  cid,
            'campanha':     camp['campanha'],
            'categoria':    cat,
            'cd':           cd,
            'gerencia':     camp.get('gerencia',''),
            'data_inicio':  inicio,
            'data_fim':     fim,
            'dn':           dn,
            'do':           do,
            'estoque_pecas':round(est, 0),
            'estoque_conv': round(est_conv, 0),
            'pallets':      round(pallets, 2),
            'status':       status,
        })

    df = pd.DataFrame(linhas)
    if df.empty:
        return df, pd.DataFrame()

    min_d = df['dn'].min()
    max_d = df['do'].max()
    serie = []
    for d in pd.date_range(min_d, max_d).date:
        ativas = df[(df['dn'] <= d) & (df['do'] >= d)]
        serie.append({
            'data':       d,
            'pecas':      ativas['estoque_pecas'].sum(),
            'pecas_conv': ativas['estoque_conv'].sum(),
            'pallets':    round(ativas['pallets'].sum(), 1),
            'campanhas':  len(ativas),
        })
    return df, pd.DataFrame(serie)

# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════

if 'ajustes' not in st.session_state:
    st.session_state['ajustes'] = {}

# ═══════════════════════════════════════════════════════════════════════════════
# INTERFACE — ABAS
# ═══════════════════════════════════════════════════════════════════════════════

st.title("Simulador ODP — Privalia")

aba_dados, aba_entradas, aba_saidas, aba_blocado, aba_rfcst, aba_sim, aba_params, aba_export = st.tabs([
    "Dados", "Entradas (ODP Vertical)", "Saídas", "Blocado",
    "Reforecast", "Simulações", "Parâmetros", "Exportar"
])

# ═══════════════════════════════════════════════════════════════════════════════
# ABA DADOS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_dados:
    st.subheader("Carregar dados")
    c1, c2, c3 = st.columns(3)
    with c1:
        arq_cal  = st.file_uploader("Calendário Salesforce (.xlsx)",
                                     type=["xlsx","xls"], key="up_cal")
    with c2:
        arq_vend = st.file_uploader("Base Vendas DBeaver (.csv ou .xlsx)",
                                     type=["csv","xlsx","xls"], key="up_vend")
    with c3:
        arq_odp  = st.file_uploader("ODP Vertical — Recebimentos (.xlsx)",
                                     type=["xlsx","xls"], key="up_odp")

    st.divider()
    cp1, cp2 = st.columns(2)
    with cp1:
        arq_params = st.file_uploader("Parâmetros ODP (.xlsx) — opcional",
                                       type=["xlsx"], key="up_params",
                                       help="parametros_odp.xlsx com Pcs/Palete, Fatores, Feriados e De-para")
    with cp2:
        arq_curvas = st.file_uploader("Curvas de Venda (.xlsx) — opcional",
                                       type=["xlsx"], key="up_curvas",
                                       help="curvas_venda_odp.xlsx com distribuição de venda por categoria × webdays")

    st.divider()
    ca, cb = st.columns(2)
    with ca:
        data_hoje = st.date_input("Data de referência (hoje)", value=date.today())
    with cb:
        sel_cd_global = st.selectbox("Centro de distribuição",
                                      ["Extrema e Jandira","Extrema","Jandira"])

    # Carregamento e validação
    if arq_cal:
        try:
            cal = carregar_calendario(arq_cal)
            st.success(f"✅ Calendário: **{len(cal)}** campanhas ODP")
            with st.expander("Prévia"):
                st.dataframe(cal.head(8), use_container_width=True, hide_index=True)
            st.session_state['cal'] = cal
        except Exception as e:
            st.error(f"Erro no Calendário: {e}")

    if arq_vend:
        try:
            vendas = carregar_vendas(arq_vend)
            st.success(f"✅ Vendas: **{len(vendas):,}** registros")
            st.session_state['vendas'] = vendas
        except Exception as e:
            st.error(f"Erro nas Vendas: {e}")

    if arq_odp:
        try:
            odp_vert = carregar_odp_vertical(arq_odp)
            st.success(f"✅ ODP Vertical: **{len(odp_vert):,}** recebimentos")
            with st.expander("Prévia"):
                st.dataframe(odp_vert.head(8), use_container_width=True, hide_index=True)
            st.session_state['odp_vert'] = odp_vert
        except Exception as e:
            st.error(f"Erro ODP Vertical: {e}")

    if arq_params or arq_curvas:
        try:
            params = carregar_parametros(arq_params, arq_curvas)
            abas_ok = list(params.keys())
            st.success(f"✅ Parâmetros carregados: {', '.join(abas_ok)}")
            st.session_state['params'] = params
        except Exception as e:
            st.error(f"Erro nos Parâmetros: {e}")

    # Persiste configurações globais
    st.session_state['data_hoje']  = data_hoje
    st.session_state['sel_cd']     = sel_cd_global

# Acesso seguro ao session state
cal       = st.session_state.get('cal',       None)
vendas    = st.session_state.get('vendas',    None)
odp_vert  = st.session_state.get('odp_vert', None)
params    = st.session_state.get('params',    {})
data_ref  = st.session_state.get('data_hoje', date.today())
ajustes   = st.session_state.get('ajustes',  {})
sel_cd    = st.session_state.get('sel_cd',   'Extrema e Jandira')
cal_ok    = cal is not None and len(cal) > 0

def filtrar_cd(df, col='cd'):
    if sel_cd == 'Extrema e Jandira' or col not in df.columns:
        return df
    return df[df[col].str.lower().str.contains(sel_cd.lower(), na=False)]

# ═══════════════════════════════════════════════════════════════════════════════
# ABA ENTRADAS (ODP VERTICAL)
# ═══════════════════════════════════════════════════════════════════════════════
with aba_entradas:
    if odp_vert is None:
        st.info("Carregue o ODP Vertical na aba Dados.")
    else:
        st.subheader("Recebimentos no CD")

        # Filtros
        ef1, ef2 = st.columns(2)
        with ef1:
            conv_ent = st.radio("Unidade", ["Peças brutas","Peças convertidas"],
                                horizontal=True, key='ent_conv')
        with ef2:
            cats_ent = sorted(odp_vert['categoria'].dropna().unique().tolist()) if 'categoria' in odp_vert.columns else []
            sel_cat_ent = st.multiselect("Categoria", cats_ent, default=cats_ent, key='ent_cat')

        df_ent = odp_vert.copy()
        if 'categoria' in df_ent.columns and sel_cat_ent:
            df_ent = df_ent[df_ent['categoria'].isin(sel_cat_ent)]

        col_v = 'pecas_convertidas' if conv_ent == 'Peças convertidas' else 'pecas'
        col_v = col_v if col_v in df_ent.columns else 'pecas'

        m1, m2, m3 = st.columns(3)
        m1.metric("Total peças brutas",     f"{df_ent['pecas'].sum():,.0f}".replace(',','.') if 'pecas' in df_ent.columns else "—")
        m2.metric("Total peças convertidas", f"{df_ent['pecas_convertidas'].sum():,.0f}".replace(',','.') if 'pecas_convertidas' in df_ent.columns else "—")
        m3.metric("Recebimentos",           str(len(df_ent)))

        # Série temporal
        if 'data_entrada' in df_ent.columns:
            st.subheader("Recebimentos por dia")
            serie_ent = df_ent.groupby(df_ent['data_entrada'].dt.date)[col_v].sum().reset_index()
            serie_ent.columns = ['data', col_v]
            serie_ent['data'] = pd.to_datetime(serie_ent['data'])
            st.bar_chart(serie_ent.set_index('data'), use_container_width=True)

        st.subheader("Detalhamento por campanha")
        agg_cols = ['id_campanha','campanha','categoria']
        agg_cols = [c for c in agg_cols if c in df_ent.columns]
        if agg_cols and col_v in df_ent.columns:
            por_camp = (df_ent.groupby(agg_cols)[col_v]
                        .sum().reset_index()
                        .sort_values(col_v, ascending=False))
            por_camp[col_v] = por_camp[col_v].round(0).astype(int)
            st.dataframe(por_camp, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA SAÍDAS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_saidas:
    if not cal_ok:
        st.info("Carregue o Calendário na aba Dados.")
    else:
        sf1, sf2, sf3, sf4 = st.columns(4)
        with sf1:
            conv_sai = st.radio("Unidade", ["Não convertido","Convertido"],
                                horizontal=True, key='sai_conv')
        with sf2:
            gers  = sorted(cal['gerencia'].dropna().unique().tolist()) if 'gerencia' in cal.columns else []
            sel_ger_s = st.multiselect("Gerência", gers, default=gers, key='sai_ger')
        with sf3:
            cats  = sorted(cal['categoria'].dropna().unique().tolist())
            sel_cat_s = st.multiselect("Categoria", cats, default=cats, key='sai_cat')
        with sf4:
            tipos_vis = st.multiselect("Tipo", ["realizado","reforecast","previsao"],
                                       default=["realizado","reforecast","previsao"],
                                       key='sai_tipo')

        cal_f = filtrar_cd(cal)
        if gers and sel_ger_s and 'gerencia' in cal_f.columns:
            cal_f = cal_f[cal_f['gerencia'].isin(sel_ger_s)]
        if cats and sel_cat_s:
            cal_f = cal_f[cal_f['categoria'].isin(sel_cat_s)]
        # filtra horizonte relevante
        horizonte_sai_ini = data_ref - timedelta(days=60)
        horizonte_sai_fim = data_ref + timedelta(days=120)
        cal_f = cal_f[
            (cal_f['data_fim']    >= pd.Timestamp(horizonte_sai_ini)) &
            (cal_f['data_inicio'] <= pd.Timestamp(horizonte_sai_fim))
        ]

        with st.spinner("Calculando saídas..."):
            df_sai = calcular_saidas(cal_f, vendas, data_ref, params, ajustes)

        if df_sai.empty:
            st.warning("Nenhum dado para os filtros selecionados.")
        else:
            col_p = 'pecas_conv' if conv_sai == 'Convertido' else 'pecas_bruto'
            df_vis = df_sai[df_sai['tipo'].isin(tipos_vis)]

            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Total peças", f"{df_vis[col_p].sum():,.0f}".replace(',','.'))
            m2.metric("Realizado",   f"{df_vis[df_vis['tipo']=='realizado'][col_p].sum():,.0f}".replace(',','.'))
            m3.metric("A realizar",  f"{df_vis[df_vis['tipo'].isin(['reforecast','previsao'])][col_p].sum():,.0f}".replace(',','.'))
            m4.metric("Campanhas",   str(df_vis['id_campanha'].nunique()))

            st.subheader("Volume de picking por dia")
            pivot = (df_vis.groupby(['data','tipo'])[col_p].sum().reset_index()
                     .pivot_table(index='data', columns='tipo', values=col_p, aggfunc='sum')
                     .fillna(0).sort_index())
            pivot.index = pd.to_datetime(pivot.index)
            pivot.columns.name = None
            pivot = pivot.rename(columns={'realizado':'Realizado','reforecast':'Reforecast','previsao':'Previsão'})
            st.bar_chart(pivot, use_container_width=True)

            st.subheader("Por gerência")
            if 'gerencia' in df_vis.columns:
                por_ger = (df_vis.groupby(['gerencia','tipo'])[col_p].sum()
                           .unstack(fill_value=0).reset_index())
                por_ger.columns.name = None
                st.dataframe(por_ger, use_container_width=True, hide_index=True)

            st.subheader("Por campanha")
            por_camp = (df_vis.groupby(['id_campanha','campanha','categoria','cd'])[col_p]
                        .sum().reset_index().sort_values(col_p, ascending=False))
            por_camp[col_p] = por_camp[col_p].round(0).astype(int)
            st.dataframe(por_camp, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA BLOCADO
# ═══════════════════════════════════════════════════════════════════════════════
with aba_blocado:
    if not cal_ok:
        st.info("Carregue o Calendário na aba Dados.")
    else:
        bf1, bf2 = st.columns(2)
        with bf1:
            conv_bl = st.radio("Unidade", ["Peças","Peças convertidas","Pallets"],
                               horizontal=True, key='bl_conv')
        with bf2:
            gers_b = sorted(cal['gerencia'].dropna().unique().tolist()) if 'gerencia' in cal.columns else []
            sel_ger_b = st.multiselect("Gerência", gers_b, default=gers_b, key='bl_ger')

        cal_fb = filtrar_cd(cal)
        if gers_b and sel_ger_b and 'gerencia' in cal_fb.columns:
            cal_fb = cal_fb[cal_fb['gerencia'].isin(sel_ger_b)]
        # filtra só campanhas com blocado no horizonte relevante (não mostra histórico de anos anteriores)
        horizonte_bl_ini = data_ref - timedelta(days=30)
        horizonte_bl_fim = data_ref + timedelta(days=120)
        cal_fb = cal_fb[
            (cal_fb['data_fim']    >= pd.Timestamp(horizonte_bl_ini)) &
            (cal_fb['data_inicio'] <= pd.Timestamp(horizonte_bl_fim))
        ]

        with st.spinner("Calculando blocado..."):
            df_bloc, serie_bloc = calcular_blocado(cal_fb, data_ref, params, ajustes)

        if df_bloc.empty:
            st.warning("Nenhum dado de blocado.")
        else:
            col_s = {'Peças':'pecas','Peças convertidas':'pecas_conv','Pallets':'pallets'}[conv_bl]

            m1,m2,m3 = st.columns(3)
            hoje_bl = df_bloc[df_bloc['status']=='blocado']
            m1.metric("Campanhas blocadas hoje", str(len(hoje_bl)))
            m2.metric(f"{conv_bl} blocadas hoje",
                      f"{hoje_bl[col_s.replace('pecas','estoque_pecas').replace('pecas_conv','estoque_conv')].sum():,.0f}".replace(',','.'))
            m3.metric("Aguardando entrada", str(len(df_bloc[df_bloc['status']=='aguardando'])))

            st.subheader(f"Saldo blocado por dia — {conv_bl}")
            if not serie_bloc.empty:
                sv = serie_bloc[['data',col_s]].set_index('data').rename(columns={col_s: conv_bl})
                sv.index = pd.to_datetime(sv.index)
                st.line_chart(sv, use_container_width=True)

            st.subheader("Campanhas no picking")
            exib = df_bloc[['id_campanha','campanha','categoria','cd','gerencia',
                             'data_inicio','data_fim','dn','do',
                             'estoque_pecas','estoque_conv','pallets','status']].copy()
            for c in ['data_inicio','data_fim','dn','do']:
                exib[c] = exib[c].astype(str)
            exib['estoque_pecas'] = exib['estoque_pecas'].astype(int)
            exib['estoque_conv']  = exib['estoque_conv'].astype(int)
            exib = exib.rename(columns={
                'id_campanha':'ID','campanha':'Campanha','categoria':'Categoria',
                'cd':'CD','gerencia':'Gerência','data_inicio':'Início','data_fim':'Fim',
                'dn':'DN (Descida)','do':'DO (Subida)',
                'estoque_pecas':'Peças','estoque_conv':'Peças conv.',
                'pallets':'Pallets','status':'Status'})
            st.dataframe(exib, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA REFORECAST
# ═══════════════════════════════════════════════════════════════════════════════
with aba_rfcst:
    if not cal_ok:
        st.info("Carregue o Calendário na aba Dados.")
    else:
        # filtra só campanhas ativas no horizonte M-2 a M+3 (evita processar histórico de 2024)
        horizonte_ini = data_ref - timedelta(days=60)
        horizonte_fim = data_ref + timedelta(days=120)
        cal_frf = filtrar_cd(cal)
        cal_frf = cal_frf[
            (cal_frf['data_fim']    >= pd.Timestamp(horizonte_ini)) &
            (cal_frf['data_inicio'] <= pd.Timestamp(horizonte_fim))
        ]
        df_sai_rf = calcular_saidas(cal_frf, vendas, data_ref, params, ajustes)

        if df_sai_rf.empty:
            st.warning("Sem dados.")
        else:
            st.subheader("Previsão original vs Realizado vs A realizar")
            prev_orig = cal_frf[['id_campanha','previsao_pecas']].copy()
            resumo_rf = (df_sai_rf.groupby(['id_campanha','campanha','categoria'])
                         .agg(
                             realizado  =('pecas_bruto', lambda x: x[df_sai_rf.loc[x.index,'tipo']=='realizado'].sum()),
                             a_realizar =('pecas_bruto', lambda x: x[df_sai_rf.loc[x.index,'tipo'].isin(['reforecast','previsao'])].sum()),
                         ).reset_index())
            resumo_rf = resumo_rf.merge(prev_orig, on='id_campanha', how='left')
            resumo_rf['total_reforecast'] = resumo_rf['realizado'] + resumo_rf['a_realizar']
            resumo_rf['desvio_pct'] = (
                (resumo_rf['total_reforecast'] - resumo_rf['previsao_pecas'])
                / resumo_rf['previsao_pecas'].replace(0, np.nan) * 100
            ).round(1)
            for c in ['previsao_pecas','realizado','a_realizar','total_reforecast']:
                resumo_rf[c] = resumo_rf[c].round(0).astype(int)
            resumo_rf = resumo_rf.rename(columns={
                'id_campanha':'ID','campanha':'Campanha','categoria':'Categoria',
                'previsao_pecas':'Previsão original','realizado':'Realizado',
                'a_realizar':'A realizar','total_reforecast':'Reforecast total',
                'desvio_pct':'Desvio % (rfcst vs prev)'})
            st.dataframe(resumo_rf, use_container_width=True, hide_index=True)

            st.subheader("Evolução acumulada")
            acum = (df_sai_rf.sort_values('data')
                    .assign(data=lambda x: pd.to_datetime(x['data']))
                    .groupby(['data','tipo'])['pecas_bruto'].sum().reset_index()
                    .pivot_table(index='data',columns='tipo',values='pecas_bruto',aggfunc='sum')
                    .fillna(0).cumsum())
            acum.columns.name = None
            acum = acum.rename(columns={'realizado':'Realizado','reforecast':'Reforecast','previsao':'Previsão'})
            st.line_chart(acum, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA SIMULAÇÕES
# ═══════════════════════════════════════════════════════════════════════════════
with aba_sim:
    if not cal_ok:
        st.info("Carregue o Calendário na aba Dados.")
    else:
        st.caption("Ajustes são efêmeros — não afetam as bases originais. Use 'Resetar' para voltar ao estado original.")

        if st.button("↺ Resetar todas as simulações"):
            st.session_state['ajustes'] = {}
            st.rerun()

        camp_ids   = cal['id_campanha'].tolist()
        camp_nomes = cal.set_index('id_campanha')['campanha'].to_dict()
        # usa '||' como separador para evitar problemas com travessão unicode
        opcoes     = [f"{cid}||{camp_nomes.get(cid,'')}" for cid in camp_ids]
        opcoes_vis = [f"{cid} — {camp_nomes.get(cid,'')}" for cid in camp_ids]

        sel_idx = st.selectbox("Campanha para ajustar", range(len(opcoes_vis)),
                               format_func=lambda i: opcoes_vis[i])
        cid_sel = camp_ids[sel_idx]
        camp_row = cal[cal['id_campanha'] == cid_sel].iloc[0]
        aj_atual = st.session_state['ajustes'].get(cid_sel, {})

        st.markdown(f"**{camp_row['campanha']}** — {camp_row['categoria']} | CD: {camp_row.get('cd','')}")

        col1, col2 = st.columns(2)
        with col1:
            new_ini = st.date_input("Data início",
                value=aj_atual.get('data_inicio', camp_row['data_inicio'].date()),
                key=f'ini_{cid_sel}')
            new_fim = st.date_input("Data fim",
                value=aj_atual.get('data_fim', camp_row['data_fim'].date()),
                key=f'fim_{cid_sel}')
            new_est = st.number_input("Estoque total (peças)",
                value=float(aj_atual.get('estoque_total', camp_row['estoque_total'])),
                min_value=0.0, step=100.0, key=f'est_{cid_sel}')
        with col2:
            feriados_sim = get_feriados_set(params, camp_row.get('cd','Extrema'))
            dn_calc, do_calc = calcular_datas_blocado(new_ini, new_fim, feriados_sim)
            new_dn = st.date_input("Data descida (DN)",
                value=aj_atual.get('dn', dn_calc), key=f'dn_{cid_sel}')
            new_do = st.date_input("Data subida (DO)",
                value=aj_atual.get('do', do_calc), key=f'do_{cid_sel}')
            new_prev = st.number_input("Previsão de vendas (peças)",
                value=float(aj_atual.get('previsao_pecas', camp_row['previsao_pecas'])),
                min_value=0.0, step=100.0, key=f'prev_{cid_sel}')

        if st.button("✓ Aplicar ajuste"):
            st.session_state['ajustes'][cid_sel] = {
                'data_inicio':   new_ini,
                'data_fim':      new_fim,
                'estoque_total': new_est,
                'dn':            new_dn,
                'do':            new_do,
                'previsao_pecas':new_prev,
            }
            st.success(f"Ajuste aplicado para {camp_row['campanha']}. Verifique as abas Saídas e Blocado.")

        if st.session_state['ajustes']:
            st.subheader("Ajustes ativos")
            rows = []
            for cid, aj in st.session_state['ajustes'].items():
                rows.append({
                    'ID': cid, 'Campanha': camp_nomes.get(cid,''),
                    'Início': str(aj.get('data_inicio','')),
                    'Fim':    str(aj.get('data_fim','')),
                    'Estoque': aj.get('estoque_total',''),
                    'DN':     str(aj.get('dn','')),
                    'DO':     str(aj.get('do',''))
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ABA PARÂMETROS
# ═══════════════════════════════════════════════════════════════════════════════
with aba_params:
    senha = st.text_input("Senha de acesso", type="password", key="senha_params")
    if senha == "privalia2024":
        st.success("Acesso liberado.")

        tab_p1, tab_p2, tab_p3, tab_p4 = st.tabs(
            ["Pcs/Palete","Fatores IN/OUT","Feriados","Curvas de Venda"])

        with tab_p1:
            st.caption("Peças por palete por categoria — edite e os cálculos de blocado atualizam automaticamente.")
            if params and 'pcs_palete' in params:
                df_edit = st.data_editor(params['pcs_palete'][['categoria','pcs_palete']],
                                         use_container_width=True, hide_index=True, num_rows="dynamic")
                params['pcs_palete'] = df_edit
            else:
                df_def = pd.DataFrame(list(PCS_DEFAULT.items()), columns=['categoria','pcs_palete'])
                df_edit = st.data_editor(df_def, use_container_width=True,
                                         hide_index=True, num_rows="dynamic")
                if 'pcs_palete' not in params: params['pcs_palete'] = df_edit

        with tab_p2:
            st.caption("Fatores de conversão Warehouse IN e OUT por grupo de categoria.")
            if params and 'fatores' in params:
                df_edit2 = st.data_editor(params['fatores'][['grupo','fator_in','fator_out']],
                                           use_container_width=True, hide_index=True, num_rows="dynamic")
                params['fatores'] = df_edit2
            else:
                df_def2 = pd.DataFrame(
                    [(g, v[0], v[1]) for g,v in FATORES_DEFAULT.items()],
                    columns=['grupo','fator_in','fator_out'])
                df_edit2 = st.data_editor(df_def2, use_container_width=True,
                                           hide_index=True, num_rows="dynamic")
                if 'fatores' not in params: params['fatores'] = df_edit2

        with tab_p3:
            st.caption("Feriados usados no cálculo de DN/DO. Local: Nacional, Extrema ou Jandira.")
            if params and 'feriados' in params:
                df_edit3 = st.data_editor(params['feriados'],
                                           use_container_width=True, hide_index=True, num_rows="dynamic")
                params['feriados'] = df_edit3

        with tab_p4:
            st.caption("Curvas de venda: distribuição percentual por dia relativo.")
            if params and 'curvas' in params:
                cols_show = ['categoria','webdays'] + [c for c in params['curvas'].columns if c.startswith('d')]
                st.dataframe(params['curvas'][cols_show], use_container_width=True, hide_index=True)

        st.session_state['params'] = params

    elif senha:
        st.error("Senha incorreta.")

# ═══════════════════════════════════════════════════════════════════════════════
# ABA EXPORTAR
# ═══════════════════════════════════════════════════════════════════════════════
with aba_export:
    st.subheader("Exportar resultados")
    st.caption("Exporta o estado atual — incluindo ajustes de simulação. Nada é salvo no sistema.")

    if not cal_ok:
        st.info("Carregue os dados primeiro.")
    else:
        cal_fexp = filtrar_cd(cal)
        df_sai_exp = calcular_saidas(cal_fexp, vendas, data_ref, params, ajustes)
        df_bloc_exp, serie_bloc_exp = calcular_blocado(cal_fexp, data_ref, params, ajustes)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            if not df_sai_exp.empty:
                pivot_exp = (df_sai_exp.groupby(['data','tipo'])['pecas_bruto']
                             .sum().reset_index()
                             .pivot_table(index='data',columns='tipo',values='pecas_bruto',aggfunc='sum')
                             .fillna(0).sort_index())
                pivot_exp.columns.name = None
                pivot_exp.to_excel(writer, sheet_name='Saídas — dia a dia')
                df_sai_exp.to_excel(writer, sheet_name='Saídas — detalhe', index=False)
            if not df_bloc_exp.empty:
                df_bloc_exp.to_excel(writer, sheet_name='Blocado — campanhas', index=False)
            if not serie_bloc_exp.empty:
                serie_bloc_exp.to_excel(writer, sheet_name='Blocado — série', index=False)
            cal_fexp.to_excel(writer, sheet_name='Calendário ODP', index=False)
            if ajustes:
                df_aj = pd.DataFrame([
                    {'ID': k, **{str(kk): str(vv) for kk,vv in v.items()}}
                    for k, v in ajustes.items()
                ])
                df_aj.to_excel(writer, sheet_name='Ajustes aplicados', index=False)

        st.download_button(
            "⬇ Baixar Excel",
            data=buf.getvalue(),
            file_name=f"simulador_odp_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
