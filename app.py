import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date, timedelta
import tempfile, os, sys
sys.path.insert(0, '/home/claude')
from engine import calcular, serie_blocado, serie_saidas, carregar_curvas
from openpyxl import load_workbook

st.set_page_config(page_title="ODP · Privalia", page_icon="📦",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.block-container{padding:1.2rem 2rem 2rem;}
[data-baseweb="tab-list"]{gap:0;border-bottom:1px solid #e5e7eb;}
[data-baseweb="tab"]{font-size:13px;font-weight:500;padding:10px 18px;color:#6b7280;}
[aria-selected="true"]{color:#111827!important;border-bottom:2px solid #111827!important;}
[data-testid="metric-container"]{background:#f9fafb;border-radius:10px;padding:14px 18px;border:1px solid #f3f4f6;}
[data-testid="stMetricValue"]{font-size:24px!important;font-weight:600;}
[data-testid="stMetricLabel"]{font-size:11px!important;color:#6b7280;}
[data-testid="stExpander"]{border:1px solid #f3f4f6!important;border-radius:10px;}
.badge-on{background:#dcfce7;color:#166534;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;}
.badge-x{background:#f3f4f6;color:#6b7280;padding:2px 8px;border-radius:4px;font-size:11px;}
.badge-fut{background:#eff6ff;color:#1d4ed8;padding:2px 8px;border-radius:4px;font-size:11px;}
.badge-aj{background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:4px;font-size:11px;}
.sec{font-size:11px;font-weight:600;color:#9ca3af;text-transform:uppercase;letter-spacing:.06em;margin:18px 0 8px;}
div[data-testid="stForm"]{border:1px solid #f3f4f6;border-radius:10px;padding:16px;}
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fn(v, dec=0):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v:,.{dec}f}".replace(",","X").replace(".",",").replace("X",".")

def fp(v, dec=1):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v*100:.{dec}f}%"

CORES = {'realizado':'#1e3a5f','reforecast':'#3b82f6','previsto':'#bfdbfe'}

# ── Session state ─────────────────────────────────────────────────────────────
for k,v in [('df',None),('curvas',{}),('data_ref',date.today()),
             ('ajustes',{}),('arquivo_path',None)]:
    if k not in st.session_state: st.session_state[k] = v

def reload():
    if not st.session_state.arquivo_path: return
    with st.spinner("Recalculando..."):
        df, curvas, dr = calcular(st.session_state.arquivo_path,
                                  st.session_state.data_ref,
                                  st.session_state.ajustes)
    st.session_state.df     = df
    st.session_state.curvas = curvas

@st.cache_data(show_spinner=False)
def processar(arq_bytes, nome):
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as t:
        t.write(arq_bytes); path = t.name
    df, curvas, dr = calcular(path)
    return df, curvas, dr, path

# ── Header ────────────────────────────────────────────────────────────────────
h1, h2, h3 = st.columns([3, 2, 2])
with h1: st.markdown("### ODP &nbsp;·&nbsp; Privalia")
with h2:
    if st.session_state.df is not None:
        dr = st.session_state.data_ref
        st.markdown(f"<div style='padding-top:6px;font-size:13px;color:#6b7280'>ref. {dr.strftime('%d/%m/%Y')}</div>",
                    unsafe_allow_html=True)
with h3:
    up = st.file_uploader("", type=['xlsm','xlsx'], label_visibility='collapsed')
    if up:
        with st.spinner("Lendo simulador..."):
            df_n, curvas_n, dr_n, path_n = processar(up.read(), up.name)
        st.session_state.df     = df_n
        st.session_state.curvas = curvas_n
        st.session_state.data_ref = dr_n
        st.session_state.arquivo_path = path_n
        st.session_state.ajustes = {}
        st.rerun()

st.markdown("<hr style='margin:0 0 12px;border:none;border-top:1px solid #f3f4f6'>",
            unsafe_allow_html=True)

df = st.session_state.df
if df is None or df.empty:
    st.info("Faça upload do arquivo do Calendário Salesforce (.xlsm) para começar.")
    st.stop()

data_ref  = st.session_state.data_ref
ontem     = data_ref - timedelta(days=1)
df_on     = df[df.status_op == 'ON'].copy()
df_ontem  = df[(df.status_op == 'x') & (df['ed'].dt.date == ontem)].copy()
df_fut    = df[df.status_op == '-'].copy()
ajustes   = st.session_state.ajustes

# ── ABAS ──────────────────────────────────────────────────────────────────────
t_dash, t_sai, t_bloc, t_rfcst, t_sim, t_dados = st.tabs([
    "Dashboard", "Saídas", "Blocado", "Reforecast", "Simulações", "Dados"
])

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with t_dash:
    alertas = df_on[df_on['desvio_abs'].abs() > 0.15]
    pal_hoje = df_on['pallets'].sum()

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Campanhas ON", len(df_on))
    m2.metric("Encerradas ontem", len(df_ontem))
    m3.metric("Pallets blocados hoje", fn(pal_hoje))
    m4.metric("Alertas (desvio >15 p.p.)", len(alertas),
              delta_color="inverse" if len(alertas) else "off",
              delta="atenção necessária" if len(alertas) else "ok")

    incl_enc = st.checkbox("Incluir encerradas ontem", False, key='d_enc')
    df_view  = pd.concat([df_on, df_ontem]) if incl_enc else df_on

    st.markdown('<div class="sec">Campanhas</div>', unsafe_allow_html=True)
    for _, c in df_view.sort_values('ed').iterrows():
        aj_flag = c['id'] in ajustes
        status_b = ({'ON':'badge-on','x':'badge-x'}).get(c['status_op'],'badge-fut')
        alerta_s = ""
        if abs(c['desvio_abs']) > 0.15:
            s = "+" if c['desvio_abs'] > 0 else ""
            alerta_s = f' &nbsp;<span style="color:#dc2626;font-size:11px;font-weight:600">⚑ {s}{c["desvio_abs"]*100:.1f} p.p.</span>'
        aj_s = ' &nbsp;<span class="badge-aj">ajustada</span>' if aj_flag else ''

        with st.expander(c['campanha']):
            st.markdown(
                f'<span class="{status_b}">{c["status_op"]}</span> &nbsp;'
                f'<b>{c["campanha"]}</b> &nbsp;'
                f'<span style="color:#9ca3af;font-size:12px">'
                f'{c["cat"]} · {c["cd"]} · '
                f'{c["sd"].strftime("%d/%m")}→{c["ed"].strftime("%d/%m")} · '
                f'dia {c["dia_atual"]} de {c["wd"]}'
                f'</span>{alerta_s}{aj_s}',
                unsafe_allow_html=True)

            a1,a2,a3,a4,a5,a6 = st.columns(6)
            a1.metric("Estoque",    fn(c['estoque']))
            a2.metric("Pallets",    fn(c['pallets']))
            a3.metric("Forecast",   fn(c['forecast']))
            a4.metric("SO congelado", fp(c['so_cong']))
            a5.metric("Venda real", fn(c['venda_real']))
            a6.metric("SO real",    fp(c['so_real_acum']))

            b1,b2,b3,b4 = st.columns(4)
            b1.metric("Reforecast", fn(c['reforecast']))
            b2.metric("Novo SO",    fp(c['novo_so']))
            b3.metric("DN (descida)", c['dn'].strftime('%d/%m/%Y'))
            b4.metric("DO (subida)", c['do_'].strftime('%d/%m/%Y'))

            prog = float(c['dia_atual']/c['wd']) if c['wd'] > 0 else 0
            st.progress(prog, text=f"Progresso: dia {c['dia_atual']} / {c['wd']}")

# ══════════════════════════════════════════════════════════════════════════════
# SAÍDAS
# ══════════════════════════════════════════════════════════════════════════════
with t_sai:
    f1,f2,f3,f4 = st.columns([2,2,2,3])
    sel_cd_s  = f1.selectbox("CD", ['Todos']+sorted(df['cd'].dropna().unique().tolist()), key='s_cd')
    sel_cat_s = f2.selectbox("Categoria", ['Todas']+sorted(df['cat'].dropna().unique().tolist()), key='s_cat')
    uni_s     = f3.radio("Unidade", ["Pç convertidas","Pç brutas"], horizontal=True, key='s_uni')
    rng_s     = f4.date_input("Período",
                value=(data_ref-timedelta(days=7), data_ref+timedelta(days=30)),
                key='s_per')

    dfi = df.copy()
    if sel_cd_s  != 'Todos': dfi = dfi[dfi['cd']  == sel_cd_s]
    if sel_cat_s != 'Todas': dfi = dfi[dfi['cat'] == sel_cat_s]

    d0 = rng_s[0] if len(rng_s)==2 else data_ref-timedelta(days=7)
    d1 = rng_s[1] if len(rng_s)==2 else data_ref+timedelta(days=30)

    uni_str = 'pecas_conv' if uni_s.startswith('Pç conv') else 'pecas_brutas'
    with st.spinner("Calculando saídas..."):
        ss = serie_saidas(dfi, d0, d1, st.session_state.curvas, uni_str)

    if not ss.empty:
        pv = ss.groupby(['data','tipo'])['volume'].sum().unstack(fill_value=0).reset_index()
        for t in ['realizado','reforecast','previsto']:
            if t not in pv.columns: pv[t] = 0

        fig = go.Figure()
        for tipo, cor, lab in [('realizado',CORES['realizado'],'Realizado'),
                                ('reforecast',CORES['reforecast'],'Reforecast'),
                                ('previsto',CORES['previsto'],'Previsto')]:
            fig.add_bar(x=pv['data'], y=pv[tipo], name=lab, marker_color=cor)
        fig.update_layout(barmode='stack', height=300,
                          margin=dict(l=0,r=0,t=10,b=0),
                          legend=dict(orientation='h',y=1.08),
                          xaxis=dict(showgrid=False),
                          yaxis=dict(gridcolor='#f3f4f6', title=uni_s),
                          plot_bgcolor='white', paper_bgcolor='white',
                          font=dict(family='DM Sans',size=12))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sec">Detalhe por campanha</div>', unsafe_allow_html=True)
    for _, c in dfi.sort_values('sd').iterrows():
        with st.expander(f"{c['campanha']} · {c['cat']} · {c['sd'].strftime('%d/%m')}→{c['ed'].strftime('%d/%m')}"):
            x1,x2,x3,x4 = st.columns(4)
            x1.metric("Reforecast", fn(c['reforecast']))
            x2.metric("Fator OUT",  f"{c['fat_out']:.3f}×")
            x3.metric("Curva",      f"{c['cat']} {c['wd']}wd")
            x4.metric("Dias rest.", fn(c['dias_rest']))
            if not ss.empty:
                sc = ss[ss['id']==c['id']]
                if not sc.empty:
                    sp = sc.groupby(['data','tipo'])['volume'].sum().unstack(fill_value=0).reset_index()
                    for t in ['realizado','reforecast','previsto']:
                        if t not in sp.columns: sp[t]=0
                    fc = go.Figure()
                    for tipo, cor in [('realizado',CORES['realizado']),
                                      ('reforecast',CORES['reforecast']),
                                      ('previsto',CORES['previsto'])]:
                        fc.add_bar(x=sp['data'], y=sp[tipo], name=tipo, marker_color=cor)
                    fc.update_layout(barmode='stack', height=160, showlegend=False,
                                     margin=dict(l=0,r=0,t=4,b=0),
                                     xaxis=dict(showgrid=False),
                                     yaxis=dict(gridcolor='#f3f4f6'),
                                     plot_bgcolor='white', paper_bgcolor='white')
                    st.plotly_chart(fc, use_container_width=True)

    if not ss.empty:
        csv = ss.groupby(['data','campanha','cat','cd','tipo'])['volume'].sum().reset_index()
        st.download_button("⬇ Exportar saídas (.csv)",
                           csv.to_csv(index=False,sep=';').encode('utf-8-sig'),
                           "saidas_odp.csv","text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# BLOCADO
# ══════════════════════════════════════════════════════════════════════════════
with t_bloc:
    b1,b2,b3 = st.columns([2,3,3])
    sel_cd_b = b1.selectbox("CD", ['Todos']+sorted(df['cd'].dropna().unique().tolist()), key='b_cd')
    uni_b    = b2.radio("Unidade", ["Pallets","Pç brutas","Pç convertidas"], horizontal=True, key='b_uni')
    rng_b    = b3.date_input("Período",
               value=(data_ref-timedelta(days=7), data_ref+timedelta(days=45)), key='b_per')

    dfb = df.copy()
    if sel_cd_b != 'Todos': dfb = dfb[dfb['cd']==sel_cd_b]
    db0 = rng_b[0] if len(rng_b)==2 else data_ref-timedelta(days=7)
    db1 = rng_b[1] if len(rng_b)==2 else data_ref+timedelta(days=45)

    uni_bstr = {'Pallets':'pallets','Pç brutas':'pecas_brutas','Pç convertidas':'pecas_conv'}[uni_b]
    with st.spinner("Calculando blocado..."):
        sb = serie_blocado(dfb, db0, db1, uni_bstr)

    if not sb.empty:
        fig_b = go.Figure()
        fig_b.add_scatter(x=sb['data'], y=sb['volume'],
                          fill='tozeroy', fillcolor='rgba(59,130,246,0.12)',
                          line=dict(color='#3b82f6',width=2), name=uni_b,
                          hovertemplate='%{x|%d/%m}<br>%{y:,.0f}<extra></extra>')
        fig_b.update_layout(height=300, margin=dict(l=0,r=0,t=10,b=0),
                             xaxis=dict(showgrid=False),
                             yaxis=dict(gridcolor='#f3f4f6', title=uni_b),
                             plot_bgcolor='white', paper_bgcolor='white',
                             font=dict(family='DM Sans',size=12))
        st.plotly_chart(fig_b, use_container_width=True)

    pico = sb['volume'].max() if not sb.empty else 0
    desc_sem = dfb[(dfb['dn'].dt.date >= data_ref) & (dfb['dn'].dt.date <= data_ref+timedelta(7))]
    sub_sem  = dfb[(dfb['do_'].dt.date >= data_ref) & (dfb['do_'].dt.date <= data_ref+timedelta(7))]
    c1,c2,c3 = st.columns(3)
    c1.metric(f"Pico blocado ({uni_b})", fn(pico))
    c2.metric("Descidas esta semana", len(desc_sem), delta="campanhas entram no picking")
    c3.metric("Subidas esta semana",  len(sub_sem),  delta="campanhas saem do picking")

    st.markdown('<div class="sec">Detalhe por campanha</div>', unsafe_allow_html=True)
    for _, c in dfb.sort_values('dn').iterrows():
        ic = {'ON':'🟢','x':'⚪','-':'🔵'}.get(c['status_op'],'')
        with st.expander(f"{ic} {c['campanha']} · DN {c['dn'].strftime('%d/%m')} → DO {c['do_'].strftime('%d/%m')}"):
            d1,d2,d3,d4,d5 = st.columns(5)
            d1.metric("Pallets",   fn(c['pallets']))
            d2.metric("Pç brutas", fn(c['estoque']))
            d3.metric("Pç/palete", fn(c['pcs_pal']))
            d4.metric("Fator IN",  f"{c['fat_in']:.3f}×")
            d5.metric("CD",        c['cd'])

    st.download_button("⬇ Exportar blocado (.csv)",
        dfb[['id','campanha','cat','setor','cd','status_op','sd','ed',
             'dn','do_','estoque','pallets','pcs_pal','fat_in','fat_out']]
        .to_csv(index=False,sep=';').encode('utf-8-sig'),
        "blocado_odp.csv","text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# REFORECAST
# ══════════════════════════════════════════════════════════════════════════════
with t_rfcst:
    r1,r2,r3 = st.columns([2,2,2])
    show_rf  = r1.radio("Campanhas", ["Só ON","Todas"], horizontal=True, key='r_sh')
    fil_rf   = r2.radio("Filtro", ["Todas","Só alertas (>15 p.p.)"], horizontal=True, key='r_fi')
    sel_cd_r = r3.selectbox("CD", ['Todos']+sorted(df['cd'].dropna().unique().tolist()), key='r_cd')

    dfr = df.copy()
    if show_rf == "Só ON":                   dfr = dfr[dfr['status_op']=='ON']
    if fil_rf  == "Só alertas (>15 p.p.)":  dfr = dfr[dfr['desvio_abs'].abs()>0.15]
    if sel_cd_r != 'Todos':                  dfr = dfr[dfr['cd']==sel_cd_r]
    dfr = dfr.sort_values('desvio_abs', key=abs, ascending=False)

    for _, c in dfr.iterrows():
        cor = '#16a34a' if abs(c['desvio_abs'])<0.05 else ('#d97706' if abs(c['desvio_abs'])<0.15 else '#dc2626')
        s   = "+" if c['desvio_abs']>=0 else ""
        ic  = {'ON':'🟢','x':'⚪','-':'🔵'}.get(c['status_op'],'')

        with st.expander(f"{ic} {c['campanha']} · dia {c['dia_atual']}/{c['wd']} · {s}{c['desvio_abs']*100:.1f} p.p."):
            pp = float(min(1, c['so_prev_acum']))
            pr = float(min(1, c['so_real_acum']))
            st.markdown(f"""
            <div style='margin-bottom:12px'>
              <div style='display:flex;justify-content:space-between;font-size:12px;color:#6b7280;margin-bottom:3px'>
                <span>SO previsto acum.</span><span>{pp*100:.1f}%</span></div>
              <div style='background:#f3f4f6;border-radius:4px;height:7px'>
                <div style='width:{pp*100:.1f}%;background:#93c5fd;border-radius:4px;height:100%'></div></div>
              <div style='display:flex;justify-content:space-between;font-size:12px;color:#6b7280;margin:7px 0 3px'>
                <span>SO real acum.</span>
                <span style='color:{cor};font-weight:600'>{pr*100:.1f}% ({s}{c["desvio_abs"]*100:.1f} p.p.)</span></div>
              <div style='background:#f3f4f6;border-radius:4px;height:7px'>
                <div style='width:{pr*100:.1f}%;background:{cor};border-radius:4px;height:100%'></div></div>
            </div>""", unsafe_allow_html=True)

            e1,e2,e3,e4,e5,e6 = st.columns(6)
            e1.metric("Venda real",   fn(c['venda_real']))
            e2.metric("Reforecast",   fn(c['reforecast']))
            e3.metric("Var vs FC",    fp(c['mudanca_fc']))
            e4.metric("Novo SO",      fp(c['novo_so']))
            e5.metric("SO congelado", fp(c['so_cong']))
            e6.metric("Fator ajuste", f"{c['fator_ajuste']:.3f}×")

            with st.expander("Racional QB → QR"):
                racional = {
                    "QB · Dias de venda":             c['dias_decor'],
                    "QC · Dias restantes":            c['dias_rest'],
                    "QD/QE · Proj. acum. (curva)":   fp(c['proj_acum']),
                    "QF · Venda real":                fn(c['venda_real']),
                    "QG · SO real acum.":             fp(c['so_real_acum']),
                    "QH · Desvio absoluto":           f"{c['desvio_abs']*100:+.2f} p.p.",
                    "QI · Desvio relativo":           fp(c['desvio_rel']),
                    "QJ · Desvio médio":              fp((c['desvio_abs']+c['desvio_rel'])/2),
                    "QK · Fator ajuste (1+QJ)":       f"{c['fator_ajuste']:.3f}×",
                    "QL · % restante curva":          fp(1-c['proj_acum']),
                    "QO · Reforecast final":          fn(c['reforecast']),
                    "QP · Var vs Forecast congelado": fp(c['mudanca_fc']),
                    "QQ · Novo sell-out (estoque)":   fp(c['novo_so']),
                    "QR · Var vs SO original":        f"{c['desvio_so']*100:+.2f} p.p.",
                }
                st.table(pd.DataFrame.from_dict(racional, orient='index', columns=['Valor']))

    cols_exp = ['id','campanha','cat','cd','status_op','sd','ed','wd',
                'estoque','forecast','so_cong','venda_real','so_real_acum',
                'desvio_abs','fator_ajuste','reforecast','novo_so','mudanca_fc']
    st.download_button("⬇ Exportar reforecast (.csv)",
        dfr[cols_exp].to_csv(index=False,sep=';').encode('utf-8-sig'),
        "reforecast_odp.csv","text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# SIMULAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
with t_sim:
    st.markdown('<div class="sec">Ajustes manuais por campanha</div>', unsafe_allow_html=True)
    st.caption("Altere datas, estoque ou forecast de uma campanha e veja o impacto imediato em todas as abas.")

    # Seleção da campanha
    opcoes = df[df['status_op'].isin(['ON','-'])].sort_values('sd')
    opcoes_nomes = [f"{r['campanha']} ({r['sd'].strftime('%d/%m')}→{r['ed'].strftime('%d/%m')})"
                    for _, r in opcoes.iterrows()]
    opcoes_ids   = list(opcoes['id'])

    if not opcoes_nomes:
        st.info("Nenhuma campanha ON ou futura disponível para simulação.")
    else:
        sel_nome = st.selectbox("Campanha", opcoes_nomes, key='sim_camp')
        sel_id   = opcoes_ids[opcoes_nomes.index(sel_nome)]
        camp_orig = df[df['id']==sel_id].iloc[0]
        aj_exist  = ajustes.get(sel_id, {})

        with st.form("form_sim"):
            st.markdown(f"**{camp_orig['campanha']}** · {camp_orig['cat']} · {camp_orig['cd']}")
            s1,s2,s3 = st.columns(3)
            nova_sd = s1.date_input("Data início (SD)",
                         value=aj_exist.get('sd', camp_orig['sd'].date()), key='sim_sd')
            nova_ed = s2.date_input("Data fim (ED)",
                         value=aj_exist.get('ed', camp_orig['ed'].date()), key='sim_ed')
            nova_wd = s3.number_input("Webdays", min_value=1, max_value=30,
                         value=int(aj_exist.get('wd', camp_orig['wd'])), key='sim_wd')
            s4,s5 = st.columns(2)
            novo_est = s4.number_input("Estoque (peças)",
                           value=float(aj_exist.get('estoque', camp_orig['estoque'])),
                           min_value=0.0, step=100.0, key='sim_est')
            novo_fc  = s5.number_input("Forecast (peças)",
                           value=float(aj_exist.get('forecast', camp_orig['forecast'])),
                           min_value=0.0, step=50.0, key='sim_fc')

            col_apply, col_reset = st.columns([2,1])
            aplicar = col_apply.form_submit_button("✓ Aplicar ajuste", use_container_width=True)
            remover = col_reset.form_submit_button("↺ Remover ajuste", use_container_width=True)

        if aplicar:
            st.session_state.ajustes[sel_id] = {
                'sd': nova_sd, 'ed': nova_ed, 'wd': nova_wd,
                'estoque': novo_est, 'forecast': novo_fc,
            }
            reload()
            st.success(f"Ajuste aplicado para {camp_orig['campanha']}. Todas as abas foram atualizadas.")
            st.rerun()

        if remover and sel_id in st.session_state.ajustes:
            del st.session_state.ajustes[sel_id]
            reload()
            st.success("Ajuste removido.")
            st.rerun()

        # Preview do impacto
        if aj_exist:
            camp_aj = df[df['id']==sel_id].iloc[0]
            st.markdown('<div class="sec">Impacto do ajuste</div>', unsafe_allow_html=True)
            i1,i2,i3,i4 = st.columns(4)
            i1.metric("Reforecast",
                       fn(camp_aj['reforecast']),
                       delta=f"{camp_aj['reforecast']-camp_orig['reforecast']:+,.0f} pç")
            i2.metric("Pallets",
                       fn(camp_aj['pallets']),
                       delta=f"{camp_aj['pallets']-camp_orig['pallets']:+,.0f}")
            i3.metric("DN (nova)",  camp_aj['dn'].strftime('%d/%m/%Y'))
            i4.metric("DO (nova)",  camp_aj['do_'].strftime('%d/%m/%Y'))

    # Lista de ajustes ativos
    if ajustes:
        st.markdown('<div class="sec">Ajustes ativos</div>', unsafe_allow_html=True)
        for aid, av in ajustes.items():
            camp_row = df[df['id']==aid]
            nome = camp_row.iloc[0]['campanha'] if not camp_row.empty else str(aid)
            st.markdown(
                f"**{nome}** · SD: {av.get('sd','—')} · ED: {av.get('ed','—')} · "
                f"Est: {fn(av.get('estoque',0))} · FC: {fn(av.get('forecast',0))}")

# ══════════════════════════════════════════════════════════════════════════════
# DADOS
# ══════════════════════════════════════════════════════════════════════════════
with t_dados:
    st.markdown("#### Arquivo carregado")
    g1,g2,g3 = st.columns(3)
    g1.metric("Total campanhas ODP", len(df))
    g2.metric("Campanhas ON", len(df_on))
    g3.metric("Referência", data_ref.strftime('%d/%m/%Y'))

    st.markdown('<div class="sec">Resumo por status</div>', unsafe_allow_html=True)
    resumo = (df.groupby('status_op')
                .agg(campanhas=('id','count'),
                     estoque_total=('estoque','sum'),
                     pallets_total=('pallets','sum'))
                .reset_index()
                .rename(columns={'status_op':'Status','campanhas':'Campanhas',
                                  'estoque_total':'Estoque Total','pallets_total':'Pallets'}))
    st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown('<div class="sec">Todas as campanhas</div>', unsafe_allow_html=True)
    vis = df[['id','campanha','status_op','cat','cd','sd','ed','wd',
              'estoque','pallets','forecast','so_cong','reforecast','novo_so']].copy()
    vis.columns = ['ID','Campanha','Status','Categoria','CD','SD','ED','WD',
                   'Estoque','Pallets','Forecast','SO Cong.','Reforecast','Novo SO']
    st.dataframe(vis, use_container_width=True, hide_index=True)
    st.download_button("⬇ Exportar todas (.csv)",
        vis.to_csv(index=False,sep=';').encode('utf-8-sig'),
        "campanhas_odp.csv","text/csv")
