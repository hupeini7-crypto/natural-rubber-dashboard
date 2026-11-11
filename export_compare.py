"""
export_compare.py —— 五国出口分析 + 中国进口/五国出口对比 模块
==============================================================
被 seasonal_dashboard.build() 在运行时调用（不破坏原有进口模块）。

数据源（唯一本地基础）：
  - 中国进口：cfg.file_path（天然橡胶进口合并报表.xlsx · 按国别）
  - 五国出口：cfg.export_file_path（五国橡胶出口数据.xlsx · 按国别 / 主要目的地（仅天然橡胶））

单位约定：进口/出口原始均为 kg，统一除以 cfg.kg2wan (=1e7) 折算为 万吨。
统计原则：
  - 进口与出口严格按「国家 + 月份 + 品种」匹配
  - 天然橡胶 / 混合橡胶 分别比较；总量 = 天然 + 混合
  - 中国进口 ÷ 该国全球出口 是「跨口径比值」，仅反映相对规模，并非实际输华比例
  - 真实对华出口占比 = 输华天然橡胶(主要目的地中国列) ÷ 该国天然橡胶全球出口
  - 数据缺失一律留空(NaN)，不填 0、不编造
"""
from __future__ import annotations
import json
import logging
from typing import Dict, List, Tuple

import pandas as pd
import numpy as np
import plotly.graph_objects as go

from seasonal_dashboard import Config, build_combo_fig, discover_country_metrics

logger = logging.getLogger("seasonal_dashboard.export_compare")

EXPORT_COUNTRIES = ['泰国', '马来西亚', '越南', '印尼', '科特迪瓦']
# 五国配色（与进口模块风格一致）
COUNTRY_COLORS = {
    '泰国': '#e02020', '马来西亚': '#2ca6a4', '越南': '#f28e2b',
    '印尼': '#4e79a7', '科特迪瓦': '#7b52ab',
}


# ===================== 数据加载 =====================
def load_export_data(cfg: Config, window: bool = True) -> pd.DataFrame:
    """读取出口 Excel「按国别」sheet，返回窗口内(万吨)宽表：
    日期/年份/月份 + exp_{c}_nat / exp_{c}_mix / exp_{c}_total
    注意：科特迪瓦无 混合橡胶 列，其 mix=NaN，total 用 sheet 提供的「总出口」(=天然)。
    window=True 时按 lookback_years 截取最近 N 年（7.x 图表用）；window=False 返回全历史（8.0 明细表用）。"""
    df = pd.read_excel(cfg.export_file_path, sheet_name=cfg.export_sheet, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    df['年份'] = df['日期'].dt.year
    df['月份'] = df['日期'].dt.month

    out = {'日期': df['日期'].values, '年份': df['年份'].values, '月份': df['月份'].values}
    for cn in EXPORT_COUNTRIES:
        nat = mix = total = None
        for col in df.columns:
            if not isinstance(col, tuple) or len(col) != 2:
                continue
            l0, l1 = col
            if not (isinstance(l0, str) and isinstance(l1, str) and cn in l0):
                continue
            if '天然橡胶' in l1:
                nat = col
            elif '混合橡胶' in l1:
                mix = col
            elif '总出口' in l1:
                total = col
        s_nat = (pd.to_numeric(df[nat], errors='coerce') / cfg.kg2wan) if nat is not None else pd.Series(float('nan'), index=df.index)
        s_mix = (pd.to_numeric(df[mix], errors='coerce') / cfg.kg2wan) if mix is not None else pd.Series(float('nan'), index=df.index)
        if total is not None:
            s_total = pd.to_numeric(df[total], errors='coerce') / cfg.kg2wan
        else:
            s_total = (s_nat.fillna(0) + s_mix.fillna(0))
            s_total[s_nat.isna() & s_mix.isna()] = float('nan')
        out[f'exp_{cn}_nat'] = s_nat.values
        out[f'exp_{cn}_mix'] = s_mix.values
        out[f'exp_{cn}_total'] = s_total.values

    res = pd.DataFrame(out)
    if not window:
        logger.info(f"[出口] 加载 {len(res)} 行（全历史，8.0 明细表用）")
        return res
    latest_year = int(res['年份'].max())
    start_year = latest_year - cfg.lookback_years
    res_w = res[res['年份'] >= start_year].copy()
    logger.info(f"[出口] 加载 {len(res)} 行，窗口 {start_year}-{latest_year}")
    return res_w


def load_export_dest_data(cfg: Config) -> pd.DataFrame:
    """读取出口 Excel「主要目的地（仅天然橡胶）」sheet，提取各国 -> 中国 出口量(万吨)。"""
    df = pd.read_excel(cfg.export_file_path, sheet_name=cfg.export_dest_sheet, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    df['年份'] = df['日期'].dt.year
    df['月份'] = df['日期'].dt.month
    out = {'日期': df['日期'].values, '年份': df['年份'].values, '月份': df['月份'].values}
    for cn in EXPORT_COUNTRIES:
        target = None
        for col in df.columns:
            if not isinstance(col, tuple) or len(col) != 2:
                continue
            l0, l1 = col
            if not (isinstance(l0, str) and isinstance(l1, str)):
                continue
            if (cn in l0) and ('中国' in l1):
                target = col
                break
        if target is not None:
            out[f'dest_{cn}_to_china'] = (pd.to_numeric(df[target], errors='coerce') / cfg.kg2wan).values
        else:
            out[f'dest_{cn}_to_china'] = float('nan')
    return pd.DataFrame(out)


def load_fine_split(cfg: Config) -> Tuple[pd.DataFrame, str]:
    """读取进口 Excel「按细分类」sheet，返回全国 天然/混合/总量(万吨) 口径。
    若「总计进口量」列全空则 fallback 到 按国别 10 国加总（返回 source 标注）。"""
    df = pd.read_excel(cfg.file_path, sheet_name=cfg.fine_sheet, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    col_nat = df.columns[6]   # ('天然橡胶合计','总计进口量')
    col_mix = df.columns[7]   # ('天然橡胶合成橡胶合计\n400280','进口量')
    nat = pd.to_numeric(df[col_nat], errors='coerce')
    mix = pd.to_numeric(df[col_mix], errors='coerce')
    if nat.notna().sum() == 0:
        # fallback：10 国加总（与 seasonal_dashboard.load_fine_data 同口径）
        from seasonal_dashboard import load_data, discover_country_metrics, compute_total_columns
        dfw = load_data(cfg)
        _, cm = discover_country_metrics(cfg, dfw)
        dfw = compute_total_columns(cfg, dfw, cm)
        tcols = [c for c in dfw.columns if isinstance(c, str) and c.startswith('总量_')]
        nat_s = dfw[tcols].sum(axis=1, min_count=1)
        # 混合部分：总量中混合橡胶分项未单独保留，这里用总量近似（仅 fallback 场景）
        out = pd.DataFrame({
            '日期': dfw['日期'].values,
            '全国天然': (nat_s.fillna(0) / cfg.kg2wan).values,
            '全国混合': float('nan'),
        })
        out['全国总量'] = out['全国天然']
        out['年份'] = dfw['年份'].values
        out['月份'] = dfw['月份'].values
        return out, 'fallback(10国加总)'
    out = pd.DataFrame({
        '日期': df['日期'].values,
        '全国天然': (nat.fillna(0) / cfg.kg2wan).values,
        '全国混合': (mix.fillna(0) / cfg.kg2wan).values,
    })
    out['全国总量'] = out['全国天然'] + out['全国混合']
    out['年份'] = df['日期'].dt.year.values
    out['月份'] = df['日期'].dt.month.values
    return out, '按细分类sheet'


def _load_full_import(cfg: Config) -> pd.DataFrame:
    """读取进口 Excel「按国别」sheet 的【全历史】（不去窗口），返回五国
    天然/混合/总量(万吨)：imp_{c}_nat / imp_{c}_mix / imp_{c}_total。
    口径与 seasonal_dashboard.load_data 一致（天然橡胶+混合橡胶，两者皆缺才为 NaN）。"""
    df = pd.read_excel(cfg.file_path, sheet_name=cfg.sheet_name, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    df['年份'] = df['日期'].dt.year
    df['月份'] = df['日期'].dt.month
    _, cm = discover_country_metrics(cfg, df)
    out = {'日期': df['日期'].values, '年份': df['年份'].values, '月份': df['月份'].values}
    for c in EXPORT_COUNTRIES:
        nat, mix, tot = _country_import_series(df, cm, c, cfg)
        out[f'imp_{c}_nat'] = nat.values
        out[f'imp_{c}_mix'] = mix.values
        out[f'imp_{c}_total'] = tot.values
    return pd.DataFrame(out)


# ===================== 通用工具 =====================
def _country_import_series(df_w, country_metrics, country, cfg):
    """返回 (nat, mix, total) Series（万吨，对齐 df_w 行），用于五国进口侧。"""
    items = country_metrics.get(country, [])
    nat_col = next((c for nm, c in items if nm == '天然橡胶中国进口量'), None)
    mix_col = next((c for nm, c in items if nm == '混合橡胶中国进口量'), None)
    nat = pd.to_numeric(df_w[nat_col], errors='coerce') / cfg.kg2wan if nat_col else pd.Series(float('nan'), index=df_w.index)
    mix = pd.to_numeric(df_w[mix_col], errors='coerce') / cfg.kg2wan if mix_col else pd.Series(float('nan'), index=df_w.index)
    total = (nat.fillna(0) + mix.fillna(0))
    total[nat.isna() & mix.isna()] = float('nan')
    return nat.reset_index(drop=True), mix.reset_index(drop=True), total.reset_index(drop=True)


def _pivot_sum(vals, dates):
    """按(月份, 年份)求和透视；min_count=1 保证整组全为缺失时保持 NaN，绝不补 0。
    （pivot_table(aggfunc='sum') 会把全 NaN 组聚合成 0，导致未更新月份折线掉到 0 线）"""
    d = pd.to_datetime(dates)
    tmp = pd.DataFrame({'年份': d.dt.year.values, '月份': d.dt.month.values,
                        'v': pd.to_numeric(pd.Series(list(vals)), errors='coerce').values})
    return tmp.groupby(['月份', '年份'])['v'].sum(min_count=1).unstack('年份')


def _mkpivot(vals, dates, agg='sum'):
    if agg == 'sum':
        return _pivot_sum(vals, dates)
    d = pd.to_datetime(dates)
    tmp = pd.DataFrame({'年份': d.dt.year.values, '月份': d.dt.month.values, 'v': vals})
    return tmp.pivot_table(index='月份', columns='年份', values='v', aggfunc=agg)


def _val_at(series, dates, y, m):
    s = pd.Series(series.values if hasattr(series, 'values') else series,
                  index=pd.to_datetime(dates).values)
    sub = s[(s.index.year == y) & (s.index.month == m)]
    if len(sub) == 0:
        return float('nan')
    v = float(sub.iloc[0])
    return v


def _sum_vals(vals):
    """忽略 None / NaN 求和；全为空返回 NaN（绝不补 0）。"""
    s = [float(v) for v in vals if v is not None and not (isinstance(v, float) and pd.isna(v))]
    return sum(s) if s else float('nan')


def _fmt(v, unit='万吨'):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return '-'
    if unit == '万吨':
        return f'{v:,.2f}'
    if unit == '%':
        return f'{v:+.1f}%'
    if unit == 'pp':
        return f'{v:+.2f}pp'
    return f'{v:g}'


def _cell_color(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if v > 0:
        return 'background:#fff1f0;color:#c0392b;font-weight:600;'
    if v < 0:
        return 'background:#f6ffed;color:#27ae60;font-weight:600;'
    return ''


# ===================== 图表辅助 =====================
def fig_countries_total(exp_df, metric, title, cumulative=False):
    """五国总出口量对比：x=月份，每国一条线（最新年）+ 五国合计线。"""
    years = sorted(exp_df['年份'].unique().tolist())
    latest = years[-1]
    fig = go.Figure()
    # 各国序列（缺失月保持 NaN，折线自然断开，不补 0）
    series_map = {}
    for c in EXPORT_COUNTRIES:
        s = exp_df[f'exp_{c}_{metric}']
        pv = _mkpivot(s.values, exp_df['日期'])
        if latest not in pv.columns:
            continue
        series_map[c] = pv[latest]
    # 五国合计：仅当该月五国均有数据时才汇总（否则 NaN，避免某国未更新导致合计虚降/虚增）
    total_all = None
    if series_map:
        total_all = pd.DataFrame(series_map).sum(axis=1, min_count=len(series_map))
    for c, y in series_map.items():
        if cumulative:
            y = y.cumsum()
        fig.add_trace(go.Scatter(
            x=y.index, y=y.values, mode='lines+markers', name=c,
            line=dict(color=COUNTRY_COLORS.get(c, '#999'), width=2.4),
            marker=dict(size=5),
            hovertemplate=f'<b>{c}</b><br>%{{x}}月: %{{y:,.2f}} 万吨<extra></extra>'))
    if total_all is not None:
        if cumulative:
            total_all = total_all.cumsum()
        fig.add_trace(go.Scatter(
            x=total_all.index, y=total_all.values, mode='lines+markers', name='五国合计',
            line=dict(color='#222', width=3.2, dash='dot'),
            marker=dict(size=6, symbol='diamond'),
            hovertemplate='<b>五国合计</b><br>%{x}月: %{y:,.2f} 万吨<extra></extra>'))
    fig.update_layout(
        title=dict(text=f'<b>{title}</b>', x=0.01, xanchor='left',
                   font=dict(size=15, family='Microsoft YaHei, Arial', color='#2c3e50')),
        height=420, template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=11, color='#444'),
        margin=dict(l=58, r=24, t=52, b=48), hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.03, x=0, xanchor='left',
                    font=dict(size=11, family='Microsoft YaHei, Arial')),
        plot_bgcolor='#fafcff')
    fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)),
                     ticktext=[f'{m:02d}' for m in range(1, 13)])
    fig.update_yaxes(title_text='万吨', gridcolor='#eee')
    return fig


def fig_import_vs_export(dates, imp, exp, title, cumulative=False, color_imp='#e02020', color_exp='#2ca6a4',
                         max_month=12):
    """中国进口 vs 该国全球出口 双线对比（最新年份逐月真实值，1..max_month）。
    max_month：只显示 1..max_month（最新有效月份），避免把未来月份画出。
    注意：展示的是最新年份各月的实际值（与 8.0/8.1 明细口径一致），
    不是历年同月加总，避免数值量级与明细表对不上。"""
    d = pd.Series(pd.to_datetime(dates))
    dfx = pd.DataFrame({'月份': d.dt.month.values, '年份': d.dt.year.values, 'imp': imp, 'exp': exp})
    ly = dfx['年份'].max()  # 最新年份（进口/出口各自的最新年份）
    cur = dfx[dfx['年份'] == ly].set_index('月份').sort_index().loc[:max_month]
    di = pd.DataFrame({'imp': cur['imp'].values, 'exp': cur['exp'].values}, index=cur.index)
    if cumulative:
        di['imp'] = di['imp'].cumsum()
        di['exp'] = di['exp'].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=di.index, y=di['imp'].values, mode='lines+markers',
                             name='中国进口', line=dict(color=color_imp, width=2.6),
                             hovertemplate='中国进口 %{x}月: %{y:,.2f} 万吨<extra></extra>'))
    fig.add_trace(go.Scatter(x=di.index, y=di['exp'].values, mode='lines+markers',
                             name='全球出口', line=dict(color=color_exp, width=2.6, dash='dash'),
                             hovertemplate='全球出口 %{x}月: %{y:,.2f} 万吨<extra></extra>'))
    fig.update_layout(
        title=dict(text=f'<b>{title}</b>', x=0.01, xanchor='left',
                   font=dict(size=14, family='Microsoft YaHei, Arial', color='#2c3e50')),
        height=360, template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=11, color='#444'),
        margin=dict(l=58, r=24, t=52, b=48), hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.03, x=0, xanchor='left',
                    font=dict(size=11, family='Microsoft YaHei, Arial')),
        plot_bgcolor='#fafcff')
    fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)),
                     ticktext=[f'{m:02d}' for m in range(1, 13)])
    fig.update_yaxes(title_text='万吨', gridcolor='#eee')
    return fig


def fig_two_lines(x_a, a, x_b, b, title, name_a, name_b, unit,
                  color_a='#e02020', color_b='#2ca6a4', height=300, small=True):
    """通用双线图：x 直接为月份序号(1-12)，无需日期转换。"""
    pa = pd.DataFrame({'月份': list(x_a), 'v': list(a)}).groupby('月份').sum(min_count=1)
    pb = pd.DataFrame({'月份': list(x_b), 'v': list(b)}).groupby('月份').sum(min_count=1)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=pa.index, y=pa['v'].values, mode='lines+markers',
                             name=name_a, line=dict(color=color_a, width=2.4),
                             hovertemplate=f'{name_a} %{{x}}月: %{{y:+.1f}}{unit}<extra></extra>'))
    fig.add_trace(go.Scatter(x=pb.index, y=pb['v'].values, mode='lines+markers',
                             name=name_b, line=dict(color=color_b, width=2.4, dash='dash'),
                             hovertemplate=f'{name_b} %{{x}}月: %{{y:+.1f}}{unit}<extra></extra>'))
    fig.update_layout(
        title=dict(text=f'<b>{title}</b>', x=0.01, xanchor='left',
                   font=dict(size=14 if not small else 13, family='Microsoft YaHei, Arial', color='#2c3e50')),
        height=height, template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=11, color='#444'),
        margin=dict(l=58, r=24, t=52, b=48), hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.03, x=0, xanchor='left',
                    font=dict(size=11, family='Microsoft YaHei, Arial')),
        plot_bgcolor='#fafcff')
    fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)),
                     ticktext=[f'{m:02d}' for m in range(1, 13)])
    fig.update_yaxes(title_text=unit, gridcolor='#eee', ticksuffix=unit if unit == '%' else '')
    return fig


# ===================== 主入口 =====================
def build_export_sections(cfg: Config, env: dict) -> Tuple[str, int, dict]:
    """生成「七、五国出口分析」「八、中国进口与五国出口对比」两段 HTML。

    返回 (html, n_figs, report_dict)。
    report_dict 含：export_latest_months / common_month / biggest_ratio_change 等。
    """
    df_w = env['df_w']
    country_metrics = env['country_metrics']
    latest_lbl = env['latest_lbl']
    prev_lbl = env['prev_lbl']
    yoy_lbl = env['yoy_lbl']
    start_year = env['start_year']
    latest_year = env['latest_year']
    latest_month = env['latest_month']
    exp_df = env['exp_df']
    dest_df = env['dest_df']
    fine_split = env['fine_split']
    fine_split_src = env['fine_split_src']

    # ---- 进口侧（五国）序列 ----
    imp = {}
    for c in EXPORT_COUNTRIES:
        nat, mix, tot = _country_import_series(df_w, country_metrics, c, cfg)
        imp[c] = {'nat': nat, 'mix': mix, 'total': tot}

    # ---- 统一月度表 U（外连接 进口 + 出口 + 输华）----
    U = pd.DataFrame({'日期': df_w['日期'].values})
    for c in EXPORT_COUNTRIES:
        U[f'imp_{c}_nat'] = imp[c]['nat'].values
        U[f'imp_{c}_mix'] = imp[c]['mix'].values
        U[f'imp_{c}_total'] = imp[c]['total'].values
    exp_cols = ['日期'] + [f'exp_{c}_{k}' for c in EXPORT_COUNTRIES for k in ('nat', 'mix', 'total')]
    U = U.merge(exp_df[exp_cols], on='日期', how='outer')
    U = U.merge(dest_df, on='日期', how='outer').sort_values('日期').reset_index(drop=True)
    U_dates = U['日期']

    # ---- 共同可比月份（五国均同时存在进口与出口数据的最后月份）----
    last = {}
    for c in EXPORT_COUNTRIES:
        mask = U[f'exp_{c}_total'].notna() & U[f'imp_{c}_total'].notna()
        ok = U.loc[mask, '日期']
        if len(ok):
            last[c] = ok.max()
    common_date = min(last.values()) if last else None
    common_lbl = common_date.strftime('%Y-%m') if common_date is not None else 'N/A'
    cy = common_date.year if common_date is not None else None
    cm = common_date.month if common_date is not None else None
    py, pm = (cy - 1, cm) if cy else (None, None)

    # 各国出口最新月份
    export_latest_months = {}
    for c in EXPORT_COUNTRIES:
        s = exp_df[f'exp_{c}_total'].dropna()
        if len(s):
            export_latest_months[c] = pd.to_datetime(exp_df.loc[s.index, '日期']).max().strftime('%Y-%m')
        else:
            export_latest_months[c] = 'N/A'

    parts = []
    n_figs = 0
    err_notes = []
    latest_note = ' · '.join(f'{c} {m}' for c, m in export_latest_months.items())

    # ================= 七、五国出口分析 =================
    parts.append(f"""
<div class='section'>
  <div class='section-title'>七、五国橡胶出口分析（泰国 / 马来西亚 / 越南 / 印尼 / 科特迪瓦）</div>
  <div class='legend-box'>
    <b>数据源：</b>五国橡胶出口数据.xlsx · 按国别 sheet（原始单位 kg，已折算 万吨）· 窗口 {start_year}-{latest_year}<br>
    <b>口径：</b>天然橡胶出口 / 混合橡胶出口 / 总出口(=天然+混合) · <b>科特迪瓦无混合橡胶出口列</b>（混合=缺失，总出口=天然）<br>
    <b>各国出口最新月份：</b>{latest_note}<br>
    <b>配色：</b><span style='color:#c0392b;'>上涨浅红</span> · <span style='color:#27ae60;'>下跌浅绿</span> · "-" 表示缺失/无数据
  </div>
""")

    # ---- 7.1 同比/环比变化表（各国按自身最新月，不强制对齐、缺失不补 0）----
    def _country_latest_ym(c):
        s = exp_df[f'exp_{c}_total'].dropna()
        if len(s) == 0:
            return None, None
        d = pd.to_datetime(exp_df.loc[s.index, '日期']).max()
        return d.year, d.month

    cmp_rows = []
    for c in EXPORT_COUNTRIES:
        ccy, ccm = _country_latest_ym(c)
        c_lbl = f'{ccy}-{ccm:02d}' if ccy else 'N/A'
        for metric, col in [('天然橡胶', f'exp_{c}_nat'), ('混合橡胶', f'exp_{c}_mix'), ('总出口', f'exp_{c}_total')]:
            cur = _val_at(exp_df[col], exp_df['日期'], ccy, ccm) if ccy else float('nan')
            # 环比：上一月（处理跨年）；同比：去年同月
            if ccy and ccm:
                prv = _val_at(exp_df[col], exp_df['日期'], ccy - 1, 12) if ccm == 1 else _val_at(exp_df[col], exp_df['日期'], ccy, ccm - 1)
                prv_lbl = f'{ccy-1}-12' if ccm == 1 else f'{ccy}-{ccm-1:02d}'
                yoy = _val_at(exp_df[col], exp_df['日期'], ccy - 1, ccm)
            else:
                prv = float('nan'); prv_lbl = 'N/A'; yoy = float('nan')
            mom = (cur - prv) / prv * 100 if (prv is not None and not pd.isna(prv) and prv != 0) else None
            yoy_pct = (cur - yoy) / yoy * 100 if (yoy is not None and not pd.isna(yoy) and yoy != 0) else None
            cmp_rows.append({'国家': c, '指标': metric, '最新月': c_lbl,
                             '最新': cur, '次新月': prv_lbl, '次新': prv,
                             '环比(%)': round(mom, 1) if mom is not None else None,
                             '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None})
    # 五国合计行：按共同可比月（五国均有数据的最后一月）汇总，避免跨月相加失真
    # 注：科特迪瓦无「混合橡胶」出口列，故五国合计的「混合橡胶」项剔除科特迪瓦，标注「不含科特」
    for metric, key, excl in [('天然橡胶', 'nat', []), ('混合橡胶(不含科特)', 'mix', ['科特迪瓦']), ('总出口', 'total', [])]:
        _cnt = [c for c in EXPORT_COUNTRIES if c not in excl]
        cur = _sum_vals([_val_at(exp_df[f'exp_{c}_{key}'], exp_df['日期'], cy, cm) for c in _cnt]) if cy else float('nan')
        prv = _sum_vals([(_val_at(exp_df[f'exp_{c}_{key}'], exp_df['日期'], cy - 1, 12) if cm == 1 else _val_at(exp_df[f'exp_{c}_{key}'], exp_df['日期'], cy, cm - 1)) for c in _cnt]) if cy else float('nan')
        prv_lbl = f'{cy-1}-12' if cm == 1 else f'{cy}-{cm-1:02d}'
        yoy = _sum_vals([_val_at(exp_df[f'exp_{c}_{key}'], exp_df['日期'], py, pm) for c in _cnt]) if py else float('nan')
        mom = (cur - prv) / prv * 100 if (prv is not None and not pd.isna(prv) and prv != 0) else None
        yoy_pct = (cur - yoy) / yoy * 100 if (yoy is not None and not pd.isna(yoy) and yoy != 0) else None
        cmp_rows.append({'国家': '五国合计', '指标': metric, '最新月': common_lbl,
                         '最新': cur, '次新月': prv_lbl, '次新': prv,
                         '环比(%)': round(mom, 1) if mom is not None else None,
                         '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None})

    rows_html = ''
    for r in cmp_rows:
        is_sum = r['国家'] == '五国合计'
        is_total = r['指标'] == '总出口'   # 各国「总出口」行加粗
        if is_sum:
            bold = 'font-weight:700;background:#f0f7ff;'
        elif is_total:
            bold = 'font-weight:700;'
        else:
            bold = ''
        rows_html += f"""<tr style="{bold}">
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td style="text-align:left;color:#555;">{r['指标']}</td>
            <td>{r['最新月']}</td>
            <td>{_fmt(r['最新'])}</td>
            <td><span style="color:#999;font-size:11px;">{r['次新月']} </span>{_fmt(r['次新'])}</td>
            <td style="{_cell_color(r['环比(%)'])}">{_fmt(r['环比(%)'],'%')}</td>
            <td style="{_cell_color(r['同比(%)'])}">{_fmt(r['同比(%)'],'%')}</td>
        </tr>"""
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>7.1 五国出口 同比 / 环比（各国按自身最新月 · 五国合计按共同可比月 <b>{common_lbl}</b>）</div>
  <div class='legend-box'>各国出口更新月份不同步，下表按<b>各国家自身最新月</b>展示最新值 / 环比 / 同比，<b>不强制跨月对齐、缺失不补 0</b>；「五国合计」行为可在同一月比较的<b>共同可比月 {common_lbl}</b>汇总。</div>
  <table>
    <thead><tr><th>国家</th><th>指标</th><th>最新月</th><th>最新值(万吨)</th><th>次新值(万吨)</th><th>环比</th><th>同比</th></tr></thead>
    <tbody>{rows_html}</tbody>
  </table>
""")

    # ---- 7.2 出口同比涨跌速览 TOP5（各国按自身最新月）----
    sv = [r for r in cmp_rows if r['同比(%)'] is not None and r['国家'] != '五国合计']
    sv_sorted = sorted(sv, key=lambda x: x['同比(%)'], reverse=True)
    top_up = sv_sorted[:5]
    top_down = sv_sorted[::-1][:5]

    def _speed_rows(rows):
        h = ''
        for r in rows:
            v = _fmt(r['最新'])
            pct = r['同比(%)']
            cls = 'up' if pct > 0 else ('down' if pct < 0 else '')
            arrow = '▲' if pct > 0 else ('▼' if pct < 0 else '—')
            h += f"""<tr>
                <td style="text-align:left;font-weight:500;">{r['国家']}</td>
                <td style="text-align:left;color:#555;">{r['指标']}</td>
                <td>{r['最新月']}</td>
                <td>{v}</td>
                <td class="pct-{cls}">{arrow} {_fmt(pct,'%')}</td>
            </tr>"""
        return h

    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>7.2 五国出口 同比涨跌速览（各国按自身最新月）</div>
  <div class='speed-grid'>
    <div class='speed-card'>
      <div class='speed-title up-title'>▲ 出口同比涨幅 TOP5</div>
      <table class='speed-table'>
        <thead><tr><th>国家</th><th>指标</th><th>最新月</th><th>最新值</th><th>同比</th></tr></thead>
        <tbody>{_speed_rows(top_up)}</tbody>
      </table>
    </div>
    <div class='speed-card'>
      <div class='speed-title down-title'>▼ 出口同比跌幅 TOP5</div>
      <table class='speed-table'>
        <thead><tr><th>国家</th><th>指标</th><th>最新月</th><th>最新值</th><th>同比</th></tr></thead>
        <tbody>{_speed_rows(top_down)}</tbody>
      </table>
    </div>
  </div>
""")

    # ---- 7.3 五国出口总量对比（月度 + 累计）----
    fig_m = fig_countries_total(exp_df, 'total', f'五国总出口量 · 月度对比（{latest_year}）', cumulative=False)
    fig_c = fig_countries_total(exp_df, 'total', f'五国总出口量 · 年内累计对比（{latest_year}）', cumulative=True)
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>7.3 五国出口总量对比（月度 + 年内累计）</div>
  <div class='legend-box'>各国更新月份不同步，每国按<b>自身最新月</b>显示：<b>{latest_note}</b>；未更新月份折线自然断开，<b>不补 0</b>。「五国合计」仅在该月五国数据齐备时汇总（截至共同可比月 <b>{common_lbl}</b>），因此不在 09 月虚降。</div>
  <div class='pair-row'>
    <div class='pair-cell'>{fig_m.to_html(include_plotlyjs='cdn', full_html=False, div_id='exp_total_m')}</div>
    <div class='pair-cell'>{fig_c.to_html(include_plotlyjs='cdn', full_html=False, div_id='exp_total_c')}</div>
  </div>
""")
    n_figs += 2

    # ---- 7.4 各国出口季节性（总出口·历史年份对比）----
    cells = []
    for c in EXPORT_COUNTRIES:
        pv = _mkpivot(exp_df[f'exp_{c}_total'].values, exp_df['日期'])
        fig = build_combo_fig(pv, '', '万吨', cfg, height=280, hover_label='总出口', small=True, show_legend=False)
        cells.append(f"""<div class='grid-cell'>
            <div class='cell-head'>{c} 总出口<span class='cell-unit'>万吨</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'exp_season_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>7.4 五国出口季节性（各国总出口·历史年份对比）</div>
  <div class='legend-box'>每格 = 某国总出口的月度季节性走势（{start_year}-{latest_year}），{latest_year} 红色加粗，灰柱=同比</div>
  <div class='share-grid'>{''.join(cells)}</div>
""")

    # ---- 7.5 各国出口占五国合计比例 ----
    # 计算每月 各国/五国合计（total）
    share_df = pd.DataFrame({'日期': exp_df['日期'].values})
    _tot_map = {}
    for c in EXPORT_COUNTRIES:
        share_df[f's_{c}'] = exp_df[f'exp_{c}_total'].values
        _tot_map[c] = exp_df[f'exp_{c}_total']
    # 分母：仅当该月五国均有数据时才算五国合计（缺失月保持 NaN，不补 0）
    share_df['tot5'] = pd.DataFrame(_tot_map).sum(axis=1, min_count=len(EXPORT_COUNTRIES)).values
    for c in EXPORT_COUNTRIES:
        share_df[f'sh_{c}'] = (share_df[f's_{c}'] / share_df['tot5'] * 100).replace([np.inf, -np.inf], np.nan)

    # 表（最新可比月 + 累计占比 + 同比变化）
    share_rows = []
    for c in EXPORT_COUNTRIES:
        cur = _val_at(share_df[f'sh_{c}'], share_df['日期'], cy, cm) if cy else float('nan')
        # 累计占比：当年1月~common月 各国累计 / 五国合计累计
        d = pd.to_datetime(share_df['日期'])
        mask_c = (d.dt.year == cy) & (d.dt.month <= cm) if cy else d.dt.year == -1
        cum_c = share_df.loc[mask_c, f's_{c}'].sum()
        cum_5 = share_df.loc[mask_c, 'tot5'].sum()
        sh_cum = cum_c / cum_5 * 100 if cum_5 else float('nan')
        # 去年同期累计
        mask_l = (d.dt.year == (cy - 1)) & (d.dt.month <= cm) if cy else d.dt.year == -1
        cum_c_l = share_df.loc[mask_l, f's_{c}'].sum()
        cum_5_l = share_df.loc[mask_l, 'tot5'].sum()
        sh_cum_l = cum_c_l / cum_5_l * 100 if cum_5_l else float('nan')
        chg = sh_cum - sh_cum_l if (not pd.isna(sh_cum) and not pd.isna(sh_cum_l)) else None
        share_rows.append({'国家': c, '单月占比(%)': cur, '累计(万吨)': cum_c,
                           '累计占比(%)': sh_cum, '去年同期累计占比(%)': sh_cum_l,
                           '变化(pp)': round(chg, 2) if chg is not None else None})
    sh_tbl = ''
    for r in share_rows:
        sh_tbl += f"""<tr>
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td style="font-weight:600;">{_fmt(r['单月占比(%)'])}</td>
            <td>{_fmt(r['累计(万吨)'])}</td>
            <td style="font-weight:600;">{_fmt(r['累计占比(%)'])}</td>
            <td>{_fmt(r['去年同期累计占比(%)'])}</td>
            <td style="{_cell_color(r['变化(pp)'])}">{_fmt(r['变化(pp)'],'pp')}</td>
        </tr>"""
    # 条形图
    bar_df = pd.DataFrame(share_rows)
    fig_bar = go.Figure()
    bar_cum = bar_df['累计占比(%)']
    fig_bar.add_trace(go.Bar(y=bar_df['国家'], x=bar_df['单月占比(%)'], orientation='h',
                             name=f'{common_lbl} 单月占比', marker_color='#4e79a7',
                             hovertemplate='<b>%{y}</b> 单月占比: %{x:.2f}%<extra></extra>'))
    fig_bar.add_trace(go.Bar(y=bar_df['国家'], x=bar_cum, orientation='h',
                             name=f'{common_lbl} 累计占比', marker_color='#f28e2b',
                             hovertemplate='<b>%{y}</b> 累计占比: %{x:.2f}%<extra></extra>'))
    fig_bar.update_layout(barmode='group', height=400, template='plotly_white',
                          font=dict(family='Microsoft YaHei, Arial', size=11, color='#444'),
                          margin=dict(l=90, r=30, t=52, b=48), xaxis_title='占比 %',
                          legend=dict(orientation='h', yanchor='bottom', y=1.03, x=0, xanchor='left',
                                      font=dict(size=11)),
                          title=dict(text='<b>各国出口占五国合计比重</b>', x=0.01, xanchor='left',
                                     font=dict(size=15, family='Microsoft YaHei, Arial', color='#2c3e50')),
                          plot_bgcolor='#fafcff')
    n_figs += 1
    # 季节性占比格
    sh_cells = []
    for c in EXPORT_COUNTRIES:
        pv = _mkpivot(share_df[f'sh_{c}'].values, share_df['日期'])
        fig = build_combo_fig(pv, '', '%', cfg, height=260, hover_label='占五国合计', small=True, show_legend=False)
        sh_cells.append(f"""<div class='grid-cell'>
            <div class='cell-head'>{c} 占五国合计<span class='cell-unit'>%</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'exp_share_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>7.5 各国出口占五国合计比例（{common_lbl}）</div>
  <table>
    <thead><tr><th>国家</th><th>单月占比 %</th><th>{cy} 年内累计(万吨)</th><th>累计占比 %</th>
      <th>{cy-1 if cy else '-'} 同期累计占比 %</th><th>变化 pp</th></tr></thead>
    <tbody>{sh_tbl}</tbody>
  </table>
  <div style='height:12px;'></div>
  {fig_bar.to_html(include_plotlyjs='cdn', full_html=False, div_id='exp_share_bar')}
  <div style='height:12px;'></div>
  <div class='share-grid'>{''.join(sh_cells)}</div>
""")

    parts.append("</div>")  # 结束 七

    # ---- 8.0 顶部明细：全历史 出口总量 / 中国进口天胶 / 进口对应出口比（2015 年以来，新→旧）----
    _ord = ['泰国', '印尼', '越南', '马来西亚', '科特迪瓦']  # 用户指定列顺序
    try:
        _imp_full = _load_full_import(cfg)
        _exp_full = load_export_data(cfg, window=False)
        M80 = pd.DataFrame({'日期': _imp_full['日期'].values})
        for _c in EXPORT_COUNTRIES:
            M80[f'imp_{_c}'] = _imp_full[f'imp_{_c}_total'].values
        _e80 = _exp_full[['日期'] + [f'exp_{_c}_total' for _c in EXPORT_COUNTRIES]].rename(
            columns={f'exp_{_c}_total': f'exp_{_c}' for _c in EXPORT_COUNTRIES})
        M80 = M80.merge(_e80, on='日期', how='outer').sort_values('日期').reset_index(drop=True)
        _d80 = pd.to_datetime(M80['日期'])
        _any80 = M80[[f'exp_{_c}' for _c in EXPORT_COUNTRIES] +
                     [f'imp_{_c}' for _c in EXPORT_COUNTRIES]].notna().any(axis=1)
        M80 = M80[(_d80 >= pd.Timestamp('2015-01-01')) & _any80].copy()
        M80 = M80.iloc[::-1].reset_index(drop=True)  # 按月份 新 → 旧
        _d80 = pd.to_datetime(M80['日期'])
    except Exception as _e:
        logger.exception(f"[8.0] 明细表数据加载失败: {_e}")
        M80 = pd.DataFrame()

    t80_rows = ''
    t80_data = []
    for _i, (_r, _row) in enumerate(M80.iterrows()):
        _y = int(_d80.loc[_r].year)
        _m = int(_d80.loc[_r].month)
        _exps = [float(_row[f'exp_{c}']) if not pd.isna(_row[f'exp_{c}']) else None for c in _ord]
        _imps = [float(_row[f'imp_{c}']) if not pd.isna(_row[f'imp_{c}']) else None for c in _ord]
        _eK = sum(1 for _v in _exps if _v is not None)   # 出口有数据国家数
        _iK = sum(1 for _v in _imps if _v is not None)   # 进口有数据国家数
        _e5 = sum(_v for _v in _exps if _v is not None) if _eK else None
        _i5 = sum(_v for _v in _imps if _v is not None) if _iK else None
        # 五国比例：分子分母用「当月同时具备进口&出口数据」的国家集合 S，保证同口径
        _S = [_j for _j in range(len(_ord)) if _exps[_j] is not None and _imps[_j] is not None]
        _rK = len(_S)
        _r5 = (sum(_imps[_j] for _j in _S) / sum(_exps[_j] for _j in _S) * 100) if _S else None
        _rats = []
        for _j in range(len(_ord)):
            _e, _im = _exps[_j], _imps[_j]
            _rats.append(_im / _e * 100 if (_e is not None and _im is not None and _e != 0) else None)
        # 单元格
        _td = []
        for _v in _exps:  # ① 各国出口
            _td.append(f"<td class='{'' if _v is not None else 'miss'}'>" + ('-' if _v is None else _fmt(_v)) + '</td>')
        if _e5 is None:
            _td.append("<td class='miss'>-</td>")
        else:
            _miss = [_c for _c, _v in zip(_ord, _exps) if _v is None]
            _sup = f" <sup>{_eK}/5</sup>" if _eK < 5 else ''
            _tip = f'仅 {_eK}/5 国有出口数据（缺失：{"、".join(_miss)}），合计不完整' if _eK < 5 else '五国均有出口数据'
            _td.append(f"<td title='{_tip}'><b>{_fmt(_e5)}</b>{_sup}</td>")
        for _v in _imps:  # ② 各国进口
            _td.append(f"<td class='{'' if _v is not None else 'miss'}'>" + ('-' if _v is None else _fmt(_v)) + '</td>')
        if _i5 is None:
            _td.append("<td class='miss'>-</td>")
        else:
            _miss = [_c for _c, _v in zip(_ord, _imps) if _v is None]
            _sup = f" <sup>{_iK}/5</sup>" if _iK < 5 else ''
            _tip = f'仅 {_iK}/5 国有进口数据（缺失：{"、".join(_miss)}），合计不完整' if _iK < 5 else '五国均有进口数据'
            _td.append(f"<td title='{_tip}'><b>{_fmt(_i5)}</b>{_sup}</td>")
        for _v in _rats:  # ③ 各国比例
            if _v is None:
                _td.append("<td class='miss'>-</td>")
            elif _v > 100:
                _td.append(f"<td class='hi'><b>{_v:.1f}%</b><sup>⚠</sup></td>")
            else:
                _td.append(f"<td>{_v:.1f}%</td>")
        if _r5 is None:  # ③ 五国比例
            _td.append("<td class='miss'>-</td>")
        else:
            _marks = ('⚠' if _r5 > 100 else '') + (f'{_rK}/5' if _rK < 5 else '')
            _sup = f" <sup>{_marks}</sup>" if _marks else ''
            _cls = " class='hi'" if _r5 > 100 else ''
            _td.append(f"<td{_cls}><b>{_r5:.1f}%</b>{_sup}</td>")
        t80_rows += (f"<tr class='{'latest' if _i == 0 else ''}' data-year='{_y}'>"
                     f"<td class='fix-dt'>{_y}-{_m:02d}</td>" + ''.join(_td) + '</tr>')
        t80_data.append({'d': f'{_y}-{_m:02d}', 'y': _y,
                         'exp': _exps, 'imp': _imps, 'rat': _rats,
                         'e5': _e5, 'i5': _i5, 'r5': _r5})

    if not t80_data:
        t80_rows = "<tr><td colspan='19' style='text-align:center;color:#c0392b;'>⚠️ 8.0 明细表数据加载失败，请检查数据文件</td></tr>"
    _years80 = sorted({_r['y'] for _r in t80_data}, reverse=True)
    _y_opts = ''.join(f"<option value='{_y}'>{_y}年</option>" for _y in _years80)
    _t80_json = json.dumps(t80_data, ensure_ascii=False)
    _range_lbl = f"{min(_years80)}-{max(_years80)}" if _years80 else '无数据'
    to_cn_table_html = f"""
  <div class='section-title' style='font-size:15px;'>8.0 五国出口总量 / 中国进口天胶 / 进口对应出口比（明细 · {_range_lbl} · 新→旧）</div>
  <div class='legend-box'>
    <b>数据来源：</b>中国进口数据采用中国海关统计，五国出口数据采用各国海关统计。由于统计口径、商品范围及报关时间不同，相关比例仅供趋势参考。<br>
    <b>口径：</b>① 各国出口总量 = 出口 Excel「按国别」sheet 的「总出口(=天然+混合)」列（<b>含混合橡胶</b>）；② 中国进口天胶 = 天然橡胶进口量 + 混合橡胶进口量（中国海关）；③ 比例 = 中国进口天胶 ÷ 该国出口总量。<br>
    <b>缺失</b>显示 "—"、不补 0；分母缺失或为 0 时不计算比例。<b>五国合计</b>仅在当月有数据的国家范围内加总，不足 5 国以 <sup>n/5</sup> 标注（鼠标悬停可见缺失国，<b>非完整五国合计</b>）；五国比例按当月同时具备进出口数据的国家集合计算，保证分子分母同口径。<br>
    <b>比例超过 100%</b>（<sup>⚠</sup>）系两国海关统计口径、商品范围及转口贸易差异所致，保留真实计算值，不代表实际输华份额。
  </div>
  <div class='chips-row'>
    <span class='chip'>年份筛选：</span>
    <select id='t80Year' onchange='t80Filter()' style='font-size:13px;padding:4px 8px;border-radius:6px;border:1px solid #ccc;'>
      <option value='ALL'>全部年份</option>
      {_y_opts}
    </select>
    <button onclick='t80Export()' style='font-size:13px;padding:4px 10px;border-radius:6px;border:1px solid #4e79a7;background:#4e79a7;color:#fff;cursor:pointer;'>导出数据（CSV · 可用 Excel 打开）</button>
    <span class='chip' style='color:#888;'>默认显示最新月份，可上下滚动查看全部历史、左右滚动浏览 19 列</span>
  </div>
  <style>
    /* 8.0 明细表容器 — 增大高度、添加阴影和圆角 */
    .t80wrap{{max-height:520px;overflow:auto;border:1px solid #d4dae3;border-radius:10px;background:#fff;
             box-shadow:0 2px 12px rgba(0,0,0,0.06);}}
    /* 表格基础 — 字号增大、行高增加 */
    .t80tbl{{border-collapse:separate;border-spacing:0;white-space:nowrap;font-size:13px;width:max-content;
              line-height:1.6;}}
    .t80tbl th,.t80tbl td{{border-bottom:1px solid #e8ecf1;border-right:1px solid #f2f4f7;padding:6px 12px;text-align:right;
                             transition:background .15s;}}
    .t80tbl thead th{{position:sticky;z-index:3;color:#1a2533;font-weight:700;letter-spacing:.3px;}}
    .t80tbl thead tr:first-child th{{top:0;height:32px;box-sizing:border-box;font-size:13px;}}
    .t80tbl thead tr:nth-child(2) th{{top:32px;height:28px;box-sizing:border-box;background:#e8eef5;font-size:12.5px;}}
    /* 三组表头 — 渐变色增强层次感 */
    .t80tbl .hd-g1{{background:linear-gradient(135deg,#7eb8da 0%,#a9c6e8 100%);border-bottom:2px solid #5a9bcf;color:#0d3a5c;}}
    .t80tbl .hd-g2{{background:linear-gradient(135deg,#8ed4aa 0%,#a8dfba 100%);border-bottom:2px solid #5cb882;color:#0d4a2a;}}
    .t80tbl .hd-g3{{background:linear-gradient(135deg,#f5c878 0%,#f6d49b 100%);border-bottom:2px solid #e5a83d;color:#5a3a0d;}}
    /* 固定日期列 — 更醒目 */
    .t80tbl th.fix-dt,.t80tbl td.fix-dt{{position:sticky;left:0;z-index:4;background:#e2e8f0;text-align:left;font-weight:700;
                                           min-width:72px;border-right:2px solid #b8c5d6;}}
    .t80tbl thead th.fix-dt{{z-index:5;background:#cdd9e8;}}
    /* 最新行高亮 */
    .t80tbl tr.latest td{{background:#fff3cc;}}
    .t80tbl tr.latest td.fix-dt{{background:#ffe082;}}
    .t80tbl tr.latest:hover td{{background:#ffeeb0;}}
    /* 缺失/异常值 */
    .t80tbl td.miss{{color:#b0b8c4;font-style:italic;}}
    .t80tbl td sup{{color:#8a94a6;font-size:10px;margin-left:2px;font-weight:400;}}
    .t80tbl td.hi{{color:#c0392b;font-weight:700;}}
    .t80tbl td.hi sup{{color:#c0392b;}}
    /* 行悬停效果 */
    .t80tbl tbody tr:hover td{{background:#f0f6ff;}}
    .t80tbl tbody tr:hover td.fix-dt{{background:#dde4ec;}}
  </style>
  <div class='t80wrap'>
    <table class='t80tbl' id='t80Table'>
      <thead>
        <tr>
          <th rowspan='2' class='fix-dt'>日期</th>
          <th colspan='6' class='hd-g1'>① 各国出口总量 (万吨)</th>
          <th colspan='6' class='hd-g2'>② 中国自各国进口天胶 (万吨)</th>
          <th colspan='6' class='hd-g3'>③ 中国进口对应出口比 (%)</th>
        </tr>
        <tr>
          <th>泰国</th><th>印尼</th><th>越南</th><th>马来西亚</th><th>科特迪瓦</th><th>五国</th>
          <th>泰国</th><th>印尼</th><th>越南</th><th>马来西亚</th><th>科特迪瓦</th><th>五国</th>
          <th>泰国</th><th>印尼</th><th>越南</th><th>马来西亚</th><th>科特迪瓦</th><th>五国</th>
        </tr>
      </thead>
      <tbody id='t80Body'>
        {t80_rows}
      </tbody>
    </table>
  </div>
  <script>
  var T80_DATA = {_t80_json};
  function t80Filter(){{
    var sel = document.getElementById('t80Year');
    var v = sel ? sel.value : 'ALL';
    var wrap = document.querySelector('.t80wrap');
    document.querySelectorAll('#t80Body tr').forEach(function(tr){{
      tr.style.display = (v==='ALL' || tr.getAttribute('data-year')===v) ? '' : 'none';
    }});
    if(wrap){{ wrap.scrollTop = 0; }}
  }}
  function t80Export(){{
    var head = ['日期','泰国出口','印尼出口','越南出口','马来西亚出口','科特迪瓦出口','五国出口','泰国进口','印尼进口','越南进口','马来西亚进口','科特迪瓦进口','五国进口','泰国比例%','印尼比例%','越南比例%','马来西亚比例%','科特迪瓦比例%','五国比例%'];
    var lines = [head.join(',')];
    T80_DATA.forEach(function(r){{
      var f2 = function(v){{ return v===null ? '' : Number(v).toFixed(2); }};
      var f1 = function(v){{ return v===null ? '' : Number(v).toFixed(1); }};
      var cells = [r.d];
      r.exp.forEach(function(v){{ cells.push(f2(v)); }});
      cells.push(f2(r.e5));
      r.imp.forEach(function(v){{ cells.push(f2(v)); }});
      cells.push(f2(r.i5));
      r.rat.forEach(function(v){{ cells.push(f1(v)); }});
      cells.push(f1(r.r5));
      lines.push(cells.join(','));
    }});
    var blob = new Blob(['\ufeff' + lines.join('\n')], {{type:'text/csv;charset=utf-8;'}});
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '进出口对比明细_{_range_lbl}.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }}
  </script>
  <div style='height:10px;'></div>
"""

    # ================= 八、中国进口与五国出口对比 =================
    parts.append(f"""
<div class='section'>
  <div class='section-title'>八、中国进口 × 五国出口 对比</div>
  <div class='legend-box'>
    <b>数据来源：</b>中国进口数据采用中国海关统计，五国出口数据采用各国海关统计。由于统计主体、商品分类、报关时间及贸易统计口径存在差异，进出口数据可能不完全一致，相关比例仅供趋势对比参考。<br>
    <b>可比口径：</b>最新可比月 <b>{common_lbl}</b>（五国均同时存在进口与出口数据的最后月份；进口最新 {latest_lbl}，出口各国最新：{latest_note}）<br>
    <b>品种分别：</b>天然橡胶 / 混合橡胶 / 总量（=天然+混合）各自比较，统一 万吨<br>
    <b>跨口径比值：</b>「中国进口 ÷ 该国全球出口」为跨口径比较，仅反映相对规模，不代表实际输华比例
  </div>
  <div class='chips-row'>
    <span class='chip'>国家筛选：</span>
    <select id='cmpCountry' onchange='cmpFilter()' style='font-size:13px;padding:4px 8px;border-radius:6px;border:1px solid #ccc;'>
      <option value='ALL'>全部国家</option>
      {''.join(f'<option value="{c}">{c}</option>' for c in EXPORT_COUNTRIES)}
    </select>
    <span class='chip' style='color:#888;'>（年份筛选可用图表图例点击切换）</span>
  </div>
  {to_cn_table_html}
""")

    # ---- 8.1 对比总表 ----
    cmp2_rows = []
    ratio_change_info = []
    for c in EXPORT_COUNTRIES:
        for metric, ik, ek in [('天然橡胶', 'nat', 'nat'), ('混合橡胶', 'mix', 'mix'), ('总量', 'total', 'total')]:
            imp_v = _val_at(imp[c][ik], df_w['日期'], cy, cm) if cy else float('nan')
            exp_v = _val_at(exp_df[f'exp_{c}_{ek}'], exp_df['日期'], cy, cm) if cy else float('nan')
            imp_yoy_v = _val_at(imp[c][ik], df_w['日期'], py, pm) if py else float('nan')
            exp_yoy_v = _val_at(exp_df[f'exp_{c}_{ek}'], exp_df['日期'], py, pm) if py else float('nan')
            ratio = imp_v / exp_v * 100 if (not pd.isna(imp_v) and not pd.isna(exp_v) and exp_v != 0) else None
            imp_yoy = (imp_v - imp_yoy_v) / imp_yoy_v * 100 if (not pd.isna(imp_yoy_v) and imp_yoy_v != 0) else None
            exp_yoy = (exp_v - exp_yoy_v) / exp_yoy_v * 100 if (not pd.isna(exp_yoy_v) and exp_yoy_v != 0) else None
            # 比值同比变化（去年同月比值）
            imp_py = _val_at(imp[c][ik], df_w['日期'], py, pm) if py else float('nan')
            exp_py = _val_at(exp_df[f'exp_{c}_{ek}'], exp_df['日期'], py, pm) if py else float('nan')
            ratio_py = imp_py / exp_py * 100 if (not pd.isna(imp_py) and not pd.isna(exp_py) and exp_py != 0) else None
            ratio_chg = ratio - ratio_py if (ratio is not None and ratio_py is not None) else None
            if metric == '总量' and ratio_chg is not None:
                ratio_change_info.append((c, ratio_chg))
            cmp2_rows.append({'国家': c, '指标': metric,
                              '中国进口': imp_v, '全球出口': exp_v, '比值(%)': ratio,
                              '进口同比(%)': round(imp_yoy, 1) if imp_yoy is not None else None,
                              '出口同比(%)': round(exp_yoy, 1) if exp_yoy is not None else None})
    # 五国合计行：天然/总量用全部五国；混合橡胶剔除科特迪瓦（无混合出口），标注「不含科特」
    for metric, ik, ek, excl in [('天然橡胶', 'nat', 'nat', []),
                                 ('混合橡胶(不含科特)', 'mix', 'mix', ['科特迪瓦']),
                                 ('总量', 'total', 'total', [])]:
        _cnt = [c for c in EXPORT_COUNTRIES if c not in excl]
        imp_v = _sum_vals([_val_at(imp[c][ik], df_w['日期'], cy, cm) for c in EXPORT_COUNTRIES]) if cy else float('nan')
        exp_v = _sum_vals([_val_at(exp_df[f'exp_{c}_{ek}'], exp_df['日期'], cy, cm) for c in _cnt]) if cy else float('nan')
        imp_yoy_v = _sum_vals([_val_at(imp[c][ik], df_w['日期'], py, pm) for c in EXPORT_COUNTRIES]) if py else float('nan')
        exp_yoy_v = _sum_vals([_val_at(exp_df[f'exp_{c}_{ek}'], exp_df['日期'], py, pm) for c in _cnt]) if py else float('nan')
        ratio = imp_v / exp_v * 100 if (not pd.isna(imp_v) and not pd.isna(exp_v) and exp_v != 0) else None
        imp_yoy = (imp_v - imp_yoy_v) / imp_yoy_v * 100 if (not pd.isna(imp_yoy_v) and imp_yoy_v != 0) else None
        exp_yoy = (exp_v - exp_yoy_v) / exp_yoy_v * 100 if (not pd.isna(exp_yoy_v) and exp_yoy_v != 0) else None
        cmp2_rows.append({'国家': '五国合计', '指标': metric,
                          '中国进口': imp_v, '全球出口': exp_v, '比值(%)': ratio,
                          '进口同比(%)': round(imp_yoy, 1) if imp_yoy is not None else None,
                          '出口同比(%)': round(exp_yoy, 1) if exp_yoy is not None else None})

    cmp2_html = ''
    for r in cmp2_rows:
        is_sum = r['国家'] == '五国合计'
        is_total = r['指标'] == '总量'
        if is_sum:
            bold = 'font-weight:700;background:#f0f7ff;'
        elif is_total:
            bold = 'font-weight:700;'
        else:
            bold = ''
        cmp2_html += f"""<tr style="{bold}">
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td style="text-align:left;color:#555;">{r['指标']}</td>
            <td>{_fmt(r['中国进口'])}</td>
            <td>{_fmt(r['全球出口'])}</td>
            <td style="font-weight:600;">{_fmt(r['比值(%)'],'%')}</td>
            <td style="{_cell_color(r['进口同比(%)'])}">{_fmt(r['进口同比(%)'],'%')}</td>
            <td style="{_cell_color(r['出口同比(%)'])}">{_fmt(r['出口同比(%)'],'%')}</td>
        </tr>"""
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.1 进口 vs 全球出口 对比总表（{common_lbl}）</div>
  <table>
    <thead><tr><th>国家</th><th>指标</th><th>中国进口</th><th>全球出口</th><th>进口/出口</th>
      <th>进口同比</th><th>出口同比</th></tr></thead>
    <tbody>{cmp2_html}</tbody>
  </table>
  <div class='legend-box' style='margin-top:10px;'>
    「进口/出口」= 中国从该国的进口量 ÷ 该国全球出口量（跨口径比值，仅供参考）
  </div>
""")

    # ---- 8.2 进口 vs 出口 量对比（按国别·按品种），支持国家筛选 ----
    def _cmp_block(metric_label, ik, ek, countries_subset):
        blk = []
        for c in countries_subset:
            # 进口/出口按日期内连接对齐（两者月份范围可能不同步，如出口新增月份），
            # 生成等长的 dates/imp/exp 再画图，避免 DataFrame 长度不一致报错。
            _imp_df = pd.DataFrame({'日期': pd.to_datetime(df_w['日期']).values,
                                    'imp': imp[c][ik].values})
            _exp_df = pd.DataFrame({'日期': pd.to_datetime(exp_df['日期']).values,
                                    f'exp_{c}_{ek}': exp_df[f'exp_{c}_{ek}'].values})
            _al = _imp_df.merge(_exp_df, on='日期', how='inner').dropna(subset=['imp', f'exp_{c}_{ek}'])
            _al_dates = _al['日期'].values
            _al_imp = _al['imp'].values
            _al_exp = _al[f'exp_{c}_{ek}'].values
            fig_m = fig_import_vs_export(_al_dates, _al_imp, _al_exp,
                                         f'中国进口 vs 全球出口 · {c} {metric_label} · 月度', cumulative=False,
                                         color_imp=COUNTRY_COLORS.get(c, '#e02020'), color_exp='#2ca6a4',
                                         max_month=cm if cm else 12)
            fig_c = fig_import_vs_export(_al_dates, _al_imp, _al_exp,
                                         f'中国进口 vs 全球出口 · {c} {metric_label} · 累计', cumulative=True,
                                         color_imp=COUNTRY_COLORS.get(c, '#e02020'), color_exp='#2ca6a4',
                                         max_month=cm if cm else 12)
            blk.append(f"""<div class='cmp-block' data-country='{c}'>
              <div class='pair-row'>
                <div class='pair-cell'>{fig_m.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'cmp_{c}_{ik}_m')}</div>
                <div class='pair-cell'>{fig_c.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'cmp_{c}_{ik}_c')}</div>
              </div>
            </div>""")
        return ''.join(blk)

    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.2 中国进口 vs 全球出口 量对比（按国别·按品种）</div>
  <div class='legend-box'>左：月度量对比；右：年内累计量对比。红=中国进口，青虚线=该国全球出口。可用上方「国家筛选」只看某国。数据仅到最新有效月（{common_lbl}），不显示未来月份。</div>
  <div><b>天然橡胶</b></div>
  {_cmp_block('天然橡胶', 'nat', 'nat', EXPORT_COUNTRIES)}
  <div style='height:6px;'></div>
  <div><b>混合橡胶</b>（科特迪瓦无混合橡胶出口，已跳过）</div>
  {_cmp_block('混合橡胶', 'mix', 'mix', [c for c in EXPORT_COUNTRIES if c != '科特迪瓦'])}
""")
    # 计数：天然 5×2 + 混合 4×2 = 18
    n_figs += len(EXPORT_COUNTRIES) * 2 + (len(EXPORT_COUNTRIES) - 1) * 2

    # ---- 8.3 进口/出口 比值 走势（跨口径）----
    ratio_cells = []
    for c in EXPORT_COUNTRIES:
        # 月度比值 = imp_total / exp_total
        dfr = pd.DataFrame({'日期': df_w['日期'].values,
                            'imp': imp[c]['total'].values})
        dfr = dfr.merge(exp_df[['日期', f'exp_{c}_total']].rename(columns={f'exp_{c}_total': 'exp'}), on='日期', how='inner')
        dfr['ratio'] = (dfr['imp'] / dfr['exp'] * 100).replace([np.inf, -np.inf], np.nan)
        pv = _mkpivot(dfr['ratio'].values, dfr['日期'])
        fig = build_combo_fig(pv, '', '%', cfg, height=280, hover_label='进口/出口', small=True, show_legend=False)
        ratio_cells.append(f"""<div class='cmp-block grid-cell' data-country='{c}'>
            <div class='cell-head'>{c} 进口/全球出口 比值<span class='cell-unit'>%</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'ratio_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.3 中国进口量 ÷ 该国全球出口量 比值走势（跨口径，≠真实对华）</div>
  <div class='legend-box'>比值 = 中国从该国的进口量 ÷ 该国全球出口量（%），反映中国需求占该国出口的「跨口径」相对强度，并非实际输华比例。</div>
  <div class='share-grid'>{''.join(ratio_cells)}</div>
""")

    # ---- 8.4 五国在中国总进口中占比（天然/混合/总量）----
    # 全国口径（窗口内）
    fine_w = fine_split[(fine_split['年份'] >= start_year)].copy()
    share_cn_rows = []
    for c in EXPORT_COUNTRIES:
        imp_nat = _val_at(imp[c]['nat'], df_w['日期'], cy, cm) if cy else float('nan')
        imp_mix = _val_at(imp[c]['mix'], df_w['日期'], cy, cm) if cy else float('nan')
        imp_tot = _val_at(imp[c]['total'], df_w['日期'], cy, cm) if cy else float('nan')
        nat_all = _val_at(fine_w['全国天然'], fine_w['日期'], cy, cm) if cy else float('nan')
        mix_all = _val_at(fine_w['全国混合'], fine_w['日期'], cy, cm) if cy else float('nan')
        tot_all = _val_at(fine_w['全国总量'], fine_w['日期'], cy, cm) if cy else float('nan')
        sh_nat = imp_nat / nat_all * 100 if (not pd.isna(imp_nat) and not pd.isna(nat_all) and nat_all and nat_all != 0) else None
        sh_mix = imp_mix / mix_all * 100 if (not pd.isna(imp_mix) and not pd.isna(mix_all) and mix_all and mix_all != 0) else None
        sh_tot = imp_tot / tot_all * 100 if (not pd.isna(imp_tot) and not pd.isna(tot_all) and tot_all and tot_all != 0) else None
        # 累计占比（用位置对齐的 numpy 掩码，避免索引不一致）
        d_imp = pd.to_datetime(df_w['日期']).values
        if cy:
            m_imp = ((d_imp >= pd.Timestamp(year=cy, month=1, day=1).to_datetime64()) &
                     (d_imp <= pd.Timestamp(year=cy, month=cm, day=1).to_datetime64()))
        else:
            m_imp = np.zeros(len(d_imp), dtype=bool)
        cum_nat = float(np.nansum(imp[c]['nat'].values[m_imp]))
        cum_mix = float(np.nansum(imp[c]['mix'].values[m_imp]))
        cum_tot = float(np.nansum(imp[c]['total'].values[m_imp]))
        d_f = pd.to_datetime(fine_w['日期']).values
        if cy:
            m_fine = ((d_f >= pd.Timestamp(year=cy, month=1, day=1).to_datetime64()) &
                      (d_f <= pd.Timestamp(year=cy, month=cm, day=1).to_datetime64()))
        else:
            m_fine = np.zeros(len(d_f), dtype=bool)
        cum_nat_all = float(np.nansum(fine_w['全国天然'].values[m_fine]))
        cum_mix_all = float(np.nansum(fine_w['全国混合'].values[m_fine]))
        cum_tot_all = float(np.nansum(fine_w['全国总量'].values[m_fine]))
        sh_nat_c = cum_nat / cum_nat_all * 100 if cum_nat_all else None
        sh_mix_c = cum_mix / cum_mix_all * 100 if cum_mix_all else None
        sh_tot_c = cum_tot / cum_tot_all * 100 if cum_tot_all else None
        share_cn_rows.append({'国家': c, '天然占比(%)': sh_nat, '混合占比(%)': sh_mix, '总量占比(%)': sh_tot,
                              '天然累计占比(%)': sh_nat_c, '混合累计占比(%)': sh_mix_c, '总量累计占比(%)': sh_tot_c})
    # 合计行
    sum_nat = sum(r['天然占比(%)'] or 0 for r in share_cn_rows)
    sum_mix = sum(r['混合占比(%)'] or 0 for r in share_cn_rows)
    sum_tot = sum(r['总量占比(%)'] or 0 for r in share_cn_rows)
    sc_html = ''
    for r in share_cn_rows:
        sc_html += f"""<tr>
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td style="font-weight:600;">{_fmt(r['天然占比(%)'])}</td>
            <td style="font-weight:600;">{_fmt(r['混合占比(%)'])}</td>
            <td style="font-weight:600;">{_fmt(r['总量占比(%)'])}</td>
            <td>{_fmt(r['天然累计占比(%)'])}</td>
            <td>{_fmt(r['混合累计占比(%)'])}</td>
            <td>{_fmt(r['总量累计占比(%)'])}</td>
        </tr>"""
    sc_html += f"""<tr style='font-weight:700;background:#f0f7ff;'>
            <td style="text-align:left;">五国合计</td>
            <td>{_fmt(sum_nat)}</td><td>{_fmt(sum_mix)}</td><td>{_fmt(sum_tot)}</td>
            <td>-</td><td>-</td><td>-</td></tr>"""
    # 条形图
    bar2 = go.Figure()
    cats = [r['国家'] for r in share_cn_rows]
    def _nz(v):
        return v if (v is not None and not (isinstance(v, float) and pd.isna(v))) else None
    bar2.add_trace(go.Bar(x=cats, y=[_nz(r['天然占比(%)']) for r in share_cn_rows], name=f'{common_lbl} 天然占比', marker_color='#e02020',
                          hovertemplate='%{x} 天然占比: %{y:.2f}%<extra></extra>'))
    bar2.add_trace(go.Bar(x=cats, y=[_nz(r['混合占比(%)']) for r in share_cn_rows], name=f'{common_lbl} 混合占比', marker_color='#2ca6a4',
                          hovertemplate='%{x} 混合占比: %{y:.2f}%<extra></extra>'))
    bar2.add_trace(go.Bar(x=cats, y=[_nz(r['总量占比(%)']) for r in share_cn_rows], name=f'{common_lbl} 总量占比', marker_color='#f28e2b',
                          hovertemplate='%{x} 总量占比: %{y:.2f}%<extra></extra>'))
    bar2.update_layout(barmode='group', height=420, template='plotly_white',
                       font=dict(family='Microsoft YaHei, Arial', size=11, color='#444'),
                       margin=dict(l=55, r=30, t=52, b=48), yaxis_title='占比 %',
                       legend=dict(orientation='h', yanchor='bottom', y=1.03, x=0, xanchor='left',
                                   font=dict(size=11)),
                       title=dict(text='<b>五国在中国总进口中的占比</b>', x=0.01, xanchor='left',
                                  font=dict(size=15, family='Microsoft YaHei, Arial', color='#2c3e50')),
                       plot_bgcolor='#fafcff')
    n_figs += 1
    # 季节性占比格（总量）
    cn_cells = []
    for c in EXPORT_COUNTRIES:
        pv_imp = _pivot_sum(imp[c]['total'].values, df_w['日期'])
        # 全国总量 pivot（缺失月保持 NaN，不补 0）
        _tot = fine_w['全国天然'].fillna(0) + fine_w['全国混合'].fillna(0)
        _tot[fine_w['全国天然'].isna() & fine_w['全国混合'].isna()] = float('nan')
        pv_tot = _pivot_sum(_tot.values, fine_w['日期'])
        # 占比 pivot（按年）
        sh = pv_imp / pv_tot * 100
        fig = build_combo_fig(sh, '', '%', cfg, height=260, hover_label='占中国总进口', small=True, show_legend=False)
        cn_cells.append(f"""<div class='cmp-block grid-cell' data-country='{c}'>
            <div class='cell-head'>{c} 占中国总进口(总量)<span class='cell-unit'>%</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'cn_share_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.4 五国在中国总进口中的占比（天然 / 混合 / 总量）</div>
  <div class='legend-box'>全国口径来源：{fine_split_src}（天然橡胶合计 + 400280 混合胶，kg 折万吨）· 比值 = 该国从中国进口量 ÷ 全国对应口径总量</div>
  <table>
    <thead><tr><th>国家</th><th>天然占比 %</th><th>混合占比 %</th><th>总量占比 %</th>
      <th>天然累计 %</th><th>混合累计 %</th><th>总量累计 %</th></tr></thead>
    <tbody>{sc_html}</tbody>
  </table>
  <div style='height:12px;'></div>
  {bar2.to_html(include_plotlyjs='cdn', full_html=False, div_id='cn_share_bar')}
  <div style='height:12px;'></div>
  <div class='share-grid'>{''.join(cn_cells)}</div>
""")

    # ---- 8.5 进口 vs 出口 同比增速对比 ----
    def _yoy_series(values, dates):
        pv = _pivot_sum(values, dates)  # 缺失月保持 NaN，避免同比被算成 -100%
        if latest_year in pv.columns and (latest_year - 1) in pv.columns:
            s = (pv[latest_year] - pv[latest_year - 1]) / pv[latest_year - 1] * 100
        else:
            s = pd.Series(float('nan'), index=range(1, 13))
        return s.reindex(range(1, 13))

    yoy_cells = []
    for c in EXPORT_COUNTRIES:
        imp_yoy_series = _yoy_series(imp[c]['total'].values, df_w['日期'])
        exp_yoy_series = _yoy_series(exp_df[f'exp_{c}_total'].values, exp_df['日期'])
        fig = fig_two_lines(list(range(1, 13)), imp_yoy_series.values,
                            list(range(1, 13)), exp_yoy_series.values,
                            f'{c} 进口 vs 出口 同比增速', '中国进口同比', '全球出口同比', '%',
                            color_a=COUNTRY_COLORS.get(c, '#e02020'), color_b='#2ca6a4')
        yoy_cells.append(f"""<div class='cmp-block grid-cell' data-country='{c}'>
            <div class='cell-head'>{c} 进口 vs 出口 同比<span class='cell-unit'>%</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'yoy_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.5 进口 vs 出口 同比增速对比（{latest_year} vs {latest_year-1}）</div>
  <div class='legend-box'>红=中国进口总量同比，青=该国全球出口总量同比（%）。可用「国家筛选」只看某国。</div>
  <div class='share-grid'>{''.join(yoy_cells)}</div>
""")

    # ---- 8.6 真实对华出口占比（输华天然橡胶 / 全球天然橡胶出口）----
    real_rows = []
    for c in EXPORT_COUNTRIES:
        to_cn = _val_at(dest_df[f'dest_{c}_to_china'], dest_df['日期'], cy, cm) if cy else float('nan')
        exp_nat = _val_at(exp_df[f'exp_{c}_nat'], exp_df['日期'], cy, cm) if cy else float('nan')
        share = to_cn / exp_nat * 100 if (not pd.isna(to_cn) and not pd.isna(exp_nat) and exp_nat != 0) else None
        real_rows.append({'国家': c, '输华天然橡胶(万吨)': to_cn, '全球天然出口(万吨)': exp_nat,
                          '真实对华占比(%)': round(share, 1) if share is not None else None})
    real_html = ''
    for r in real_rows:
        real_html += f"""<tr>
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td>{_fmt(r['输华天然橡胶(万吨)'])}</td>
            <td>{_fmt(r['全球天然出口(万吨)'])}</td>
            <td style="font-weight:600;">{_fmt(r['真实对华占比(%)'],'%') if r['真实对华占比(%)'] is not None else '-'}</td>
        </tr>"""
    # 走势图
    real_cells = []
    for c in EXPORT_COUNTRIES:
        dfr = pd.DataFrame({'日期': dest_df['日期'].values,
                            'to_cn': dest_df[f'dest_{c}_to_china'].values})
        dfr = dfr.merge(exp_df[['日期', f'exp_{c}_nat']].rename(columns={f'exp_{c}_nat': 'exp_nat'}), on='日期', how='inner')
        dfr['share'] = dfr['to_cn'] / dfr['exp_nat'] * 100
        pv = _mkpivot(dfr['share'].values, dfr['日期'])
        fig = build_combo_fig(pv, '', '%', cfg, height=280, hover_label='真实对华占比', small=True, show_legend=False)
        real_cells.append(f"""<div class='cmp-block grid-cell' data-country='{c}'>
            <div class='cell-head'>{c} 真实对华占比(天然)<span class='cell-unit'>%</span></div>
            {fig.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'real_{c}', config={'displayModeBar': False})}
        </div>""")
        n_figs += 1
    parts.append(f"""
  <div class='section-title' style='font-size:15px;'>8.6 真实对华出口占比（输华天然橡胶 ÷ 全球天然橡胶出口）</div>
  <div class='legend-box'>数据来自出口 Excel「主要目的地（仅天然橡胶）」sheet 的中国列（天然橡胶口径）。该值为「该国天然橡胶出口中运往中国的真实份额」。</div>
  <table>
    <thead><tr><th>国家</th><th>输华天然橡胶(万吨)</th><th>全球天然出口(万吨)</th><th>真实对华占比 %</th></tr></thead>
    <tbody>{real_html}</tbody>
  </table>
  <div style='height:12px;'></div>
  <div class='share-grid'>{''.join(real_cells)}</div>
""")

    parts.append("</div>")  # 结束 八

    # ---- 报告信息 ----
    biggest = None
    if ratio_change_info:
        biggest = max(ratio_change_info, key=lambda x: abs(x[1]))
    report = {
        'export_latest_months': export_latest_months,
        'common_month': common_lbl,
        'biggest_ratio_change': (biggest[0], round(biggest[1], 2)) if biggest else None,
        'fine_split_src': fine_split_src,
    }

    # ====== 拆分为 七（界面二）/ 八（界面三）两段，并剔除已删除的 8.6 ======
    idx8 = next((i for i, p in enumerate(parts) if '八、中国进口' in p), len(parts))
    parts7 = parts[:idx8]
    parts8 = parts[idx8:]
    # 彻底移除 8.6「真实对华出口占比」相关片段（含其表格与走势图）
    parts8 = [p for p in parts8
              if '8.6 真实对华出口占比' not in p
              and '真实对华占比' not in p
              and 'real_' not in p
              and '输华天然橡胶' not in p]

    html7 = '\n'.join(parts7)
    html8 = '\n'.join(parts8)
    # JS 国家筛选（置于界面三）
    html8 += """
<script>
function cmpFilter(){
  var sel = document.getElementById('cmpCountry');
  var v = sel ? sel.value : 'ALL';
  document.querySelectorAll('.cmp-block').forEach(function(el){
    var c = el.getAttribute('data-country');
    if(v==='ALL' || c===v){ el.style.display=''; } else { el.style.display='none'; }
  });
}
</script>
"""
    n_figs7 = html7.count('Plotly.newPlot')
    n_figs8 = html8.count('Plotly.newPlot')
    return html7, html8, n_figs7, n_figs8, report
