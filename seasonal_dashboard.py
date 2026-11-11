"""
seasonal_dashboard.py —— 核心模块
=================================
职责：读取《天然橡胶进口合并报表.xlsx（按国别）》，生成：
  - dashboard_total.html（同比环比表 + 涨跌速览 + 10 国总量图 + 50 格季节性 Dashboard）
  - 同比环比变化表.xlsx

设计原则：
  - 纯函数 build(FILE, OUT_DIR) -> dict，输入文件路径 + 输出目录，返回生成产物信息
  - 数据窗口、指标映射、配色、列重命名、剔除项均集中在顶部 CONFIG 常量，方便维护
  - 任何月度数据新增/更新后再次执行即可自动反映最新月份（latest = df['日期'].max()）
"""
from __future__ import annotations
import io
import os
import sys
import shutil
import socket
import json
import urllib.request
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ====== 日志（脚本被 import 时不输出；CLI 直接运行时显示）======
logger = logging.getLogger("seasonal_dashboard")
if not logger.handlers:
    h = logging.StreamHandler(io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8'))
    h.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s',
                                    datefmt='%H:%M:%S'))
    logger.addHandler(h)
logger.setLevel(logging.INFO)


# ====== 配置常量（所有可调项集中在此）======
@dataclass
class Config:
    file_path: str = r'D:\AAA东方嘉信\备份数据\qinrex\天然橡胶进口合并报表.xlsx'
    out_dir: str = r'D:\AAA东方嘉信\备份数据\qinrex\_analysis\seasonal'
    sheet_name: str = '按国别'

    # 列重命名映射
    rename_map: Dict[str, str] = field(default_factory=lambda: {
        '400110进口量': '乳胶中国进口量',
        '400121进口量': '烟片中国进口量',
        '400122进口量': '中国标胶进口量',
        '总计进口量': '天然橡胶中国进口量',
        '进口量': '混合橡胶中国进口量',
    })
    # 剔除指标（不参与季节性 / 同比")
    exclude_metrics: set = field(default_factory=lambda: {'400129进口量', '400130进口量'})
    # 指标展示顺序
    order_new: List[str] = field(default_factory=lambda: [
        '乳胶中国进口量', '烟片中国进口量', '中国标胶进口量',
        '天然橡胶中国进口量', '混合橡胶中国进口量',
    ])
    # 分析窗口：最新年份往前 N 年（含最新共 N+1 年）
    lookback_years: int = 7
    # kg -> 万吨
    kg2wan: float = 1e7
    # 年份配色（最新年份必须 = 红色 #e02020）
    year_colors: Dict[int, str] = field(default_factory=lambda: {
        2019: '#c8c8c8', 2020: '#ddc7a8',
        2021: '#f28e2b', 2022: '#9fb8cc', 2023: '#4e79a7',
        2024: '#2ca6a4', 2025: '#7b52ab', 2026: '#e02020',
    })

    # 输出文件名
    out_html: str = 'dashboard_total.html'
    out_xlsx: str = '同比环比变化表.xlsx'

    # 局域网发布：写入本机 http.server(8899) 服务目录，手机可直接访问
    publish_dir: str = r'C:\Users\Lenovo\WorkBuddy\20260923182637'
    lan_port: int = 8899
    # lan_base 留空则每次构建时自动检测当前局域网 IP（IP 变化后 URL 依然正确）
    lan_base: str = ''

    # 组合国别分析（组内各国天胶总量加总后做季节性分析）
    groups: List[List[str]] = field(default_factory=lambda: [
        ['老挝', '缅甸', '柬埔寨'],
        ['老挝', '缅甸', '柬埔寨', '越南'],
        ['柬埔寨', '越南'],
    ])
    # 占比分析：全国总量口径来源 sheet（按细分类 = 全国合计，无国别）
    fine_sheet: str = '按细分类'
    share_top_n: int = 6  # 占比季节性图展示前 N 大来源国

    # 页面主标题
    page_title: str = '橡胶进出口分析'

    # ===== 五国出口 + 进出口对比（唯一本地基础数据源）======
    export_file_path: str = r'D:\AAA东方嘉信\备份数据\qinrex\五国橡胶出口数据.xlsx'
    export_sheet: str = '按国别'
    export_dest_sheet: str = '主要目的地（仅天然橡胶）'
    export_countries: List[str] = field(default_factory=lambda: [
        '泰国', '马来西亚', '越南', '印尼', '科特迪瓦'])


# ====== 工具函数 ======
def load_data(cfg: Config) -> pd.DataFrame:
    """读取 Excel 并做基本解析，返回宽表（含 年份 / 月份 / 天胶总量 衍生列）"""
    df = pd.read_excel(cfg.file_path, sheet_name=cfg.sheet_name, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    df['年份'] = df['日期'].dt.year
    df['月份'] = df['日期'].dt.month

    latest_year = int(df['年份'].max())
    start_year = latest_year - cfg.lookback_years
    df_w = df[df['年份'] >= start_year].copy()
    logger.info(f"加载数据: {len(df)} 行，分析窗口 {start_year}-{latest_year}")
    return df_w


def discover_country_metrics(cfg: Config, df: pd.DataFrame) -> Tuple[List[str], Dict[str, List[Tuple[str, tuple]]]]:
    """扫描列头，按「{国家}{天然橡胶|混合橡胶}」拆分并应用重命名，返回 (countries, country_metrics)"""
    import re
    country_metrics: Dict[str, List[Tuple[str, tuple]]] = {}
    for col in df.columns[1:-2]:  # 排除 日期 / 年份 / 月份
        if not isinstance(col, tuple) or len(col) != 2:
            continue
        ctype, metric = col
        if metric in cfg.exclude_metrics:
            continue
        m = re.match(r'(.+?)(天然橡胶|混合橡胶)$', ctype)
        if not m:
            continue
        country, _ = m.groups()
        new_metric = cfg.rename_map.get(metric)
        if not new_metric:
            continue
        country_metrics.setdefault(country, []).append((new_metric, col))
    for c in country_metrics:
        country_metrics[c] = sorted(country_metrics[c],
                                     key=lambda x: cfg.order_new.index(x[0]) if x[0] in cfg.order_new else 99)
    return list(country_metrics.keys()), country_metrics


def compute_total_columns(cfg: Config, df: pd.DataFrame,
                          country_metrics: Dict[str, List]) -> pd.DataFrame:
    """计算天胶总量（kg -> 万吨），加 总量_{country} 列"""
    total_cols = {}
    for country, items in country_metrics.items():
        nat = next((c for nm, c in items if nm == '天然橡胶中国进口量'), None)
        mix = next((c for nm, c in items if nm == '混合橡胶中国进口量'), None)
        s = (pd.to_numeric(df[nat], errors='coerce').fillna(0)
             + pd.to_numeric(df[mix], errors='coerce').fillna(0)) / cfg.kg2wan
        both_na = pd.to_numeric(df[nat], errors='coerce').isna() & pd.to_numeric(df[mix], errors='coerce').isna()
        s[both_na] = float('nan')
        total_cols[f'总量_{country}'] = s
    return pd.concat([df, pd.DataFrame(total_cols)], axis=1)


def load_fine_data(cfg: Config) -> Tuple[pd.DataFrame, str]:
    """读取全国天胶总量(万吨)口径。

    优先从「按细分类」sheet 的"总计进口量 + 400280 进口量"两列加总；
    若该 sheet 的"总计进口量"列全为空（数据缺失），自动 fallback 到
    从「按国别」sheet 把 10 国"天胶总计进口量 + 混合橡胶"加总作为"全国总量"近似。

    返回 (df, source_label)，df 列：日期/年份/月份/全国天胶总量(万吨)。
    """
    df = pd.read_excel(cfg.file_path, sheet_name=cfg.fine_sheet, header=[0, 1])
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df = df.dropna(subset=['日期'])
    col_nat = df.columns[6]  # ('天然橡胶合计', '总计进口量')
    col_mix = df.columns[7]  # ('天然橡胶合成橡胶合计\n400280', '进口量')
    nat = pd.to_numeric(df[col_nat], errors='coerce')
    mix = pd.to_numeric(df[col_mix], errors='coerce')
    nat_valid = nat.notna().sum()
    if nat_valid == 0:
        # === Fallback：按细分类 sheet 的"总计进口量"列全空，改用按国别 10 国加总 ===
        logger.warning(
            "「按细分类」sheet 的「总计进口量」列全空，"
            "自动 fallback 到「按国别」sheet 10 国加总作为全国总量近似")
        df_w = load_data(cfg)
        _, country_metrics = discover_country_metrics(cfg, df_w)
        df_w = compute_total_columns(cfg, df_w, country_metrics)
        total_cols = [c for c in df_w.columns
                      if isinstance(c, str) and c.startswith('总量_')]
        out = pd.DataFrame({
            '日期': df_w['日期'].values,
            '全国天胶总量': df_w[total_cols].sum(axis=1, min_count=1).values,
        })
        out['年份'] = out['日期'].dt.year
        out['月份'] = out['日期'].dt.month
        return out[['日期', '年份', '月份', '全国天胶总量']], 'fallback(10国加总)'
    both_na = nat.isna() & mix.isna()
    df['全国天胶总量'] = (nat.fillna(0) + mix.fillna(0)) / cfg.kg2wan
    df.loc[both_na, '全国天胶总量'] = float('nan')
    df['年份'] = df['日期'].dt.year
    df['月份'] = df['日期'].dt.month
    return df[['日期', '年份', '月份', '全国天胶总量']], '按细分类sheet'


def pivot_sum(df: pd.DataFrame, valcol: str) -> pd.DataFrame:
    """按(月份, 年份)求和透视。min_count=1 保证「整组全缺失」时保持 NaN，
    绝不把未更新月份补成 0（pivot_table(aggfunc='sum') 会补 0，造成折线掉到 0 线）。"""
    tmp = pd.DataFrame({'年份': df['年份'].values, '月份': df['月份'].values,
                        'v': pd.to_numeric(df[valcol], errors='coerce').values})
    return tmp.groupby(['月份', '年份'])['v'].sum(min_count=1).unstack('年份')


def build_combo_fig(pivot: pd.DataFrame, title: str, unit_str: str,
                    cfg: Config, *, height: int = 400, title_size: int = 16,
                    hover_label: str = '天胶总量', small: bool = False,
                    show_legend: bool = True) -> go.Figure:
    """统一图表函数：彩色折线 + 2026红粗 + 灰色同比柱（右轴，置底）"""
    years = sorted(pivot.columns.tolist())
    if not years:
        # pivot 全部为 NaN（极端 fallback 场景）—— 返回占位空图，避免崩溃
        fig = go.Figure()
        fig.add_annotation(text='数据为空', xref='paper', yref='paper',
                           x=0.5, y=0.5, showarrow=False,
                           font=dict(size=14, color='#999'))
        fig.update_layout(height=height, xaxis=dict(visible=False),
                          yaxis=dict(visible=False),
                          paper_bgcolor='rgba(0,0,0,0)',
                          plot_bgcolor='rgba(0,0,0,0)')
        return fig
    latest = years[-1]
    yoy_series = (pivot[latest] - pivot[latest - 1]) / pivot[latest - 1] * 100

    fig = make_subplots(specs=[[{'secondary_y': True}]])
    yy = yoy_series.dropna()
    fig.add_trace(
        go.Bar(x=yy.index, y=yy.values, name=f'{latest} 同比',
               marker_color='#c9c9c9', opacity=0.85, width=0.72,
               zorder=0,
               hovertemplate=f'<b>{latest}同比</b><br>月%{{x}}: %{{y:+.1f}}%<extra></extra>'),
        secondary_y=True)
    for year in years:
        y = pivot[year]
        is_last = (year == latest)
        fig.add_trace(
            go.Scatter(x=y.index, y=y.values, mode='lines+markers',
                       name=str(year),
                       zorder=2,
                       line=dict(color=cfg.year_colors.get(year, '#999'),
                                 width=3.4 if is_last else 1.6),
                       marker=dict(size=7 if is_last else 4,
                                   line=dict(color='white', width=1) if is_last else dict(width=0)),
                       legendrank=year,
                       hovertemplate=f'<b>{year}年 %{{x}}月</b><br>{hover_label}: %{{y:,.2f}} {unit_str}<extra></extra>'),
            secondary_y=False)
    valid_months = pivot[latest].dropna()
    last_month = int(valid_months.index.max()) if len(valid_months) else 12
    if small:
        margin = dict(l=48, r=44, t=14, b=30)
        legend_font = 9
        legend_cfg = dict(orientation='h', yanchor='bottom', y=1.0,
                          xanchor='left', x=0, title=None,
                          font=dict(size=legend_font))
    else:
        margin = dict(l=55, r=60, t=72, b=40)
        legend_font = 11
        legend_cfg = dict(orientation='h', yanchor='bottom', y=1.0,
                          xanchor='right', x=1.0, title=None,
                          font=dict(size=legend_font))
    title_cfg = None
    if title:
        title_cfg = dict(
            text=f'<b>{title}</b><span style="float:right;font-size:{10 if small else 12}px;color:#888;font-weight:400;">UNIT {unit_str}</span>',
            x=0.01, xanchor='left', y=0.97, yanchor='top',
            font=dict(size=title_size, family='Microsoft YaHei, Arial'))
    fig.update_layout(
        title=title_cfg, height=height,
        showlegend=show_legend, legend=legend_cfg,
        template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=10 if small else 11),
        margin=margin, hovermode='x unified',
    )
    fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)),
                     ticktext=[f'{m:02d}' for m in range(1, 13)])
    fig.update_yaxes(title_text=unit_str, secondary_y=False, showgrid=True,
                     gridcolor='#eee', title_font=dict(size=10))
    fig.update_yaxes(title_text='同比 %', secondary_y=True, showgrid=False,
                     zeroline=True, zerolinecolor='#bbb', zerolinewidth=1,
                     ticksuffix='%', title_font=dict(size=10))
    if not small:
        fig.add_annotation(
            text=f'注：灰色柱为 {latest} 年同比（右轴）· {latest} 年数据截至 {last_month:02d} 月',
            xref='paper', yref='paper', x=0, y=-0.22, showarrow=False,
            font=dict(size=10, color='#999'), xanchor='left')
    return fig


# ====== 主流程：build() ======
def build(cfg: Config | None = None) -> Dict[str, str]:
    """
    主入口。读取 cfg.file_path 生成 dashboard + xlsx，返回产物信息：
    {'html': '...path...', 'xlsx': '...path...', 'latest_month': 'YYYY-MM',
     'rows_in_table': N, 'figures': N}
    """
    cfg = cfg or Config()
    os.makedirs(cfg.out_dir, exist_ok=True)

    df_w = load_data(cfg)
    countries, country_metrics = discover_country_metrics(cfg, df_w)
    df_w = compute_total_columns(cfg, df_w, country_metrics)

    latest_row = df_w[df_w['日期'] == df_w['日期'].max()].iloc[0]
    prev_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(months=1))].iloc[0]
    yoy_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(years=1))].iloc[0]
    latest_lbl = latest_row['日期'].strftime('%Y-%m')
    prev_lbl = prev_row['日期'].strftime('%Y-%m')
    yoy_lbl = yoy_row['日期'].strftime('%Y-%m')
    start_year = int(df_w['年份'].min())
    latest_year = int(df_w['年份'].max())
    latest_month = int(latest_row['月份'])

    # ====== 同比环比表 ======
    cmp_rows = []
    for country in countries:
        items = country_metrics[country] + [('天胶总量(万吨)', f'总量_{country}')]
        for new_metric, col in items:
            cur = pd.to_numeric(latest_row[col], errors='coerce')
            prv = pd.to_numeric(prev_row[col], errors='coerce')
            yoy = pd.to_numeric(yoy_row[col], errors='coerce')
            is_total = (new_metric == '天胶总量(万吨)')
            div = 1  # 天胶总量在 compute_total_columns 已折万吨
            mom = (cur - prv) / prv * 100 if not pd.isna(prv) and prv != 0 else None
            yoy_pct = (cur - yoy) / yoy * 100 if not pd.isna(yoy) and yoy != 0 else None
            cmp_rows.append({
                '国家': country, '指标': new_metric,
                latest_lbl: cur / div if not pd.isna(cur) else None,
                prev_lbl: prv / div if not pd.isna(prv) else None,
                yoy_lbl: yoy / div if not pd.isna(yoy) else None,
                '环比(%)': round(mom, 1) if mom is not None else None,
                '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None,
            })
    # ====== 各国天胶总量累计(YTD)行：当年 1 月 ~ 最新月 ======
    s_by_date = df_w.set_index('日期')
    for country in countries:
        s = s_by_date[f'总量_{country}']
        cum_cur = s[s.index.year == latest_year].sum()  # 当年数据天然只到最新月
        cum_pm = s[(s.index.year == latest_year) & (s.index.month < latest_month)].sum()
        cum_ly = s[(s.index.year == latest_year - 1) & (s.index.month <= latest_month)].sum()
        mom = (cum_cur - cum_pm) / cum_pm * 100 if cum_pm else None
        yoy_pct = (cum_cur - cum_ly) / cum_ly * 100 if cum_ly else None
        cmp_rows.append({
            '国家': country, '指标': '天胶总量累计(万吨)',
            latest_lbl: cum_cur if cum_cur else None,
            prev_lbl: cum_pm if cum_pm else None,
            yoy_lbl: cum_ly if cum_ly else None,
            '环比(%)': round(mom, 1) if mom is not None else None,
            '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None,
        })
    cmp_df = pd.DataFrame(cmp_rows)

    def fmt_num(v, is_total=False):
        if v is None or pd.isna(v): return '-'
        if is_total: return f'{v:,.2f}'
        return f'{v:,.2f}' if abs(v) < 1000 else f'{v:,.0f}'

    def fmt_pct(v):
        if v is None or pd.isna(v): return '-'
        return f'{v:+.1f}%'

    def cell_color(v):
        if v is None or pd.isna(v): return ''
        if v > 0: return 'background:#fff1f0;color:#c0392b;font-weight:600;'
        if v < 0: return 'background:#f6ffed;color:#27ae60;font-weight:600;'
        return ''

    rows_html = ''
    for _, r in cmp_df.iterrows():
        is_total = r['指标'].startswith('天胶总量')
        bold = 'font-weight:700;background:#f0f7ff;' if is_total else ''
        rows_html += f'''<tr style="{bold}">
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td style="text-align:left;color:#555;">{r['指标']}</td>
            <td>{fmt_num(r[latest_lbl], is_total)}</td>
            <td>{fmt_num(r[prev_lbl], is_total)}</td>
            <td>{fmt_num(r[yoy_lbl], is_total)}</td>
            <td style="{cell_color(r['环比(%)'])}">{fmt_pct(r['环比(%)'])}</td>
            <td style="{cell_color(r['同比(%)'])}">{fmt_pct(r['同比(%)'])}</td>
        </tr>'''
    cmp_table_html = f"""
<table>
    <thead><tr>
        <th>国家</th><th>指标</th><th>{latest_lbl}（最新）</th>
        <th>{prev_lbl}</th><th>{yoy_lbl}</th><th>环比</th><th>同比</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
</table>"""

    # ====== 涨跌速览（保持单月口径，不含累计行）======
    sv = cmp_df[(cmp_df['同比(%)'].notna()) & (cmp_df['指标'] != '天胶总量累计(万吨)')].copy()
    top_up = sv.sort_values('同比(%)', ascending=False).head(5)
    top_down = sv.sort_values('同比(%)', ascending=True).head(5)

    def speed_rows(df_part, color_cls):
        h = ''
        for _, r in df_part.iterrows():
            unit = ' 万吨' if r['指标'].startswith('天胶总量') else ''
            val = f"{r[latest_lbl]:,.2f}{unit}" if not pd.isna(r[latest_lbl]) else '-'
            pct = r['同比(%)']
            cls = 'up' if pct > 0 else ('down' if pct < 0 else '')
            arrow = '▲' if pct > 0 else ('▼' if pct < 0 else '—')
            h += f'''<tr>
                <td style="text-align:left;font-weight:500;">{r['国家']}</td>
                <td style="text-align:left;color:#555;">{r['指标']}</td>
                <td>{val}</td>
                <td class="{color_cls} {cls}">{arrow} {fmt_pct(pct)}</td>
            </tr>'''
        return h

    speed_html = f"""
<div class='speed-grid'>
  <div class='speed-card'>
    <div class='speed-title up-title'>▲ 同比涨幅 TOP5</div>
    <table class='speed-table'>
      <thead><tr><th>国家</th><th>指标</th><th>{latest_lbl}</th><th>同比</th></tr></thead>
      <tbody>{speed_rows(top_up, 'pct-up')}</tbody>
    </table>
  </div>
  <div class='speed-card'>
    <div class='speed-title down-title'>▼ 同比跌幅 TOP5</div>
    <table class='speed-table'>
      <thead><tr><th>国家</th><th>指标</th><th>{latest_lbl}</th><th>同比</th></tr></thead>
      <tbody>{speed_rows(top_down, 'pct-down')}</tbody>
    </table>
  </div>
</div>"""

    # ====== 10 国总量图（月度 + 年内累计，一行两图）======
    total_pair_html = []
    for country in countries:
        pivot = pivot_sum(df_w, f'总量_{country}')
        fig_m = build_combo_fig(pivot, f'中国进口{country}天胶总量 · 月度', '万吨',
                                cfg, height=400, hover_label='天胶总量')
        fig_c = build_combo_fig(pivot.cumsum(), f'中国进口{country}天胶总量 · 年内累计',
                                '万吨', cfg, height=400, hover_label='累计天胶总量')
        total_pair_html.append(
            f"<div class='pair-row'>"
            f"<div class='pair-cell'>{fig_m.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'total_{country}')}</div>"
            f"<div class='pair-cell'>{fig_c.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'totalcum_{country}')}</div>"
            f"</div>")

    # ====== 50 格季节性 Dashboard ======
    country_blocks = []
    for country in countries:
        cells = []
        for new_metric, col in country_metrics[country]:
            df_w['_tmp_val'] = pd.to_numeric(df_w[col], errors='coerce')
            pivot = pivot_sum(df_w, '_tmp_val')
            fig = build_combo_fig(pivot, '', '千克', cfg, height=280,
                                  hover_label=new_metric, small=True, show_legend=False)
            cell = fig.to_html(include_plotlyjs='cdn', full_html=False,
                               div_id=f'cell_{country}_{new_metric}',
                               config={'displayModeBar': False})
            cells.append(
                f"<div class='grid-cell'>"
                f"<div class='cell-head'>{new_metric}<span class='cell-unit'>千克</span></div>"
                f"{cell}</div>")
        country_blocks.append(
            f"<div class='country-block'>"
            f"<div class='country-head'>{country} <span class='country-sub'>5 个指标 · 顺序：乳胶 / 烟片 / 标胶 / 天然橡胶 / 混合橡胶</span></div>"
            f"<div class='grid-wrap'>{''.join(cells)}</div>"
            f"</div>")

    # ====== 组合国别分析（组内各国天胶总量加总，月度 + 累计，一行两图）======
    group_pair_html = []
    for members in cfg.groups:
        gname = '+'.join(members)
        cols = [f'总量_{c}' for c in members]
        df_w[f'组_{gname}'] = df_w[cols].sum(axis=1, min_count=1)
        pivot = pivot_sum(df_w, f'组_{gname}')
        fig_m = build_combo_fig(pivot, f'中国进口{gname}天胶总量 · 月度', '万吨',
                                cfg, height=380, hover_label='天胶总量')
        fig_c = build_combo_fig(pivot.cumsum(), f'中国进口{gname}天胶总量 · 年内累计',
                                '万吨', cfg, height=380, hover_label='累计天胶总量')
        group_pair_html.append(
            f"<div class='pair-row'>"
            f"<div class='pair-cell'>{fig_m.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'grp_{gname}')}</div>"
            f"<div class='pair-cell'>{fig_c.to_html(include_plotlyjs='cdn', full_html=False, div_id=f'grpcum_{gname}')}</div>"
            f"</div>")

    # ====== 占比分析（分母 = 按细分类 sheet 的全国天胶总量口径）======
    df_fine, fine_source = load_fine_data(cfg)
    logger.info(f"占比分析·全国口径来源: {fine_source}")

    # ====== 新增：五国出口分析 + 中国进口/五国出口对比（界面二 / 界面三）======
    sec7_html = ''
    sec8_html = ''
    n_figs7 = n_figs8 = 0
    export_report = {}
    try:
        from export_compare import (build_export_sections, load_export_data,
                                    load_export_dest_data, load_fine_split)
        exp_df = load_export_data(cfg)
        dest_df = load_export_dest_data(cfg)
        fine_split, fine_split_src = load_fine_split(cfg)
        sec7_html, sec8_html, n_figs7, n_figs8, export_report = build_export_sections(
            cfg, dict(
                df_w=df_w, country_metrics=country_metrics,
                latest_lbl=latest_lbl, prev_lbl=prev_lbl, yoy_lbl=yoy_lbl,
                start_year=start_year, latest_year=latest_year, latest_month=latest_month,
                exp_df=exp_df, dest_df=dest_df,
                fine_split=fine_split, fine_split_src=fine_split_src,
            ))
        logger.info(f"[导出] 五国出口/对比模块生成完成，图表 七={n_figs7} 八={n_figs8} 张")
    except Exception as e:
        logger.exception(f"[导出] 五国出口/对比模块生成失败（保留原有进口模块）: {e}")
        sec7_html = ("<div class='section'><div class='section-title'>七、五国橡胶出口分析</div>"
                     f"<div class='legend-box'>⚠️ 五国出口模块生成异常，已跳过：{e}</div></div>")
        sec8_html = ("<div class='section'><div class='section-title'>八、中国进口 × 五国出口 对比</div>"
                     f"<div class='legend-box'>⚠️ 进出口对比模块生成异常，已跳过：{e}</div></div>")
        export_report = {'error': str(e)}
    # 导出模块提供的各国出口最新月 / 共同可比月（用于界面二、三副标题）
    _elm = export_report.get('export_latest_months') or {}
    latest_note = ' · '.join(f'{c} {m}' for c, m in _elm.items()) or 'N/A'
    common_lbl = export_report.get('common_month') or 'N/A'
    nat_map = df_fine.set_index('日期')['全国天胶总量']
    nat_cur = nat_map.get(latest_row['日期'])
    nat_cur = float(nat_cur) if nat_cur is not None and not pd.isna(nat_cur) else None
    nf = nat_map
    nat_cum_cur = nf[nf.index.year == latest_year].sum()
    nat_cum_ly = nf[(nf.index.year == latest_year - 1) & (nf.index.month <= latest_month)].sum()

    share_rows = []
    for country in countries:
        s = s_by_date[f'总量_{country}']
        cum_cur = s[s.index.year == latest_year].sum()
        cum_ly = s[(s.index.year == latest_year - 1) & (s.index.month <= latest_month)].sum()
        cur_val = latest_row[f'总量_{country}']
        cur_val = float(cur_val) if not pd.isna(cur_val) else None
        sh_m = cur_val / nat_cur * 100 if (cur_val is not None and nat_cur) else None
        sh_cum = cum_cur / nat_cum_cur * 100 if nat_cum_cur else None
        sh_cum_ly = cum_ly / nat_cum_ly * 100 if nat_cum_ly else None
        share_rows.append({
            '国家': country, '最新月(万吨)': cur_val,
            '单月占比(%)': round(sh_m, 2) if sh_m is not None else None,
            '年内累计(万吨)': cum_cur if cum_cur else None,
            '累计占比(%)': round(sh_cum, 2) if sh_cum is not None else None,
            '去年同期累计占比(%)': round(sh_cum_ly, 2) if sh_cum_ly is not None else None,
            '变化(pp)': round(sh_cum - sh_cum_ly, 2) if (sh_cum is not None and sh_cum_ly is not None) else None,
        })
    # 其它（10 国之外的全国剩余份额）
    ten_m = sum(r['最新月(万吨)'] or 0 for r in share_rows)
    ten_cum = sum(r['年内累计(万吨)'] or 0 for r in share_rows)
    other_m_val = (nat_cur - ten_m) if nat_cur is not None else None
    other_cum_val = (nat_cum_cur - ten_cum) if nat_cum_cur else None
    sh_other_m = other_m_val / nat_cur * 100 if (other_m_val is not None and nat_cur) else None
    sh_other_cum = other_cum_val / nat_cum_cur * 100 if (other_cum_val is not None and nat_cum_cur) else None
    share_rows.append({
        '国家': '其它(10国外)', '最新月(万吨)': other_m_val,
        '单月占比(%)': round(sh_other_m, 2) if sh_other_m is not None else None,
        '年内累计(万吨)': other_cum_val,
        '累计占比(%)': round(sh_other_cum, 2) if sh_other_cum is not None else None,
        '去年同期累计占比(%)': None, '变化(pp)': None,
    })
    sh_df = pd.DataFrame(share_rows)

    # 占比表 HTML
    def fmt_share_num(v):
        if v is None or pd.isna(v): return '-'
        return f'{v:,.2f}'

    def fmt_pp(v):
        if v is None or pd.isna(v): return '-'
        return f'{v:+.2f}pp'

    share_rows_html = ''
    for _, r in sh_df.iterrows():
        is_other = r['国家'].startswith('其它')
        bold = 'font-weight:700;background:#f0f7ff;' if not is_other else 'color:#888;'
        share_rows_html += f'''<tr style="{bold}">
            <td style="text-align:left;font-weight:500;">{r['国家']}</td>
            <td>{fmt_share_num(r['最新月(万吨)'])}</td>
            <td style="font-weight:600;">{fmt_share_num(r['单月占比(%)'])}</td>
            <td>{fmt_share_num(r['年内累计(万吨)'])}</td>
            <td style="font-weight:600;">{fmt_share_num(r['累计占比(%)'])}</td>
            <td>{fmt_share_num(r['去年同期累计占比(%)'])}</td>
            <td style="{cell_color(r['变化(pp)'])}">{fmt_pp(r['变化(pp)'])}</td>
        </tr>'''
    share_table_html = f"""
<table>
    <thead><tr>
        <th>国家</th><th>{latest_lbl}（万吨）</th><th>单月占比 %</th>
        <th>{latest_year} 年内累计（万吨）</th><th>累计占比 %</th>
        <th>{latest_year-1} 同期累计占比 %</th><th>变化 pp</th>
    </tr></thead>
    <tbody>{share_rows_html}</tbody>
</table>"""

    # 占比条形图（单月 vs 年内累计）
    bar_df = sh_df[sh_df['单月占比(%)'].notna()].sort_values('累计占比(%)')
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        y=bar_df['国家'], x=bar_df['单月占比(%)'], orientation='h',
        name=f'{latest_lbl} 单月占比', marker_color='#4e79a7',
        hovertemplate='<b>%{y}</b> 单月占比: %{x:.2f}%<extra></extra>'))
    fig_bar.add_trace(go.Bar(
        y=bar_df['国家'], x=bar_df['累计占比(%)'], orientation='h',
        name=f'{latest_lbl} 年内累计占比', marker_color='#f28e2b',
        hovertemplate='<b>%{y}</b> 累计占比: %{x:.2f}%<extra></extra>'))
    fig_bar.update_layout(
        barmode='group', height=460, template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=11),
        margin=dict(l=90, r=40, t=40, b=40),
        xaxis_title='占比 %', legend=dict(orientation='h', y=1.08, x=0),
        title=dict(text=f'<b>各国进口天胶占全国比重</b>', x=0.01, xanchor='left',
                   font=dict(size=15)))
    bar_html = fig_bar.to_html(include_plotlyjs='cdn', full_html=False, div_id='share_bar')

    # 前 N 大来源国占比季节性图（每月占比 %，同款季节性格式）
    top_countries = sh_df[~sh_df['国家'].str.startswith('其它')].sort_values(
        '累计占比(%)', ascending=False).head(cfg.share_top_n)['国家'].tolist()
    df_w['_nat'] = df_w['日期'].map(nat_map)
    share_cells = []
    for c in top_countries:
        df_w[f'占比_{c}'] = df_w[f'总量_{c}'] / df_w['_nat'] * 100
        pivot = df_w.pivot_table(index='月份', columns='年份',
                                 values=f'占比_{c}', aggfunc='mean')
        fig = build_combo_fig(pivot, '', '%', cfg, height=280,
                              hover_label='占全国比重', small=True, show_legend=False)
        cell_html = fig.to_html(include_plotlyjs='cdn', full_html=False,
                                div_id=f'share_{c}', config={'displayModeBar': False})
        share_cells.append(
            f"<div class='grid-cell'>"
            f"<div class='cell-head'>{c} 占全国进口比重<span class='cell-unit'>%</span></div>"
            f"{cell_html}</div>")

    # 共享图例色块
    years_all = sorted(df_w['年份'].unique().tolist())
    latest_all = years_all[-1]
    chips = ''
    for year in years_all:
        is_last = (year == latest_all)
        style = f"background:{cfg.year_colors.get(year, '#999')};"
        if is_last:
            style += 'height:4px;border-radius:2px;'
        chips += f"<span class='chip'><i style='{style}'></i>{year}{'（最新）' if is_last else ''}</span>"
    chips += "<span class='chip'><i class='bar-chip'></i>灰色柱 = 2026同比%（右轴）</span>"
    legend_chips_html = f"<div class='chips-row'>{chips}</div>"

    # ====== 整合 HTML（三界面导航）======
    def _lazy_wrap(html, qname):
        """将 html 内的 Plotly.newPlot 脚本延迟到所属界面首次切换时执行（按需加载）。"""
        import re as _re
        chunks = _re.split(r'(<script[^>]*>.*?</script>)', html, flags=_re.DOTALL)
        out = []
        for ch in chunks:
            if ch.startswith('<script') and 'Plotly.newPlot' in ch:
                m = _re.match(r'<script[^>]*>(.*)</script>', ch, flags=_re.DOTALL)
                inner = m.group(1)
                out.append("<script>%s.push(function(){{ %s }});</script>" % (qname, inner))
            else:
                out.append(ch)
        return ''.join(out)

    nav_html = """
<div class='tab-nav'>
  <button data-tab='tab1' class='active' onclick="showTab('tab1')">中国月度进口数据</button>
  <button data-tab='tab2' onclick="showTab('tab2')">五国月度出口数据</button>
  <button data-tab='tab3' onclick="showTab('tab3')">进出口数据对比</button>
</div>"""

    sec7_lazy = _lazy_wrap(sec7_html, 'TAB2_Q')
    sec8_lazy = _lazy_wrap(sec8_html, 'TAB3_Q')

    combined_html = f"""<!DOCTYPE html>
<html lang='zh-CN'>
<head>
<meta charset='utf-8'>
<title>{cfg.page_title}</title>
<style>
  body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; padding: 24px;
          background: linear-gradient(135deg, #f0f4f8 0%, #e8eef3 100%); margin:0; }}
  .container {{ max-width: 1750px; margin: 0 auto; }}
  h1 {{ color: #2c3e50; border-bottom: 3px solid #1890ff; padding-bottom: 12px;
        margin-bottom: 8px; font-size: 26px; }}
  .subtitle {{ color: #7f8c8d; font-size: 14px; margin-bottom: 24px; }}
  .section {{ background: white; padding: 24px; border-radius: 12px;
              box-shadow: 0 4px 16px rgba(0,0,0,0.06); margin-bottom: 32px; }}
  .section-title {{ font-size: 18px; font-weight: 700; color: #2c3e50;
                    margin: 0 0 16px 0; padding-left: 12px; border-left: 5px solid #1890ff; }}
  .legend-box {{ background: #fffbe6; padding: 14px 18px; border-left: 4px solid #faad14;
                 border-radius: 6px; margin-bottom: 20px; font-size: 13px; color: #5c4d00; }}
  .legend-box b {{ color: #2c3e50; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th {{ background: #1890ff; color: white; padding: 12px 8px; text-align: center; font-weight: 600; }}
  td {{ padding: 9px 8px; border-bottom: 1px solid #ecf0f1; text-align: center; color: #2c3e50; }}
  tr:nth-child(even) {{ background: #fafbfc; }}
  tr:hover {{ background: #e6f7ff !important; }}
  .formula {{ background:#e6f7ff; border-left:4px solid #1890ff; padding:10px 16px;
              border-radius:6px; margin-bottom:16px; font-size:14px; }}
  .speed-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  .speed-card {{ border: 1px solid #e8ecef; border-radius: 10px; overflow: hidden; }}
  .speed-title {{ padding: 10px 16px; font-weight: 700; font-size: 14px; color: white; }}
  .up-title {{ background: #c0392b; }}
  .down-title {{ background: #27ae60; }}
  .speed-table th {{ background: #f5f7fa; color: #2c3e50; font-size: 12px; padding: 8px; }}
  .speed-table td {{ font-size: 13px; }}
  .pct-up {{ color: #c0392b; font-weight: 700; }}
  .pct-down {{ color: #27ae60; font-weight: 700; }}
  .grid-wrap {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }}
  .pair-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }}
  .pair-cell {{ border: 1px solid #eef1f4; border-radius: 8px; background: #fff; padding: 4px; }}
  .share-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }}
  .grid-cell {{ border: 1px solid #eef1f4; border-radius: 8px; overflow: hidden; background: #fff; }}
  .country-block {{ margin-bottom: 26px; }}
  .country-head {{ font-size: 15px; font-weight: 700; color: #2c3e50; padding: 8px 12px;
                   background: linear-gradient(90deg, #eaf4ff 0%, #f7fbff 100%);
                   border-left: 4px solid #1890ff; border-radius: 6px; margin-bottom: 10px; }}
  .country-sub {{ font-size: 11px; font-weight: 400; color: #999; margin-left: 8px; }}
  .cell-head {{ font-size: 12.5px; font-weight: 700; color: #2c3e50; padding: 7px 10px 2px;
                background: #f8fafc; border-bottom: 1px solid #eef1f4;
                display: flex; justify-content: space-between; align-items: center; }}
  .cell-unit {{ font-size: 10px; font-weight: 400; color: #999; }}
  .chips-row {{ display: flex; flex-wrap: wrap; gap: 14px; align-items: center;
                padding: 10px 14px; background: #f8fafc; border: 1px solid #eef1f4;
                border-radius: 8px; margin-bottom: 14px; }}
  .chip {{ display: inline-flex; align-items: center; gap: 5px;
           font-size: 12px; color: #555; }}
  .chip i {{ display: inline-block; width: 18px; height: 3px; border-radius: 2px; }}
  .chip .bar-chip {{ width: 10px; height: 12px; background: #c9c9c9; opacity: 0.85;
                     border-radius: 2px; }}
  .tab-nav {{ display:flex; gap:10px; margin:0 0 18px; flex-wrap:wrap; }}
  .tab-nav button {{ padding:10px 20px; font-size:14px; font-weight:600; cursor:pointer;
                     border:1px solid #1890ff; background:#fff; color:#1890ff; border-radius:8px;
                     transition:all .15s; }}
  .tab-nav button:hover {{ background:#e6f7ff; }}
  .tab-nav button.active {{ background:#1890ff; color:#fff; }}
  .tab-pane {{ display:none; }}
  .tab-pane.active {{ display:block; }}
</style>
</head>
<body>
<div class='container'>

<h1>{cfg.page_title}</h1>

{nav_html}

<div class='tab-pane active' id='tab1'>
<div class='subtitle'>
  数据源：天然橡胶进口合并报表.xlsx（按国别） · 窗口 {start_year}-{latest_year} · 天胶总量单位：万吨（原始 kg 折算）<br>
  生成时间：<span id='gen-ts'></span> · 进口数据最新月：<b>{latest_lbl}</b>
</div>
<div class='section'>
  <div class='section-title'>一、同比 / 环比 变化表（最新月 {latest_lbl}）</div>
  <div class='formula'>
    <b>天胶总量 = 天然橡胶中国进口量 + 混合橡胶中国进口量</b>（按月求和，折算万吨）
  </div>
  <div class='legend-box'>
    <b>对比口径：</b>{latest_lbl}（最新） · {prev_lbl}（环比） · {yoy_lbl}（同比）<br>
    <b>累计行口径：</b>当年 1 月 ~ {latest_lbl} 累计 · 环比 = 较上月末累计（即当月增量贡献） · 同比 = 较 {yoy_lbl[:4]} 年同期（1~{latest_month:02d} 月）累计<br>
    <b>配色：</b><span style='color:#c0392b;'>上涨浅红</span> · <span style='color:#27ae60;'>下跌浅绿</span> · "-"表示零/缺失<br>
    <b>指标：</b>乳胶=HS400110 · 烟片=HS400121 · 标胶=HS400122 · 混合橡胶=HS400280 · <b>底色行 = 天胶总量（单月 / 累计）</b>
  </div>
  {cmp_table_html}
</div>

<div class='section'>
  <div class='section-title'>二、同比涨跌速览（{latest_lbl} vs {yoy_lbl}）</div>
  {speed_html}
</div>

<div class='section'>
  <div class='section-title'>三、各国进口天胶总量 · 月度 + 年内累计（每国一行两图）</div>
  <div class='legend-box'>
    <b>左图：</b>月度季节性走势 + 当月同比柱 · <b>右图：</b>当年 1 月起逐月累加的累计曲线（每年一条线）+ 累计同比柱（右轴）<br>
    累计曲线反映年内到货节奏，同比柱 = 截至该月的累计量较去年同期变化
  </div>
  {''.join(total_pair_html)}
</div>

<div class='section'>
  <div class='section-title'>四、组合国别分析（组内各国天胶总量加总）</div>
  <div class='legend-box'>
    <b>组合口径：</b>{' · '.join('+'.join(g) for g in cfg.groups)} · 每组一行两图（月度 / 年内累计）
  </div>
  {''.join(group_pair_html)}
</div>

<div class='section'>
  <div class='section-title'>五、中国进口总量占比分析（分母 = 按细分类 sheet 全国天胶总量）</div>
  <div class='legend-box'>
    <b>全国口径：</b>天然橡胶总计进口量 + 400280 混合胶（来自「按细分类」sheet，kg 折万吨）<br>
    <b>占比 = </b>该国天胶总量 ÷ 全国天胶总量 · "其它(10国外)" = 全国 - 10 国合计 · pp = 百分点
  </div>
  {share_table_html}
  <div style='height:20px;'></div>
  {bar_html}
  <div style='height:12px;'></div>
  <div class='chips-row'>
    <span class='chip'>前 {cfg.share_top_n} 大来源国 · 月度占比季节性（{start_year}-{latest_year}，{latest_year} 红色加粗，灰柱 = 占比同比变化）</span>
  </div>
  <div class='share-grid'>{''.join(share_cells)}</div>
</div>

<div class='section'>
  <div class='section-title'>六、季节性 Dashboard（10 国 × 5 指标，按国家分块 · 每行 3 图）</div>
  <div class='legend-box'>
    <b>格式说明：</b>每格 = 某国某指标的月度季节性走势（{start_year}-{latest_year} 共 8 年，{latest_year} 红色加粗）<br>
    已剔除 400129 / 400130 · 灰色同比柱置底不遮挡折线 · 鼠标悬停查看数值
  </div>
  {legend_chips_html}
  {''.join(country_blocks)}
</div>
</div><!-- /tab1 -->

<div class='tab-pane' id='tab2'>
<div class='subtitle'>
  数据源：五国橡胶出口数据.xlsx（按国别） · 窗口 {start_year}-{latest_year} · 单位：万吨（原始 kg 折算）<br>
  五国出口最新月：<b>{latest_note}</b> · 共同可比月：{common_lbl}
</div>
{sec7_lazy}
</div>

<div class='tab-pane' id='tab3'>
<div class='subtitle'>
  数据源：中国进口（天然橡胶进口合并报表.xlsx）+ 五国出口（五国橡胶出口数据.xlsx） · 单位：万吨<br>
  进口最新月：{latest_lbl} · 五国出口最新月：{latest_note} · 共同可比月：{common_lbl}
</div>
{sec8_lazy}
</div>

</div><!-- /container -->
"""

    # 去除重复的 Plotly CDN <script>，并把懒加载定义脚本一并放到 <body> 顶部
    # （必须早于 tab2/tab3 内 TAB2_Q.push / TAB3_Q.push 的执行，否则队列定义晚于使用会报 undefined）
    JS_DEF = """
<script>
  var TAB2_Q = [];
  var TAB3_Q = [];
  function runLazy(q){ if(q && q.length){ q.forEach(function(f){ try{ f(); }catch(e){ console.error(e); } }); q.length = 0; } }
  function showTab(id){
    var panes = document.querySelectorAll('.tab-pane');
    for(var i=0;i<panes.length;i++){ panes[i].classList.remove('active'); }
    var pane = document.getElementById(id);
    if(pane){ pane.classList.add('active'); }
    var btns = document.querySelectorAll('.tab-nav button');
    for(var j=0;j<btns.length;j++){ btns[j].classList.toggle('active', btns[j].getAttribute('data-tab') === id); }
    if(id === 'tab2'){ runLazy(TAB2_Q); }
    if(id === 'tab3'){ runLazy(TAB3_Q); }
    setTimeout(function(){
      if(pane){ var plots = pane.querySelectorAll('.js-plotly-plot');
        for(var k=0;k<plots.length;k++){ try{ Plotly.Plots.resize(plots[k]); }catch(e){} } }
    }, 60);
  }
</script>
"""
    import re as _re2
    _cdn = _re2.compile(r"<script[^>]*src=['\"]https://cdn\.plot\.ly/[^'\"]*['\"][^>]*></script>")
    _cdn_list = _cdn.findall(combined_html)
    if _cdn_list:
        _first = _cdn_list[0]
        combined_html = combined_html.replace(_first, '')
        _pos = combined_html.find('<body>')
        combined_html = combined_html[:_pos + len('<body>')] + '\n' + _first + '\n' + JS_DEF + combined_html[_pos + len('<body>'):]

    combined_html += """
<script>document.getElementById('gen-ts').textContent = new Date().toLocaleString('zh-CN');</script>
</body>
</html>
"""

    html_path = os.path.join(cfg.out_dir, cfg.out_html)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(combined_html)
    xlsx_path = os.path.join(cfg.out_dir, cfg.out_xlsx)
    cmp_df.to_excel(xlsx_path, index=False)

    # 发布到局域网服务器目录（若目录存在）
    url = None
    url_kind = None  # 'ngrok' / 'lan' / 'custom' / 'localhost'
    if cfg.publish_dir and os.path.isdir(cfg.publish_dir):
        pub_path = os.path.join(cfg.publish_dir, cfg.out_html)
        shutil.copyfile(html_path, pub_path)
        base = None
        # 1) 优先 ngrok 公网 URL（如本机有 ngrok 进程，会暴露 127.0.0.1:4040 API）
        if not base:
            try:
                r = urllib.request.urlopen('http://127.0.0.1:4040/api/tunnels', timeout=2)
                data = json.loads(r.read().decode('utf-8'))
                for t in data.get('tunnels', []):
                    if t.get('proto') == 'https' and t.get('public_url'):
                        base = t['public_url']
                        url_kind = 'ngrok'
                        break
            except Exception:
                pass
        # 2) 用户自定义 lan_base
        if not base and cfg.lan_base:
            base = cfg.lan_base
            url_kind = 'custom'
        # 3) 自动检测当前局域网 IP（UDP connect 不真发包）
        if not base:
            try:
                _s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                _s.connect(('8.8.8.8', 80))
                _ip = _s.getsockname()[0]
                _s.close()
                base = f'http://{_ip}:{cfg.lan_port}'
                url_kind = 'lan'
            except Exception:
                base = f'http://127.0.0.1:{cfg.lan_port}'
                url_kind = 'localhost'
        url = f'{base}/{cfg.out_html}'
        logger.info(f"发布 URL（{url_kind}）: {url}")

    n_figs = (len(total_pair_html) * 2 + len(group_pair_html) * 2
              + sum(len(v) for v in country_metrics.values())
              + 1 + len(share_cells) + n_figs7 + n_figs8)
    logger.info(f"[OK] HTML  -> {html_path}")
    logger.info(f"[OK] XLSX -> {xlsx_path}")
    logger.info(f"表格 {len(cmp_df)} 行，最新月 {latest_lbl}，图表 {n_figs} 张")
    if url:
        logger.info(f"[OK] 局域网 -> {url}")

    return {
        'html': html_path, 'xlsx': xlsx_path,
        'latest_month': latest_lbl, 'rows_in_table': len(cmp_df),
        'figures': n_figs,
        'countries': countries,
        'url': url,
        'export': export_report,
    }


if __name__ == '__main__':
    build()