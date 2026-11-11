"""
v4 最终版：
1. 修复天胶总量表格显示 0.00 的 bug（重复除以 1e7）
2. 图表格式统一为参考图样式：彩色折线 + 2026 红粗 + 灰色同比柱（右轴）
3. 末尾补回 10行 x 5列 = 50 格季节性 Dashboard（5 指标，已剔除 400129/400130）
4. 新增同比涨跌速览（TOP5 涨 / TOP5 跌）
输出：dashboard_total.html
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import re
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

FILE = r'D:\AAA东方嘉信\备份数据\qinrex\天然橡胶进口合并报表.xlsx'
OUT_DIR = r'D:\AAA东方嘉信\备份数据\qinrex\_analysis\seasonal'

df = pd.read_excel(FILE, sheet_name='按国别', header=[0, 1])
df = df.iloc[1:].reset_index(drop=True)
df.columns = ['日期'] + list(df.columns[1:])
df['日期'] = pd.to_datetime(df['日期'])
df['年份'] = df['日期'].dt.year
df['月份'] = df['日期'].dt.month

LATEST_YEAR = int(df['年份'].max())
START_YEAR = LATEST_YEAR - 7
df_w = df[df['年份'] >= START_YEAR].copy()

KG2WAN = 1e7  # kg -> 万吨

RENAME_MAP = {
    '400110进口量': '乳胶中国进口量',
    '400121进口量': '烟片中国进口量',
    '400122进口量': '中国标胶进口量',
    '总计进口量': '天然橡胶中国进口量',
    '进口量': '混合橡胶中国进口量',
}
EXCLUDE = {'400129进口量', '400130进口量'}
ORDER_NEW = ['乳胶中国进口量', '烟片中国进口量', '中国标胶进口量', '天然橡胶中国进口量', '混合橡胶中国进口量']

# 国家 -> 指标列
country_metrics = {}
for col in df_w.columns[1:-2]:
    ctype, metric = col
    if metric in EXCLUDE:
        continue
    m = re.match(r'(.+?)(天然橡胶|混合橡胶)$', ctype)
    if not m:
        continue
    country, _ = m.groups()
    new_metric = RENAME_MAP.get(metric)
    if not new_metric:
        continue
    country_metrics.setdefault(country, []).append((new_metric, col))

for country in country_metrics:
    country_metrics[country] = sorted(country_metrics[country],
                                      key=lambda x: ORDER_NEW.index(x[0]) if x[0] in ORDER_NEW else 99)

countries = list(country_metrics.keys())

# ===== 计算天胶总量（kg -> 万吨）=====
total_cols = {}
for country in countries:
    nat = None; mix = None
    for new_metric, col in country_metrics[country]:
        if new_metric == '天然橡胶中国进口量':
            nat = col
        elif new_metric == '混合橡胶中国进口量':
            mix = col
    s = (pd.to_numeric(df_w[nat], errors='coerce').fillna(0)
         + pd.to_numeric(df_w[mix], errors='coerce').fillna(0)) / KG2WAN
    both_na = pd.to_numeric(df_w[nat], errors='coerce').isna() & pd.to_numeric(df_w[mix], errors='coerce').isna()
    s[both_na] = float('nan')
    total_cols[f'总量_{country}'] = s

df_w = pd.concat([df_w, pd.DataFrame(total_cols)], axis=1)
print('天胶总量列已计算（万吨）')

# ===== 同比环比表（5指标 + 天胶总量）=====
latest_row = df_w[df_w['日期'] == df_w['日期'].max()].iloc[0]
prev_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(months=1))].iloc[0]
yoy_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(years=1))].iloc[0]
LATEST_LBL = latest_row['日期'].strftime('%Y-%m')
PREV_LBL = prev_row['日期'].strftime('%Y-%m')
YOY_LBL = yoy_row['日期'].strftime('%Y-%m')

cmp_rows = []
for country in countries:
    items = country_metrics[country] + [('天胶总量(万吨)', f'总量_{country}')]
    for new_metric, col in items:
        cur = pd.to_numeric(latest_row[col], errors='coerce')
        prv = pd.to_numeric(prev_row[col], errors='coerce')
        yoy = pd.to_numeric(yoy_row[col], errors='coerce')
        is_total = (new_metric == '天胶总量(万吨)')
        # FIX: 天胶总量在计算阶段已折算为万吨，此处不再除以 KG2WAN
        div = 1 if is_total else 1
        mom = (cur - prv) / prv * 100 if not pd.isna(prv) and prv != 0 else None
        yoy_pct = (cur - yoy) / yoy * 100 if not pd.isna(yoy) and yoy != 0 else None
        cmp_rows.append({
            '国家': country,
            '指标': new_metric,
            LATEST_LBL: cur / div if not pd.isna(cur) else None,
            PREV_LBL: prv / div if not pd.isna(prv) else None,
            YOY_LBL: yoy / div if not pd.isna(yoy) else None,
            '环比(%)': round(mom, 1) if mom is not None else None,
            '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None,
        })
cmp_df = pd.DataFrame(cmp_rows)

def fmt_num(v, is_total=False):
    if v is None or pd.isna(v):
        return '-'
    if is_total:
        return f'{v:,.2f}'
    return f'{v:,.2f}' if abs(v) < 1000 else f'{v:,.0f}'

def fmt_pct(v):
    if v is None or pd.isna(v):
        return '-'
    return f'{v:+.1f}%'

def cell_color(v):
    if v is None or pd.isna(v):
        return ''
    if v > 0:
        return 'background:#fff1f0;color:#c0392b;font-weight:600;'
    elif v < 0:
        return 'background:#f6ffed;color:#27ae60;font-weight:600;'
    return ''

rows_html = ''
for _, r in cmp_df.iterrows():
    is_total = r['指标'].startswith('天胶总量')
    bold = 'font-weight:700;background:#f0f7ff;' if is_total else ''
    rows_html += f'''<tr style="{bold}">
        <td style="text-align:left;font-weight:500;">{r['国家']}</td>
        <td style="text-align:left;color:#555;">{r['指标']}</td>
        <td>{fmt_num(r[LATEST_LBL], is_total)}</td>
        <td>{fmt_num(r[PREV_LBL], is_total)}</td>
        <td>{fmt_num(r[YOY_LBL], is_total)}</td>
        <td style="{cell_color(r['环比(%)'])}">{fmt_pct(r['环比(%)'])}</td>
        <td style="{cell_color(r['同比(%)'])}">{fmt_pct(r['同比(%)'])}</td>
    </tr>'''

cmp_table_html = f"""
<table>
    <thead><tr>
        <th>国家</th><th>指标</th><th>{LATEST_LBL}（最新）</th>
        <th>{PREV_LBL}</th><th>{YOY_LBL}</th><th>环比</th><th>同比</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
</table>"""

# ===== 同比涨跌速览 =====
sv = cmp_df[cmp_df['同比(%)'].notna()].copy()
top_up = sv.sort_values('同比(%)', ascending=False).head(5)
top_down = sv.sort_values('同比(%)', ascending=True).head(5)

def speed_rows(df_part, color_cls):
    h = ''
    for _, r in df_part.iterrows():
        unit = ' 万吨' if r['指标'].startswith('天胶总量') else ''
        val = f"{r[LATEST_LBL]:,.2f}{unit}" if not pd.isna(r[LATEST_LBL]) else '-'
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
      <thead><tr><th>国家</th><th>指标</th><th>{LATEST_LBL}</th><th>同比</th></tr></thead>
      <tbody>{speed_rows(top_up, 'pct-up')}</tbody>
    </table>
  </div>
  <div class='speed-card'>
    <div class='speed-title down-title'>▼ 同比跌幅 TOP5</div>
    <table class='speed-table'>
      <thead><tr><th>国家</th><th>指标</th><th>{LATEST_LBL}</th><th>同比</th></tr></thead>
      <tbody>{speed_rows(top_down, 'pct-down')}</tbody>
    </table>
  </div>
</div>
"""

# ===== 图表工具（参考图样式）=====
YEARS = sorted(df_w['年份'].unique().tolist())
LATEST = YEARS[-1]
YEAR_COLORS = {
    2019: '#c8c8c8', 2020: '#ddc7a8',
    2021: '#f28e2b', 2022: '#9fb8cc', 2023: '#4e79a7',
    2024: '#2ca6a4', 2025: '#7b52ab', 2026: '#e02020',
}

def build_combo_fig(pivot, title, unit_str, height=400, title_size=16,
                    hover_label='天胶总量', small=False, show_legend=True):
    """彩色折线 + 2026红粗 + 灰色同比柱（右轴）"""
    yoy_series = (pivot[LATEST] - pivot[LATEST - 1]) / pivot[LATEST - 1] * 100
    fig = make_subplots(specs=[[{'secondary_y': True}]])
    yy = yoy_series.dropna()
    fig.add_trace(
        go.Bar(x=yy.index, y=yy.values, name=f'{LATEST} 同比',
               marker_color='#c9c9c9', opacity=0.85, width=0.72,
               zorder=0,
               hovertemplate=f'<b>{LATEST}同比</b><br>月%{{x}}: %{{y:+.1f}}%<extra></extra>'),
        secondary_y=True)
    for year in YEARS:
        y = pivot[year]
        is_last = (year == LATEST)
        fig.add_trace(
            go.Scatter(x=y.index, y=y.values, mode='lines+markers',
                       name=str(year),
                       zorder=2,
                       line=dict(color=YEAR_COLORS.get(year, '#999'),
                                 width=3.4 if is_last else 1.6),
                       marker=dict(size=7 if is_last else 4,
                                   line=dict(color='white', width=1) if is_last else dict(width=0)),
                       legendrank=year,
                       hovertemplate=f'<b>{year}年 %{{x}}月</b><br>{hover_label}: %{{y:,.2f}} {unit_str}<extra></extra>'),
            secondary_y=False)
    valid_months = pivot[LATEST].dropna()
    last_month = int(valid_months.index.max()) if len(valid_months) else 12
    if small:
        margin = dict(l=48, r=44, t=14, b=30)
        legend_font = 9
        title_font = title_size
        legend_cfg = dict(orientation='h', yanchor='bottom', y=1.0,
                          xanchor='left', x=0, title=None,
                          font=dict(size=legend_font))
    else:
        margin = dict(l=55, r=60, t=72, b=40)
        legend_font = 11
        title_font = title_size
        legend_cfg = dict(orientation='h', yanchor='bottom', y=1.0,
                          xanchor='right', x=1.0, title=None,
                          font=dict(size=legend_font))
    title_cfg = None
    if title:
        title_cfg = dict(
            text=f'<b>{title}</b><span style="float:right;font-size:{10 if small else 12}px;color:#888;font-weight:400;">UNIT {unit_str}</span>',
            x=0.01, xanchor='left', y=0.97, yanchor='top',
            font=dict(size=title_font, family='Microsoft YaHei, Arial'))
    fig.update_layout(
        title=title_cfg,
        height=height,
        showlegend=show_legend,
        legend=legend_cfg,
        template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=10 if small else 11),
        margin=margin,
        hovermode='x unified',
    )
    fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)),
                     ticktext=[f'{m:02d}' for m in range(1, 13)])
    fig.update_yaxes(title_text=unit_str, secondary_y=False, showgrid=True, gridcolor='#eee',
                     title_font=dict(size=10))
    fig.update_yaxes(title_text='同比 %', secondary_y=True, showgrid=False,
                     zeroline=True, zerolinecolor='#bbb', zerolinewidth=1, ticksuffix='%',
                     title_font=dict(size=10))
    if not small:
        fig.add_annotation(
            text=f'注：灰色柱为 {LATEST} 年同比（右轴）· {LATEST} 年数据截至 {last_month:02d} 月',
            xref='paper', yref='paper', x=0, y=-0.22, showarrow=False,
            font=dict(size=10, color='#999'), xanchor='left')
    return fig

# ===== 10 国天胶总量图 =====
total_figs_html = []
for country in countries:
    pivot = df_w.pivot_table(index='月份', columns='年份', values=f'总量_{country}', aggfunc='sum')
    fig = build_combo_fig(pivot, f'中国进口{country}天胶总量', '万吨', height=400,
                          hover_label='天胶总量')
    total_figs_html.append(fig.to_html(include_plotlyjs='cdn', full_html=False,
                                       div_id=f'total_{country}'))
    print(f'[OK] {country} 天胶总量图')

# ===== 50 格季节性 Dashboard（每行 3 图，按国家分块）=====
country_blocks = []
for country in countries:
    cells = []
    for new_metric, col in country_metrics[country]:
        df_w['_tmp_val'] = pd.to_numeric(df_w[col], errors='coerce')
        pivot = df_w.pivot_table(index='月份', columns='年份', values='_tmp_val', aggfunc='sum')
        fig = build_combo_fig(pivot, '', '千克', height=280,
                              hover_label=new_metric, small=True, show_legend=False)
        cell = fig.to_html(include_plotlyjs='cdn', full_html=False,
                           div_id=f'cell_{country}_{new_metric}',
                           config={'displayModeBar': False})
        cells.append(
            f"<div class='grid-cell'>"
            f"<div class='cell-head'>{new_metric}<span class='cell-unit'>千克</span></div>"
            f"{cell}</div>")
        print(f'[OK] {country} {new_metric} 季节性格子')
    country_blocks.append(
        f"<div class='country-block'>"
        f"<div class='country-head'>{country} <span class='country-sub'>5 个指标 · 顺序：乳胶 / 烟片 / 标胶 / 天然橡胶 / 混合橡胶</span></div>"
        f"<div class='grid-wrap'>{''.join(cells)}</div>"
        f"</div>")
    print(f'[OK] {country} 分块完成')

grid_html = ''.join(country_blocks)

# 共享图例（年份色块 + 灰色同比柱）
chips = ''
for year in YEARS:
    is_last = (year == LATEST)
    style = f"background:{YEAR_COLORS.get(year, '#999')};"
    if is_last:
        style += 'height:4px;border-radius:2px;'
    chips += f"<span class='chip'><i style='{style}'></i>{year}{'（最新）' if is_last else ''}</span>"
chips += "<span class='chip'><i class='bar-chip'></i>灰色柱 = 2026同比%（右轴）</span>"
legend_chips_html = f"<div class='chips-row'>{chips}</div>"

# ===== 整合 HTML =====
combined_html = f"""<!DOCTYPE html>
<html lang='zh-CN'>
<head>
<meta charset='utf-8'>
<title>主产国进口天胶分析</title>
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
</style>
</head>
<body>
<div class='container'>

<h1>中国月度进口（分国别）</h1>
<div class='subtitle'>
  数据源：天然橡胶进口合并报表.xlsx（按国别） · 窗口 {START_YEAR}-{LATEST_YEAR} · 天胶总量单位：万吨（原始 kg 折算）
</div>

<div class='section'>
  <div class='section-title'>一、同比 / 环比 变化表（最新月 {LATEST_LBL}）</div>
  <div class='formula'>
    <b>天胶总量 = 天然橡胶中国进口量 + 混合橡胶中国进口量</b>（按月求和，折算万吨）
  </div>
  <div class='legend-box'>
    <b>对比口径：</b>{LATEST_LBL}（最新） · {PREV_LBL}（环比） · {YOY_LBL}（同比）<br>
    <b>配色：</b><span style='color:#c0392b;'>上涨浅红</span> · <span style='color:#27ae60;'>下跌浅绿</span> · "-"表示零/缺失<br>
    <b>指标：</b>乳胶=HS400110 · 烟片=HS400121 · 标胶=HS400122 · 混合橡胶=HS400280 · <b>底色行 = 天胶总量</b>
  </div>
  {cmp_table_html}
</div>

<div class='section'>
  <div class='section-title'>二、同比涨跌速览（{LATEST_LBL} vs {YOY_LBL}）</div>
  {speed_html}
</div>

<div class='section'>
  <div class='section-title'>三、各国进口天胶总量 · 季节性 + 同比（折线 + 同比柱）</div>
  {''.join(total_figs_html)}
</div>

<div class='section'>
  <div class='section-title'>四、季节性 Dashboard（10 国 × 5 指标，按国家分块 · 每行 3 图）</div>
  <div class='legend-box'>
    <b>格式说明：</b>每格 = 某国某指标的月度季节性走势（{START_YEAR}-{LATEST} 共 8 年，{LATEST} 红色加粗）<br>
    已剔除 400129 / 400130 · 灰色同比柱置底不遮挡折线 · 鼠标悬停查看数值
  </div>
  {legend_chips_html}
  {grid_html}
</div>

</div>
</body>
</html>
"""

out_path = os.path.join(OUT_DIR, 'dashboard_total.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(combined_html)
print(f'\n[OK] 输出: {out_path}')
print(f'文件大小: {os.path.getsize(out_path)/1024/1024:.2f} MB')

xlsx_path = os.path.join(OUT_DIR, '同比环比变化表.xlsx')
cmp_df.to_excel(xlsx_path, index=False)
print(f'[OK] xlsx: {xlsx_path}')

# 打印天胶总量速览验证
print('\n=== 天胶总量验证（表格数值应为万吨）===')
for country in countries:
    v = latest_row[f'总量_{country}']
    print(f'{country}: {v:.2f} 万吨')
