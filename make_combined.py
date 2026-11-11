"""
整合版：同比环比变化表 + 季节性图（50格）合并到一个 HTML
列重命名：
    400110进口量 → 乳胶中国进口量
    400121进口量 → 烟片中国进口量
    400122进口量 → 中国标胶进口量
    400129 / 400130 → 不做分析（排除）
    总计进口量  → 天然橡胶中国进口量
    进口量（混合胶）→ 混合橡胶中国进口量
季节性窗口：2019-2026，最新2026红色加粗
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
os.makedirs(OUT_DIR, exist_ok=True)

# 用户定义的指标重命名（仅保留 5 个）
RENAME_MAP = {
    '400110进口量': '乳胶中国进口量',
    '400121进口量': '烟片中国进口量',
    '400122进口量': '中国标胶进口量',
    '总计进口量': '天然橡胶中国进口量',
    '进口量': '混合橡胶中国进口量',
}
# 排除
EXCLUDE = {'400129进口量', '400130进口量'}
# 排序顺序
ORDER_NEW = ['乳胶中国进口量', '烟片中国进口量', '中国标胶进口量', '天然橡胶中国进口量', '混合橡胶中国进口量']

df = pd.read_excel(FILE, sheet_name='按国别', header=[0, 1])
df = df.iloc[1:].reset_index(drop=True)
df.columns = ['日期'] + list(df.columns[1:])
df['日期'] = pd.to_datetime(df['日期'])
df['年份'] = df['日期'].dt.year
df['月份'] = df['日期'].dt.month

LATEST_YEAR = int(df['年份'].max())
START_YEAR = LATEST_YEAR - 7
df_w = df[df['年份'] >= START_YEAR].copy()
print(f'窗口: {START_YEAR} - {LATEST_YEAR}')

# 解析国家 -> [(新指标名, 列名)]
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

# 排序
for country in country_metrics:
    country_metrics[country] = sorted(
        country_metrics[country],
        key=lambda x: ORDER_NEW.index(x[0]) if x[0] in ORDER_NEW else 99
    )

countries = list(country_metrics.keys())
print('国家:', countries)
for c in countries:
    print(f'  {c}: {[m[0] for m in country_metrics[c]]}')

YEARS = sorted(df_w['年份'].unique().tolist())
LATEST = YEARS[-1]
HIST_COLOR = {2019: '#fee5d9', 2020: '#fcae91', 2021: '#fb6a4a', 2022: '#de2d26',
              2023: '#cb1817', 2024: '#67000d', 2025: '#08519c'}
LATEST_COLOR = '#ff0000'
MONTH_NAMES = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']

# ========== 1. 计算同比环比数据 ==========
latest_row = df_w[df_w['日期'] == df_w['日期'].max()].iloc[0]
prev_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(months=1))].iloc[0]
yoy_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(years=1))].iloc[0]

LATEST_LBL = latest_row['日期'].strftime('%Y-%m')
PREV_LBL = prev_row['日期'].strftime('%Y-%m')
YOY_LBL = yoy_row['日期'].strftime('%Y-%m')

cmp_rows = []
for country in countries:
    for new_metric, col in country_metrics[country]:
        cur = pd.to_numeric(latest_row[col], errors='coerce')
        prv = pd.to_numeric(prev_row[col], errors='coerce')
        yoy = pd.to_numeric(yoy_row[col], errors='coerce')
        mom = (cur - prv) / prv * 100 if prv and prv != 0 and not pd.isna(prv) else None
        yoy_pct = (cur - yoy) / yoy * 100 if yoy and yoy != 0 and not pd.isna(yoy) else None
        cmp_rows.append({
            '国家': country,
            '指标': new_metric,
            LATEST_LBL: cur,
            PREV_LBL: prv,
            YOY_LBL: yoy,
            '环比(%)': round(mom, 1) if mom is not None else None,
            '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None
        })
cmp_df = pd.DataFrame(cmp_rows)

# ========== 2. 同比环比表 HTML ==========
def fmt_num(v):
    if v is None or pd.isna(v):
        return '-'
    return f'{v:,.0f}'

def fmt_pct(v):
    if v is None or pd.isna(v):
        return '-'
    return f'{v:+.1f}%'

def cell_color(v):
    if v is None or pd.isna(v):
        return ''
    if v > 0:
        return 'background:#fff1f0;color:#c0392b;font-weight:600;'  # 涨红
    elif v < 0:
        return 'background:#f6ffed;color:#27ae60;font-weight:600;'  # 跌绿
    return ''

rows_html = ''
for _, r in cmp_df.iterrows():
    mom_v = r['环比(%)']
    yoy_v = r['同比(%)']
    rows_html += f'''<tr>
        <td style="text-align:left;font-weight:500;">{r['国家']}</td>
        <td style="text-align:left;color:#555;">{r['指标']}</td>
        <td>{fmt_num(r[LATEST_LBL])}</td>
        <td>{fmt_num(r[PREV_LBL])}</td>
        <td>{fmt_num(r[YOY_LBL])}</td>
        <td style="{cell_color(mom_v)}">{fmt_pct(mom_v)}</td>
        <td style="{cell_color(yoy_v)}">{fmt_pct(yoy_v)}</td>
    </tr>'''

cmp_table_html = f"""
<table>
    <thead>
        <tr>
            <th>国家</th>
            <th>指标</th>
            <th>{LATEST_LBL}（最新）</th>
            <th>{PREV_LBL}</th>
            <th>{YOY_LBL}</th>
            <th>环比</th>
            <th>同比</th>
        </tr>
    </thead>
    <tbody>{rows_html}</tbody>
</table>
"""

# ========== 3. 季节性图（50格 = 10国 × 5指标） ==========
n_rows = len(countries)
n_cols = 5
sub_titles = []
for country in countries:
    for new_metric, _ in country_metrics[country]:
        sub_titles.append(f'<b>{country}</b><br>{new_metric}')

fig = make_subplots(
    rows=n_rows, cols=n_cols,
    subplot_titles=sub_titles,
    horizontal_spacing=0.035,
    vertical_spacing=0.03
)

for ri, country in enumerate(countries):
    for ci, (new_metric, col) in enumerate(country_metrics[country]):
        r = ri + 1
        c = ci + 1
        for year in YEARS:
            sub = df_w[df_w['年份'] == year].sort_values('月份')
            x_vals = sub['月份'].values
            y_vals = pd.to_numeric(sub[col], errors='coerce').values
            is_last = (year == LATEST)
            color = LATEST_COLOR if is_last else HIST_COLOR.get(year, '#999')
            width = 3.0 if is_last else 1.0
            marker_size = 6 if is_last else 2.5
            marker_line = dict(color='black', width=1) if is_last else dict(width=0)
            showlegend = (ri == 0 and ci == 0)
            fig.add_trace(
                go.Scatter(
                    x=x_vals, y=y_vals,
                    mode='lines+markers',
                    name=f'{year} (最新)' if is_last else str(year),
                    legendgroup=str(year),
                    showlegend=showlegend,
                    line=dict(color=color, width=width),
                    marker=dict(size=marker_size, line=marker_line),
                    hovertemplate=f'<b>{year}-{country}</b><br>月%{{x}} {new_metric}: %{{y:,.0f}}<extra></extra>'
                ),
                row=r, col=c
            )
        fig.update_xaxes(
            tickmode='array', tickvals=[1, 4, 7, 10],
            ticktext=['1', '4', '7', '10'],
            row=r, col=c, showgrid=False
        )
        fig.update_yaxes(row=r, col=c, showgrid=False, zeroline=False)

fig.update_layout(
    title=dict(
        text=f'<b>天然橡胶中国进口量 · 季节性 + 同比环比</b><br>'
             f'<sub>窗口 {START_YEAR}-{LATEST_YEAR} · 10 国家 × 5 指标（已剔除 400129/400130）· '
             f'<span style="color:red">红色加粗=最新年({LATEST})</span></sub>',
        x=0.5, xanchor='center',
        font=dict(size=18, family='Microsoft YaHei, Arial')
    ),
    height=2600, width=1700,
    legend=dict(orientation='h', yanchor='bottom', y=-0.003, xanchor='center', x=0.5, title='年份'),
    template='plotly_white',
    font=dict(family='Microsoft YaHei, Arial', size=10),
    margin=dict(l=50, r=20, t=140, b=40)
)

# 把图表转成完整 HTML 片段（含 div + script）
plotly_div = fig.to_html(include_plotlyjs='cdn', full_html=False, div_id='seasonal_dashboard')

# ========== 4. 整合到一个 HTML ==========
combined_html = f"""<!DOCTYPE html>
<html lang='zh-CN'>
<head>
<meta charset='utf-8'>
<title>天然橡胶中国进口量 · 季节性 + 同比环比</title>
<style>
  body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; padding: 24px;
          background: linear-gradient(135deg, #f0f4f8 0%, #e8eef3 100%); margin:0; }}
  .container {{ max-width: 1800px; margin: 0 auto; }}
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
  th {{ background: #1890ff; color: white; padding: 12px 8px;
        text-align: center; font-weight: 600; }}
  td {{ padding: 9px 8px; border-bottom: 1px solid #ecf0f1;
        text-align: center; color: #2c3e50; }}
  tr:nth-child(even) {{ background: #fafbfc; }}
  tr:hover {{ background: #e6f7ff !important; }}
</style>
</head>
<body>
<div class='container'>

<h1>天然橡胶中国进口量 · 季节性 + 同比环比 一体化分析</h1>
<div class='subtitle'>
  数据源：天然橡胶进口合并报表.xlsx（按国别） · 窗口 2019-2026 · 最新月份 {LATEST_LBL}
</div>

<div class='section'>
  <div class='section-title'>一、同比 / 环比 变化表（最新月 {LATEST_LBL}）</div>
  <div class='legend-box'>
    <b>对比口径：</b>{LATEST_LBL}（最新） · {PREV_LBL}（环比） · {YOY_LBL}（同比）<br>
    <b>配色：</b>按中国股市惯例，<span style='color:#c0392b;'>上涨浅红</span> · <span style='color:#27ae60;'>下跌浅绿</span> · <b>数据"- "</b>表示基期或本期为零/缺失<br>
    <b>指标体系：</b>乳胶=HS 400110 · 烟片=HS 400121 · 中国标胶=HS 400122 · 天然橡胶=该国天然胶汇总 · 混合橡胶=HS 400280
  </div>
  {cmp_table_html}
</div>

<div class='section'>
  <div class='section-title'>二、季节性分析 Dashboard（10 国家 × 5 指标 = 50 格）</div>
  {plotly_div}
</div>

</div>
</body>
</html>
"""

OUT_COMBINED = os.path.join(OUT_DIR, 'dashboard_all.html')
with open(OUT_COMBINED, 'w', encoding='utf-8') as f:
    f.write(combined_html)
print(f'[OK] 综合 HTML: {OUT_COMBINED}')
print(f'文件大小: {os.path.getsize(OUT_COMBINED)/1024/1024:.2f} MB')

# 同时也保留单独的xlsx供二次加工
xlsx_path = os.path.join(OUT_DIR, '同比环比变化表.xlsx')
cmp_df.to_excel(xlsx_path, index=False)
print(f'[OK] xlsx: {xlsx_path}')

# 打印关键发现
print('\n=== 关键发现 ===')
non_null = cmp_df.dropna(subset=['同比(%)'])
top5 = non_null.nlargest(5, '同比(%)')[['国家', '指标', '同比(%)']]
bot5 = non_null.nsmallest(5, '同比(%)')[['国家', '指标', '同比(%)']]
print('同比涨幅 TOP5:')
print(top5.to_string(index=False))
print('同比降幅 TOP5:')
print(bot5.to_string(index=False))