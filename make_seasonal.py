"""
天然橡胶进口合并报表 - 季节性图（v2）
- 最新年份(2026)线条标红加粗突出
- 生成每个国家最新月份同比/环比变化表
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

country_metrics = {}
for col in df_w.columns[1:-2]:
    ctype, metric = col
    m = re.match(r'(.+?)(天然橡胶|混合橡胶)$', ctype)
    if not m:
        continue
    country, _ = m.groups()
    country_metrics.setdefault(country, []).append((metric, col))

ORDER = ['400110进口量', '400121进口量', '400122进口量', '400129进口量', '400130进口量', '总计进口量', '进口量']
for country in country_metrics:
    country_metrics[country] = sorted(country_metrics[country], key=lambda x: ORDER.index(x[0]) if x[0] in ORDER else 99)

YEARS = sorted(df_w['年份'].unique().tolist())
LATEST = YEARS[-1]
# 历史年份（红橙渐变）
HIST_COLOR = {2019: '#fee5d9', 2020: '#fcae91', 2021: '#fb6a4a', 2022: '#de2d26', 2023: '#cb1817', 2024: '#67000d', 2025: '#08519c'}
HIST_W = {2019: 1.2, 2020: 1.2, 2021: 1.2, 2022: 1.2, 2023: 1.2, 2024: 1.2, 2025: 1.5}
HIST_MS = {2019: 3, 2020: 3, 2021: 3, 2022: 3, 2023: 3, 2024: 3, 2025: 4}
# 最新年份红色加粗
LATEST_COLOR = '#ff0000'
LATEST_W = 3.5
LATEST_MS = 7

MONTH_NAMES = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']


def style_for(year):
    if year == LATEST:
        return dict(color=LATEST_COLOR, width=LATEST_W, marker_size=LATEST_MS, dash='solid')
    else:
        return dict(color=HIST_COLOR.get(year, '#999'),
                    width=HIST_W.get(year, 1.2),
                    marker_size=HIST_MS.get(year, 3),
                    dash='solid')


def build_country_figure(country, metrics_list):
    n = len(metrics_list)
    rows = 2
    cols = 4
    subplot_titles = [m[0].replace('进口量', '') for m in metrics_list] + [''] * (rows * cols - n)
    fig = make_subplots(
        rows=rows, cols=cols,
        subplot_titles=subplot_titles,
        horizontal_spacing=0.05,
        vertical_spacing=0.20
    )
    for idx, (metric_name, col) in enumerate(metrics_list):
        r = idx // cols + 1
        c = idx % cols + 1
        for year in YEARS:
            sub = df_w[df_w['年份'] == year].sort_values('月份')
            x_vals = sub['月份'].values
            y_vals = pd.to_numeric(sub[col], errors='coerce').values
            st = style_for(year)
            showlegend = (idx == 0)
            fig.add_trace(
                go.Scatter(
                    x=x_vals, y=y_vals,
                    mode='lines+markers',
                    name=str(year),
                    legendgroup=str(year),
                    showlegend=showlegend,
                    line=dict(color=st['color'], width=st['width'], dash=st['dash']),
                    marker=dict(size=st['marker_size'],
                                line=dict(color='black', width=1) if year == LATEST else dict(width=0)),
                    hovertemplate=f'<b>{year}年 %{{x}}月</b><br>{metric_name}: %{{y:,.0f}}<extra></extra>'
                ),
                row=r, col=c
            )
        fig.update_xaxes(tickmode='array', tickvals=list(range(1, 13)), ticktext=MONTH_NAMES, row=r, col=c)
    fig.update_layout(
        title=dict(
            text=f'<b>{country} 进口量季节性分析</b> · {START_YEAR}-{LATEST_YEAR}<br><sub>红色加粗 = 最新年({LATEST})</sub>',
            x=0.5, xanchor='center', font=dict(size=18, family='Microsoft YaHei, Arial')
        ),
        height=720, width=1500,
        legend=dict(orientation='h', yanchor='bottom', y=-0.10, xanchor='center', x=0.5, title='年份'),
        template='plotly_white',
        font=dict(family='Microsoft YaHei, Arial', size=11)
    )
    return fig


for country, metrics_list in country_metrics.items():
    fig = build_country_figure(country, metrics_list)
    out = os.path.join(OUT_DIR, f'seasonal_{country}.html')
    fig.write_html(out, include_plotlyjs='cdn')
    print(f'[OK] {country}')

# ========== 同比环比变化表 ==========
# 最新月 / 上月 / 去年同月
latest_row = df_w[df_w['日期'] == df_w['日期'].max()].iloc[0]
prev_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(months=1))].iloc[0]
yoy_row = df_w[df_w['日期'] == (latest_row['日期'] - pd.DateOffset(years=1))].iloc[0]

latest_label = latest_row['日期'].strftime('%Y-%m')
prev_label = prev_row['日期'].strftime('%Y-%m')
yoy_label = yoy_row['日期'].strftime('%Y-%m')

print(f'\n最新: {latest_label} | 上月: {prev_label} | 去年同期: {yoy_label}')

# 收集所有指标列
all_country_rows = []
for country, metrics_list in country_metrics.items():
    for metric_name, col in metrics_list:
        cur = pd.to_numeric(latest_row[col], errors='coerce')
        prv = pd.to_numeric(prev_row[col], errors='coerce')
        yoy = pd.to_numeric(yoy_row[col], errors='coerce')
        mom = (cur - prv) / prv * 100 if prv and prv != 0 else None
        yoy_pct = (cur - yoy) / yoy * 100 if yoy and yoy != 0 else None
        all_country_rows.append({
            '国家': country,
            '指标': metric_name,
            latest_label: cur,
            prev_label: prv,
            yoy_label: yoy,
            '环比(%)': round(mom, 1) if mom is not None else None,
            '同比(%)': round(yoy_pct, 1) if yoy_pct is not None else None
        })

cmp_df = pd.DataFrame(all_country_rows)
# 保存为 xlsx 和 html
xlsx_path = os.path.join(OUT_DIR, '同比环比变化表.xlsx')
cmp_df.to_excel(xlsx_path, index=False)
print(f'[OK] Excel: {xlsx_path}')

# HTML 表格（带同比环比颜色）
def color_pct(v):
    if v is None or pd.isna(v):
        return ''
    if v > 0:
        return 'background-color:#fee; color:#c00;'  # 涨红
    elif v < 0:
        return 'background-color:#efe; color:#060;'  # 跌绿（中国惯例）
    return ''

styled = cmp_df.style.format({
    latest_label: '{:,.0f}',
    prev_label: '{:,.0f}',
    yoy_label: '{:,.0f}',
    '环比(%)': lambda v: f'{v:+.1f}%' if v is not None and not pd.isna(v) else '-',
    '同比(%)': lambda v: f'{v:+.1f}%' if v is not None and not pd.isna(v) else '-'
}, na_rep='-').applymap(color_pct, subset=['环比(%)', '同比(%)'])

html_path = os.path.join(OUT_DIR, '同比环比变化表.html')
html_str = styled.to_html()
# 包装完整 HTML
wrapper = f"""<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<title>同比环比变化表</title>
<style>
body {{ font-family: 'Microsoft YaHei', Arial; padding: 20px; background: #f5f7fa; }}
h2 {{ color: #2c3e50; }}
.note {{ background: #fffbe6; padding: 10px; border-left: 4px solid #faad14; margin-bottom: 20px; }}
table {{ border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }}
th {{ background: #1890ff; color: white; padding: 10px; text-align: center; }}
td {{ padding: 8px; border-bottom: 1px solid #e8e8e8; text-align: center; }}
tr:hover {{ background: #f0f7ff; }}
td:first-child {{ text-align: left; font-weight: 500; }}
td:nth-child(2) {{ text-align: left; color: #555; }}
</style></head>
<body>
<h2>天然橡胶进口 - 最新数据同比环比变化</h2>
<div class='note'>
<b>对比口径：</b>最新月 <b>{latest_label}</b> · 上月 <b>{prev_label}</b>（环比）· 去年同期 <b>{yoy_label}</b>（同比）<br>
<b>配色规则（按中国股市惯例）：</b>同比/环比上涨 → 浅红底；下跌 → 浅绿底
</div>
{html_str}
</body></html>
"""
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(wrapper)
print(f'[OK] HTML: {html_path}')

# 终端打印前几行
print('\n部分数据预览:')
print(cmp_df.head(10).to_string(index=False))