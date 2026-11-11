"""
汇总仪表板 v2：所有国家×指标 70 格季节性图
最新年份(2026)线条红色加粗突出
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import re
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

FILE = r'D:\AAA东方嘉信\备份数据\qinrex\天然橡胶进口合并报表.xlsx'
OUT = r'D:\AAA东方嘉信\备份数据\qinrex\_analysis\seasonal\dashboard_summary.html'

df = pd.read_excel(FILE, sheet_name='按国别', header=[0, 1])
df = df.iloc[1:].reset_index(drop=True)
df.columns = ['日期'] + list(df.columns[1:])
df['日期'] = pd.to_datetime(df['日期'])
df['年份'] = df['日期'].dt.year
df['月份'] = df['日期'].dt.month

LATEST_YEAR = int(df['年份'].max())
START_YEAR = LATEST_YEAR - 7
df_w = df[df['年份'] >= START_YEAR].copy()

country_metrics = {}
for col in df_w.columns[1:-2]:
    ctype, metric = col
    m = re.match(r'(.+?)(天然橡胶|混合橡胶)$', ctype)
    if not m:
        continue
    country, _ = m.groups()
    country_metrics.setdefault(country, []).append((metric, col))

ORDER = ['400110进口量', '400121进口量', '400122进口量', '400129进口量', '400130进口量', '总计进口量', '进口量']
countries = list(country_metrics.keys())
for country in country_metrics:
    country_metrics[country] = sorted(country_metrics[country], key=lambda x: ORDER.index(x[0]) if x[0] in ORDER else 99)

YEARS = sorted(df_w['年份'].unique().tolist())
LATEST = YEARS[-1]
HIST_COLOR = {2019: '#fee5d9', 2020: '#fcae91', 2021: '#fb6a4a', 2022: '#de2d26', 2023: '#cb1817', 2024: '#67000d', 2025: '#08519c'}
LATEST_COLOR = '#ff0000'
MONTH_NAMES = ['1', '4', '7', '10']

n_rows = len(countries)
n_cols = 7
sub_titles = []
for country in countries:
    for m, _ in country_metrics[country]:
        short = m.replace('进口量', '')
        sub_titles.append(f'<b>{country}</b><br>{short}')

fig = make_subplots(
    rows=n_rows, cols=n_cols,
    subplot_titles=sub_titles,
    horizontal_spacing=0.025,
    vertical_spacing=0.025
)

for ri, country in enumerate(countries):
    for ci, (metric_name, col) in enumerate(country_metrics[country]):
        r = ri + 1
        c = ci + 1
        for year in YEARS:
            sub = df_w[df_w['年份'] == year].sort_values('月份')
            x_vals = sub['月份'].values
            y_vals = pd.to_numeric(sub[col], errors='coerce').values
            is_last = year == LATEST
            color = LATEST_COLOR if is_last else HIST_COLOR.get(year, '#999')
            width = 3.2 if is_last else 1.0
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
                    hovertemplate=f'<b>{year}-{country}</b><br>月%{{x}} {metric_name}: %{{y:,.0f}}<extra></extra>'
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
        text=f'<b>天然橡胶进口合并报表 · 季节性分析 Dashboard</b><br><sub>窗口 {START_YEAR}-{LATEST_YEAR} · 10 国家 × 7 指标 · <span style="color:red">红色加粗=最新年({LATEST})</span></sub>',
        x=0.5, xanchor='center',
        font=dict(size=18, family='Microsoft YaHei, Arial')
    ),
    height=2400, width=1800,
    legend=dict(orientation='h', yanchor='bottom', y=-0.005, xanchor='center', x=0.5, title='年份'),
    template='plotly_white',
    font=dict(family='Microsoft YaHei, Arial', size=9),
    margin=dict(l=50, r=20, t=120, b=50)
)

fig.write_html(OUT, include_plotlyjs='cdn')
print(f'[OK] {OUT}')
print(f'文件大小: {os.path.getsize(OUT)/1024/1024:.2f} MB')