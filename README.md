# 中国月度进口（分国别）— 自动化分析流程

## 文件清单

| 文件 | 角色 |
|---|---|
| `seasonal_dashboard.py` | 核心模块：所有计算、图表、HTML 生成的函数；可被任意脚本 import |
| `update_dashboard.py` | CLI 入口：一行命令重新生成 dashboard + xlsx |
| `dashboard_total.html` | 输出：网页（含同比环比表 / 涨跌速览 / 10 国总量图 / 50 格季节性） |
| `同比环比变化表.xlsx` | 输出：表格 |
| `make_total.py` | 旧脚本（保留可读，已被 `update_dashboard.py` 替代） |

## 一行命令执行

```bash
# 默认路径
python update_dashboard.py

# 自定义 xlsx 路径
python update_dashboard.py "D:\\AAA东方嘉信\\备份数据\\qinrex\\天然橡胶进口合并报表.xlsx"

# 自定义输出目录
python update_dashboard.py "D:\\data.xlsx" "D:\\out"
```

## 作为模块调用

```python
from seasonal_dashboard import Config, build
cfg = Config(file_path=r'D:\xxx.xlsx', out_dir=r'D:\out')
result = build(cfg)
print(result['latest_month'], result['figures'])
```

## 触发方式（任选）

1. 手动：每日跑一次 `update_dashboard.py`，或接入到 `merge_and_scrape.py` 末尾追加一行调用 `build()`
2. 定时：已注册 WorkBuddy 自动化任务「中国进口分析-每日更新」，每日 08:30 自动重建页面
3. 监控：如果需要检测文件变化后自动刷新，可注册 mtime 触发（当前未启用）

## 维护要点

- 所有可调项集中在 `seasonal_dashboard.Config` 类（列重命名、剔除、配色、窗口等）
- 新增国家：自动识别（按列头「{国家}{天然橡胶|混合橡胶}」拆分），无需改代码
- 新增指标：在 `Config.rename_map` 加一行映射
- 剔除指标（如新增 4002XX 编码）：在 `Config.exclude_metrics` 加一项
- 最新年份颜色必须为 `#e02020`（红色加粗），否则参考图样式失效
- 任何月度数据更新后再次执行即可，无须改代码

## 输出产物验证

执行成功后控制台会打印：
```
[OK] HTML  -> D:\...\dashboard_total.html
[OK] XLSX -> D:\...\同比环比变化表.xlsx
表格 N 行，最新月 YYYY-MM，图表 N 张
```