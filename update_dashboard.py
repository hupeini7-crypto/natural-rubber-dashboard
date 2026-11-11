"""
update_dashboard.py —— CLI 入口
================================
执行：python update_dashboard.py [可选：xlsx_path] [可选：out_dir]
不传参则使用 seasonal_dashboard.Config 的默认路径。

示例：
  python update_dashboard.py
  python update_dashboard.py "D:\\AAA东方嘉信\\备份数据\\qinrex\\天然橡胶进口合并报表.xlsx"
  python update_dashboard.py "D:\\new.xlsx" "D:\\new_out"
"""
import sys
from pathlib import Path

# 让同目录的模块可被 import
HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from seasonal_dashboard import Config, build

def main():
    cfg = Config()
    if len(sys.argv) >= 2:
        cfg.file_path = sys.argv[1]
    if len(sys.argv) >= 3:
        cfg.out_dir = sys.argv[2]
    print(f"[入口] 数据源: {cfg.file_path}")
    print(f"[入口] 输出目录: {cfg.out_dir}")
    result = build(cfg)
    print(f"\n[完成] 表格 {result['rows_in_table']} 行 · {result['figures']} 张图表 · 最新月 {result['latest_month']}")
    print(f"   HTML : {result['html']}")
    print(f"   XLSX : {result['xlsx']}")
    if result.get('url'):
        print(f"   手机访问 : {result['url']}")
    return 0

if __name__ == '__main__':
    sys.exit(main())