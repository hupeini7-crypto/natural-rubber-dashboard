"""
daily_dashboard.py —— 天然橡胶进出口 Dashboard 每日自动发布流程
=============================================================
每日 08:30 由自动化调用（或手动运行）。
完整执行「检查数据 → 备份 → 生成 → 校验 → 公网服务 → 汇报」。

步骤：
  1) 检查两个 Excel 是否存在/可读/最新月份/异常值
  2) 备份上一版 dashboard_total.html（异常时回滚）
  3) 运行 build() 生成（含五国出口 + 进出口对比）
  4) 校验网页（原模块 + 新增模块 + 图表数 + 无异常占位）
  5) 检查 8899（被占用则 8901）HTTP 服务与 ngrok，自动启动，重试一次
  6) 从 127.0.0.1:4040 获取公网 URL 并验证 /dashboard_total.html 可访问
  7) 输出每日更新报告（打印 + 写日志）

公网目录仅发布 HTML；不公开原始 Excel / 代码 / 密钥。
"""
from __future__ import annotations
import os, sys, shutil, subprocess, socket, urllib.request, json, time, datetime, traceback
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

PY = r'D:\Program Files\Python312\python.exe'
NGROK = r'C:\Users\Lenovo\tools\ngrok\ngrok.exe'
IMPORT_XLSX = r'D:\AAA东方嘉信\备份数据\qinrex\天然橡胶进口合并报表.xlsx'
EXPORT_XLSX = r'D:\AAA东方嘉信\备份数据\qinrex\五国橡胶出口数据.xlsx'
OUT_DIR = HERE
HTML_NAME = 'dashboard_total.html'
BACKUP_DIR = HERE / 'backups'
LOG_DIR = HERE / 'logs'
SERVE_PORT_PREF = 8899
SERVE_PORT_FALLBACK = 8901

BACKUP_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)


def log(msg):
    ts = datetime.datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}')


# ---------------- 1. 数据检查 ----------------
def check_excel(path, label):
    res = {'label': label, 'ok': False, 'latest': 'N/A', 'rows': 0, 'note': ''}
    try:
        import pandas as pd
        xls = pd.ExcelFile(path)
        res['sheets'] = xls.sheet_names
        res['ok'] = True
    except Exception as e:
        res['note'] = f'读取失败: {e}'
    return res


def latest_month_of_sheet(path, sheet, header=(0, 1)):
    import pandas as pd
    df = pd.read_excel(path, sheet_name=sheet, header=list(header))
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = ['日期'] + list(df.columns[1:])
    d = pd.to_datetime(df['日期'], errors='coerce').dropna()
    return d.max().strftime('%Y-%m') if len(d) else 'N/A'


# ---------------- 5. 端口 / ngrok ----------------
def port_responding(port):
    try:
        r = urllib.request.urlopen(f'http://127.0.0.1:{port}/{HTML_NAME}', timeout=2)
        return r.status == 200
    except Exception:
        return False


def is_port_listening(port):
    """端口是否有任何服务在监听（不一定是我们的）。"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(1)
        ok = s.connect_ex(('127.0.0.1', port)) == 0
        s.close()
        return ok
    except Exception:
        return False


def start_http_server(port, serve_dir):
    subprocess.Popen(
        [PY, '-m', 'http.server', str(port), '--bind', '0.0.0.0'],
        cwd=str(serve_dir),
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )


def ensure_http_server():
    """返回实际使用的端口（8899 优先，被其它项目占用则 8901）。"""
    if port_responding(SERVE_PORT_PREF):
        log(f'HTTP 服务已在 {SERVE_PORT_PREF} 提供本 Dashboard，复用。')
        return SERVE_PORT_PREF
    if not is_port_listening(SERVE_PORT_PREF):
        log(f'启动 HTTP 服务 {SERVE_PORT_PREF} ...')
        start_http_server(SERVE_PORT_PREF, OUT_DIR)
        time.sleep(2.5)
        if port_responding(SERVE_PORT_PREF):
            return SERVE_PORT_PREF
    # 端口在监听但未响应本文件 -> 短暂重试，仍失败再判为被占用
    time.sleep(1.5)
    if port_responding(SERVE_PORT_PREF):
        return SERVE_PORT_PREF
    # 8899 被占用且非本服务 -> 改用 8901
    log(f'{SERVE_PORT_PREF} 被其它项目占用，改用 {SERVE_PORT_FALLBACK}。')
    if not is_port_listening(SERVE_PORT_FALLBACK):
        start_http_server(SERVE_PORT_FALLBACK, OUT_DIR)
        time.sleep(2.5)
    return SERVE_PORT_FALLBACK


def ngrok_url():
    try:
        r = urllib.request.urlopen('http://127.0.0.1:4040/api/tunnels', timeout=3)
        d = json.loads(r.read().decode('utf-8'))
        for t in d.get('tunnels', []):
            if t.get('proto') == 'https' and t.get('public_url'):
                return t['public_url']
    except Exception:
        pass
    return None


def ensure_ngrok(port):
    u = ngrok_url()
    if u:
        # 校验是否指向当前端口
        try:
            r = urllib.request.urlopen('http://127.0.0.1:4040/api/tunnels', timeout=3)
            d = json.loads(r.read().decode('utf-8'))
            for t in d.get('tunnels', []):
                if t.get('proto') == 'https' and str(port) in t.get('config', {}).get('addr', ''):
                    return u
        except Exception:
            pass
        # 已有隧道但指向其它端口 -> 重启 ngrok 指向本端口（hostname 不变）
        log('现有 ngrok 隧道端口不匹配，重启指向本端口 ...')
        subprocess.run(['taskkill', '/IM', 'ngrok.exe', '/F'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1)
    else:
        log('启动 ngrok 隧道 ...')
    subprocess.Popen(
        [NGROK, 'http', str(port)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    time.sleep(6)
    return ngrok_url()


# ---------------- 校验 ----------------
def validate_html(path):
    txt = Path(path).read_text(encoding='utf-8')
    need = ['七、五国橡胶出口分析', '八、中国进口 × 五国出口 对比',
            '8.1 进口 vs 全球出口 对比总表',
            '中国月度进口数据', '五国月度出口数据', '进出口数据对比']
    forbidden = ['8.6 真实对华出口占比', '真实对华占比', '真实输华份额']
    miss = [n for n in need if n not in txt]
    bad = [n for n in forbidden if n in txt]
    figs = txt.count('Plotly.newPlot')
    err = '本模块生成异常' in txt
    ok = (len(miss) == 0 and len(bad) == 0 and not err and figs > 0)
    return ok, miss, figs, bad


# ---------------- 主流程 ----------------
def main():
    today = datetime.date.today().strftime('%Y-%m-%d')
    log('=' * 60)
    log(f'天然橡胶进出口 Dashboard 每日更新 · {today}')
    report = {'date': today, 'steps': {}}

    # 1) 数据检查
    imp_ok = check_excel(IMPORT_XLSX, '中国进口')
    exp_ok = check_excel(EXPORT_XLSX, '五国出口')
    imp_latest = exp_latest = 'N/A'
    try:
        imp_latest = latest_month_of_sheet(IMPORT_XLSX, '按国别')
    except Exception as e:
        imp_ok['note'] = str(e)
    try:
        exp_latest = latest_month_of_sheet(EXPORT_XLSX, '按国别')
    except Exception as e:
        exp_ok['note'] = str(e)
    data_ok = imp_ok['ok'] and exp_ok['ok']
    report['import'] = {'ok': imp_ok['ok'], 'latest': imp_latest, 'note': imp_ok.get('note', '')}
    report['export'] = {'ok': exp_ok['ok'], 'latest': exp_latest, 'note': exp_ok.get('note', '')}
    log(f"数据检查: 进口={'OK' if imp_ok['ok'] else 'FAIL'} 最新{imp_latest} | "
        f"出口={'OK' if exp_ok['ok'] else 'FAIL'} 最新{exp_latest}")

    # 2) 备份上一版
    html_path = OUT_DIR / HTML_NAME
    backup_path = BACKUP_DIR / f'{HTML_NAME}.{today}.bak'
    if html_path.exists():
        shutil.copyfile(html_path, backup_path)
        log(f'已备份上一版: {backup_path.name}')

    # 3) 生成
    build_ok = False
    build_info = {}
    try:
        from seasonal_dashboard import Config, build
        cfg = Config()
        res = build(cfg)
        build_ok = True
        build_info = {
            'latest_month': res.get('latest_month'),
            'figures': res.get('figures'),
            'rows': res.get('rows_in_table'),
            'export': res.get('export', {}),
        }
        report['export_report'] = res.get('export', {})
        log(f"生成成功: 最新月 {res.get('latest_month')} · 图表 {res.get('figures')} 张")
    except Exception as e:
        log(f'生成失败: {e}')
        log(traceback.format_exc())

    # 4) 校验
    valid = False
    miss = []
    figs = 0
    if build_ok and html_path.exists():
        valid, miss, figs, bad = validate_html(html_path)
    report['web'] = {'ok': valid, 'figures': figs, 'missing': miss, 'forbidden': bad}
    if not valid:
        log(f'校验失败: 缺失 {miss}，违禁 {bad}，准备回滚到备份')
        if backup_path.exists():
            shutil.copyfile(backup_path, html_path)
            log(f'已回滚到上一版: {backup_path.name}')

    # 5/6) 公网服务
    public_url = None
    port = None
    if valid:
        port = ensure_http_server()
        for attempt in (1, 2):
            public_url = ensure_ngrok(port)
            if public_url:
                break
            log(f'ngrok 第 {attempt} 次重试...')
            time.sleep(3)
        # 验证公网可访问
        if public_url:
            full = f'{public_url}/{HTML_NAME}'
            try:
                r = urllib.request.urlopen(full, timeout=8)
                ok = r.status == 200
                report['public'] = {'ok': ok, 'url': full}
                log(f'公网访问: {"OK" if ok else "FAIL"} · {full}')
            except Exception as e:
                report['public'] = {'ok': False, 'url': full, 'err': str(e)}
                log(f'公网访问验证失败: {e}')
        else:
            report['public'] = {'ok': False, 'url': None}
            log('未能获取 ngrok 公网 URL')
    else:
        report['public'] = {'ok': False, 'url': None, 'err': '网页校验未通过，未启动公网'}

    # 7) 报告
    print_report(report)
    # 写日志
    log_path = LOG_DIR / f'report_{today}.md'
    log_path.write_text(format_report_md(report), encoding='utf-8')
    log(f'报告已写: {log_path}')
    return report


def print_report(r):
    imp = r['import']; exp = r['export']; web = r['web']; pub = r['public']
    exp_repo = r.get('export_report', {})
    print('\n' + '=' * 56)
    print('[报告] 天然橡胶进出口 Dashboard 更新报告')
    print('=' * 56)
    print(f"- 中国进口数据：{'成功' if imp['ok'] else '失败'}，最新月份 {imp['latest']}")
    print(f"- 五国出口数据：{'成功' if exp['ok'] else '失败'}，各国最新月份见下")
    if isinstance(exp_repo, dict) and exp_repo.get('export_latest_months'):
        for c, m in exp_repo['export_latest_months'].items():
            print(f"    · {c}: {m}")
    print(f"- 五国共同可比月份：{exp_repo.get('common_month','N/A') if isinstance(exp_repo,dict) else 'N/A'}")
    print(f"- 网页生成：{'成功' if web['ok'] else '失败'}（图表 {web.get('figures',0)} 张）")
    print(f"- 公网访问：{'成功' if pub.get('ok') else '失败'}")
    if pub.get('url'):
        print(f"  公网链接：{pub['url']}")
    # 重点变化
    print('- 重点变化：')
    if isinstance(exp_repo, dict):
        bc = exp_repo.get('biggest_ratio_change')
        if bc:
            print(f"    · 进口/出口比值变化最大国家：{bc[0]}（{bc[1]:+.2f} pp）")
        print(f"    · 全国口径来源：{exp_repo.get('fine_split_src','-')}")
    # 异常
    print('- 异常说明：')
    notes = []
    if not imp['ok']: notes.append(f"进口Excel: {imp.get('note','')}")
    if not exp['ok']: notes.append(f"出口Excel: {exp.get('note','')}")
    if web.get('missing'): notes.append(f"缺失模块: {web['missing']}")
    if pub.get('err'): notes.append(f"公网: {pub['err']}")
    if not notes:
        notes.append('无')
    for n in notes:
        print(f"    · {n}")
    print('=' * 56 + '\n')


def format_report_md(r):
    imp = r['import']; exp = r['export']; web = r['web']; pub = r['public']
    er = r.get('export_report', {})
    lines = [f"# 天然橡胶进出口 Dashboard 更新报告 · {r['date']}", '',
             f"- 中国进口数据：{'成功' if imp['ok'] else '失败'}，最新月份 {imp['latest']}",
             f"- 五国出口数据：{'成功' if exp['ok'] else '失败'}，各国最新：{exp.get('latest','N/A')}",
             f"- 网页生成：{'成功' if web['ok'] else '失败'}（图表 {web.get('figures',0)} 张）",
             f"- 公网访问：{'成功' if pub.get('ok') else '失败'}",
             f"- 🌐 公网链接：{pub.get('url') or '无'}",
             f"- 五国共同可比月份：{er.get('common_month','N/A')}",
             f"- 全国口径来源：{er.get('fine_split_src','-')}", '']
    return '\n'.join(lines)


if __name__ == '__main__':
    main()
