#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DK 变化记录表 —— 解析 / 分析 / 生成网页 引擎

用法:
  处理某一天的两个表:
    python process.py --k 2026-08-14_K.xlsx --d 2026-08-14_D.xlsx
    (日期优先取文件名里的 YYYY-MM-DD，也可用 --date 2026-08-14 指定)

  仅用已有数据重新生成网页:
    python process.py --regen

  自检(合成数据，不碰真实数据):
    python process.py --selftest
"""
import argparse
import datetime
import json
import os
import re
import urllib.request

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE, "config.json")
RECORDS_PATH = os.path.join(BASE, "data", "records.json")
STOCKS_PATH = os.path.join(BASE, "data", "stocks.json")
HTML_PATH = os.path.join(BASE, "index.html")
OBS_PATH = os.path.join(BASE, "data", "obs.json")
APP_TEMPLATE_PATH = os.path.join(BASE, "app_template.html")
OBS_WINDOW = 30            # 观察窗口（自然日）
BENCH_SECID = "1.000985"   # 中证全指（前复权）
BENCH_NAME = "中证全指"


# ---------- 基础工具 ----------
def load_config(path=CONFIG_PATH):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def find_col(headers, candidates):
    low = {h.lower(): i for i, h in enumerate(headers) if h}
    # 1) 精确匹配
    for c in candidates:
        if c in low:
            return low[c]
        if c.lower() in low:
            return low[c.lower()]
    # 2) 子串回退（兼容「涨跌幅(13:25)」这类带后缀表头）
    for c in candidates:
        cl = c.lower()
        if not cl:
            continue
        for h, i in low.items():
            if cl in h:
                return i
    return None


def normalize_code(val, pad):
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    s = s.lstrip("'").strip()
    if pad and s.isdigit():
        s = s.zfill(6)
    return s


def parse_cap(val):
    """把「总市值」列文本（如 '230.24亿' / '4.14亿' / '12.3万'）解析为 亿元 数值。无则返回 None。"""
    if val is None:
        return None
    s = str(val).strip().replace(",", "").replace("，", "")
    if not s:
        return None
    m = re.match(r"^([\d.]+)\s*(千亿|万亿|亿|万)?", s)
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    unit = m.group(2)
    if unit == "亿":
        return num
    if unit == "万":
        return num / 1e4
    if unit in ("千亿", "万亿"):
        return num * 1000.0
    return num / 1e8  # 裸数字视为「元」


def parse_date_from_filename(name, config):
    rx = config.get("file_naming", {}).get("date_in_filename_regex")
    if not rx:
        return None
    m = re.search(rx, os.path.basename(name))
    return m.group(1) if m else None


# ---------- 读取 xlsx ----------
def read_xlsx(path, config, kind=None):
    """读取一个 xlsx。
    kind='K'/'D' 时优先按「选股格式」处理：找 K点/D点 列，值为「符合」的才是信号。
    找不到标记列则按「精简清单格式」处理：整表每行都是信号。
    """
    if load_workbook is None:
        raise SystemExit("[错误] 未安装 openpyxl，无法读取 xlsx。请先 pip install openpyxl。")
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    hr = config.get("sheet", {}).get("header_row", 0)
    headers = [str(c).strip() if c is not None else "" for c in rows[hr]]
    code_i = find_col(headers, config["sheet"]["code_col_candidates"])
    name_i = find_col(headers, config["sheet"].get("name_col_candidates", []))
    metric_i = find_col(headers, config["sheet"].get("metric_col_candidates", []))
    cap_i = find_col(headers, config["sheet"].get("cap_col_candidates", []))
    if code_i is None:
        raise SystemExit(
            "[错误] 在 %s 找不到代码列。\n表头: %s\n候选名: %s"
            % (os.path.basename(path), headers, config["sheet"]["code_col_candidates"])
        )
    # 选股格式：K点/D点 列值为「符合」的才是信号
    marker_i = None
    if kind == "K":
        marker_i = find_col(headers, ["K点"])
    elif kind == "D":
        marker_i = find_col(headers, ["D点"])
    mode = "选股(符合筛选)" if marker_i is not None else "精简清单(整表)"
    out = []
    for r in rows[hr + 1:]:
        if not r or all(c is None for c in r):
            continue
        if marker_i is not None:
            v = r[marker_i]
            if v is None or str(v).strip() != "符合":
                continue
        raw = r[code_i]
        if raw is None:
            continue
        code = normalize_code(raw, config["analysis"]["code_pad_to_6"])
        if not code:
            continue
        name = str(r[name_i]).strip() if (name_i is not None and r[name_i] is not None) else ""
        metric = None
        if metric_i is not None and r[metric_i] is not None:
            try:
                metric = float(r[metric_i])
            except Exception:
                metric = None
        cap = parse_cap(r[cap_i]) if cap_i is not None else None
        out.append({"code": code, "name": name, "metric": metric, "cap": cap})
    print("  [统计] %s  解析模式=%s  命中=%d行  含市值=%d行" % (
        os.path.basename(path), mode, len(out), sum(1 for x in out if x["cap"] is not None)))
    return out


# ---------- 处理一天 ----------
def process_day(k_path, d_path, date, config):
    k = read_xlsx(k_path, config, "K")
    d = read_xlsx(d_path, config, "D")
    print("  [统计] %s  K表=%d行  D表=%d行" % (date, len(k), len(d)))

    stocks = load_json(STOCKS_PATH, {})
    for it in k + d:
        if it["code"] and it["name"]:
            stocks[it["code"]] = it["name"]
    save_json(STOCKS_PATH, stocks)

    k_codes = sorted({i["code"] for i in k})
    d_codes = sorted({i["code"] for i in d})
    metrics = {}
    caps = {}
    for i in k + d:
        if i["code"] and i["metric"] is not None:
            metrics[i["code"]] = i["metric"]
        if i["code"] and i["cap"] is not None:
            caps[i["code"]] = i["cap"]

    day = {"date": date, "k": k_codes, "d": d_codes, "metrics": metrics, "caps": caps}

    records = load_json(RECORDS_PATH, {"updated": None, "days": []})
    before = len(records.get("days", []))
    days = [x for x in records.get("days", []) if x["date"] != date]
    days.append(day)
    days.sort(key=lambda x: x["date"])
    records["days"] = days
    records["updated"] = date
    save_json(RECORDS_PATH, records)
    print("  [统计] 历史天数 %d -> %d；本日 K=%d  D=%d" % (before, len(days), len(k_codes), len(d_codes)))
    return day


# ---------- 分析 ----------
def build_analysis(records, stocks, config):
    days = records.get("days", [])
    updated = records.get("updated")

    stocks_set = set(stocks.keys())
    for day in days:
        stocks_set.update(day.get("k", []))
        stocks_set.update(day.get("d", []))

    latest_metrics = days[-1].get("metrics", {}) if days else {}
    latest_caps = days[-1].get("caps", {}) if days else {}
    timelines = {}
    for code in stocks_set:
        tl = []
        for day in days:
            if code in day.get("k", []):
                tl.append({"date": day["date"], "state": "K"})
            if code in day.get("d", []):
                tl.append({"date": day["date"], "state": "D"})
        timelines[code] = tl

    rows = []
    total_reversals = 0
    for code, tl in timelines.items():
        nK = sum(1 for e in tl if e["state"] == "K")
        nD = sum(1 for e in tl if e["state"] == "D")
        reversals = []
        for a, b in zip(tl, tl[1:]):
            # 仅跨天状态变化算「反转」；同日既K又D视为口径冲突，不计入预警
            if a["state"] != b["state"] and a["date"] != b["date"]:
                reversals.append({
                    "from": a["state"], "to": b["state"],
                    "from_date": a["date"], "to_date": b["date"],
                    "same_day": False,
                })
        total_reversals += len(reversals)
        last = tl[-1] if tl else None
        first = tl[0] if tl else None
        cons = 1
        if tl:
            for i in range(len(tl) - 1, 0, -1):
                if tl[i]["state"] == tl[i - 1]["state"]:
                    cons += 1
                else:
                    break
        latest_rev = reversals[-1] if reversals else None
        rows.append({
            "code": code,
            "name": stocks.get(code, ""),
            "metric": latest_metrics.get(code),
            "cap": latest_caps.get(code),
            "nK": nK, "nD": nD,
            "first_date": first["date"] if first else None,
            "last_date": last["date"] if last else None,
            "current": last["state"] if last else None,
            "reversals": reversals,
            "n_reversals": len(reversals),
            "latest_reversal": latest_rev,
            "trailing": cons,
            "timeline": tl,
        })

    def sort_key(r):
        lr = r["latest_reversal"]
        lr_date = lr["to_date"] if lr else "0000-00-00"
        return (lr is not None, lr_date, r["last_date"] or "0000-00-00")

    rows.sort(key=sort_key, reverse=True)

    today = days[-1]["date"] if days else None
    today_k = days[-1].get("k", []) if days else []
    today_d = days[-1].get("d", []) if days else []

    return {
        "updated": updated,
        "today": today,
        "today_k": today_k,
        "today_d": today_d,
        "total_stocks": len(rows),
        "total_reversals": total_reversals,
        "stocks_with_reversal": sum(1 for r in rows if r["reversals"]),
        "rows": rows,
    }


# ---------- 观察池（首次D后30天涨跌幅，验证D参考性）----------
def secid_of(code):
    """沪市(6开头，含科创板688) -> 1.x；深市(0/3开头) -> 0.x。"""
    return ("1." + code) if code.startswith("6") else ("0." + code)


def fetch_kline(secid, beg, end):
    """拉前复权日K线，返回 [{date, close}] 或 None（网络失败时）。"""
    url = ("https://push2his.eastmoney.com/api/qt/stock/kline/get"
           "?secid=%s&fields1=f1,f2,f3&fields2=f51,f53&klt=101&fqt=1&beg=%s&end=%s"
           % (secid, beg, end))
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://quote.eastmoney.com/"})
        with urllib.request.urlopen(req, timeout=20) as r:
            b = json.loads(r.read().decode("utf-8", "ignore"))
        data = b.get("data")
        if not data or not data.get("klines"):
            return None
        out = []
        for line in data["klines"]:
            parts = line.split(",")
            out.append({"date": parts[0], "close": float(parts[2])})
        return out
    except Exception as e:
        print("  [价格] 拉取失败 %s: %s" % (secid, e))
        return None


def _close_on_or_before(klines, date):
    best = None
    for k in klines:
        if k["date"] <= date:
            best = k
        else:
            break
    return best


def build_obs(records, config):
    """构建观察池：每只首次出现 D 的个股，计算其 D日收盘至今(<=30天)累计涨幅、同期基准涨幅、超额。"""
    days = records.get("days", [])
    if not days:
        return {"updated": records.get("updated"), "benchmark": BENCH_NAME, "stocks": []}
    today = datetime.date.today()
    first_d = {}
    for day in days:
        for c in day.get("d", []):
            if c not in first_d:
                first_d[c] = day["date"]
    names = load_json(STOCKS_PATH, {})
    cache = {s["code"]: s for s in load_json(OBS_PATH, {}).get("stocks", [])}
    bench_cache = None
    result = []
    for code, entry_date in first_d.items():
        ed = datetime.date.fromisoformat(entry_date)
        holding = (today - ed).days
        exited = holding > OBS_WINDOW
        cached = cache.get(code)
        if cached and cached.get("exited"):
            result.append(cached)  # 已流出则冻结，不再拉取
            continue
        klines = fetch_kline(secid_of(code), entry_date, today.isoformat())
        if klines is None:
            # 网络不可用：沿用缓存或占位，状态标「待同步」
            ec = cached.get("entry_close") if cached else None
            ret = cached.get("ret") if cached else None
            br = cached.get("bench_ret") if cached else None
            result.append({
                "code": code, "name": names.get(code, ""), "entry_date": entry_date,
                "entry_close": ec, "ret": ret, "bench_ret": br,
                "excess": (ret - br) if (ret is not None and br is not None) else None,
                "holding_days": holding, "exited": exited,
                "status": "待同步" if ret is None else ("已流出" if exited else "活跃"),
                "updated": today.isoformat(),
            })
            continue
        ek = _close_on_or_before(klines, entry_date) or klines[0]
        entry_close = ek["close"]
        ret = klines[-1]["close"] / entry_close - 1
        if bench_cache is None:
            bench_cache = fetch_kline(BENCH_SECID, entry_date, today.isoformat())
        bench_ret = None
        if bench_cache:
            b0 = _close_on_or_before(bench_cache, entry_date)
            if b0:
                bench_ret = bench_cache[-1]["close"] / b0["close"] - 1
        result.append({
            "code": code, "name": names.get(code, ""), "entry_date": entry_date,
            "entry_close": entry_close, "ret": ret, "bench_ret": bench_ret,
            "excess": (ret - bench_ret) if bench_ret is not None else None,
            "holding_days": holding, "exited": exited,
            "status": "已流出" if exited else "活跃",
            "updated": today.isoformat(),
        })
    # 按累计涨幅降序（None 垫底）
    result.sort(key=lambda x: (x["ret"] if x["ret"] is not None else -1e9), reverse=True)
    obs = {"updated": today.isoformat(), "benchmark": BENCH_NAME, "stocks": result}
    save_json(OBS_PATH, obs)
    print("  [观察池] 首次D=%d  活跃=%d  已流出=%d  待同步=%d" % (
        len(result), sum(1 for x in result if not x["exited"] and x["status"] != "待同步"),
        sum(1 for x in result if x["exited"]),
        sum(1 for x in result if x["status"] == "待同步")))
    return obs


# ---------- 生成网页 ----------
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DK 变化记录表</title>
<style>
  :root{
    --bg:#f6f7f9; --card:#ffffff; --fg:#1c2024; --muted:#6b7280;
    --line:#e5e7eb; --k:#dc2626; --d:#16a34a; --warn:#b45309; --warnbg:#fffbeb;
    --accent:#2563eb;
  }
  @media (prefers-color-scheme: dark){
    :root{
      --bg:#0f1115; --card:#171a21; --fg:#e8eaed; --muted:#9aa3af;
      --line:#262b33; --k:#f87171; --d:#4ade80; --warn:#fbbf24; --warnbg:#2a2113;
      --accent:#60a5fa;
    }
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--fg);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
    line-height:1.5;padding:16px;max-width:960px;margin:0 auto}
  h1{font-size:20px;margin:0 0 4px}
  .sub{color:var(--muted);font-size:13px;margin-bottom:16px}
  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:18px}
  @media(max-width:640px){.cards{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px}
  .card .n{font-size:22px;font-weight:700}
  .card .l{font-size:12px;color:var(--muted);margin-top:2px}
  .card.k .n{color:var(--k)} .card.d .n{color:var(--d)}
  .warnbox{background:var(--warnbg);border:1px solid var(--warn);border-radius:12px;
    padding:12px 14px;margin-bottom:18px}
  .warnbox h2{font-size:15px;margin:0 0 8px;color:var(--warn)}
  .witem{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:7px 0;
    border-top:1px dashed var(--line);font-size:14px}
  .witem:first-of-type{border-top:none}
  .badge{font-size:11px;padding:1px 7px;border-radius:999px;font-weight:700}
  .bK{background:color-mix(in srgb,var(--k) 18%,transparent);color:var(--k)}
  .bD{background:color-mix(in srgb,var(--d) 18%,transparent);color:var(--d)}
  .bSame{background:var(--warn);color:#fff}
  .arrow{color:var(--muted);font-weight:700}
  .controls{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
  .controls input{flex:1;min-width:160px;padding:8px 10px;border-radius:9px;
    border:1px solid var(--line);background:var(--card);color:var(--fg);font-size:14px}
  .controls button{padding:8px 12px;border-radius:9px;border:1px solid var(--line);
    background:var(--card);color:var(--fg);font-size:13px;cursor:pointer}
  .controls button.on{background:var(--accent);color:#fff;border-color:var(--accent)}
  table{width:100%;border-collapse:collapse;font-size:13px;background:var(--card);
    border:1px solid var(--line);border-radius:12px;overflow:hidden}
  th,td{padding:8px 10px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap}
  th{color:var(--muted);font-weight:600;position:sticky;top:0;background:var(--card)}
  tr:last-child td{border-bottom:none}
  td.wrap{white-space:normal}
  .sK{color:var(--k);font-weight:700} .sD{color:var(--d);font-weight:700}
  .tl{color:var(--muted);font-size:12px}
  .empty{color:var(--muted);padding:20px;text-align:center}
  a{color:var(--accent)}
</style>
</head>
<body>
<h1>DK 变化记录表</h1>
<div class="sub">最后更新：__UPDATED__ ｜ K→D / D→K 反转自动置顶预警</div>

<div class="cards" id="cards"></div>
<div class="warnbox" id="warnbox">
  <h2>⚠ 反转预警（重点提示）</h2>
  <div id="warns"></div>
</div>

<div class="controls">
  <input id="q" placeholder="搜索代码 / 名称…" oninput="render()">
  <button id="f_all" class="on" onclick="setF('all',this)">全部</button>
  <button id="f_rev" onclick="setF('rev',this)">仅反转</button>
  <button id="f_k" onclick="setF('k',this)">当前K</button>
  <button id="f_d" onclick="setF('d',this)">当前D</button>
  <select id="sort" onchange="render()" style="padding:8px 10px;border-radius:9px;border:1px solid var(--line);background:var(--card);color:var(--fg);font-size:13px">
    <option value="cap_desc">市值↓ 大→小</option>
    <option value="cap_asc">市值↑ 小→大</option>
    <option value="rev">反转优先</option>
    <option value="code">代码</option>
  </select>
</div>

<table>
  <thead><tr>
    <th>代码</th><th>名称</th><th>市值</th><th>当前</th><th>涨幅%</th><th>首现</th><th>末次</th>
    <th>K</th><th>D</th><th>反转</th><th>连续</th><th>状态时间线</th>
  </tr></thead>
  <tbody id="tbody"></tbody>
</table>

<script>
const DATA = __DATA__;
let FILTER = "all";

function setF(v,el){
  FILTER=v;
  document.querySelectorAll('.controls button').forEach(b=>b.classList.remove('on'));
  el.classList.add('on');
  render();
}
function esc(s){return (s==null?'':String(s));}
function fmtMetric(m){
  if(m==null) return '-';
  const cls = m>0?'sK':(m<0?'sD':'');
  const sign = m>0?'+':'';
  return '<span class="'+cls+'">'+sign+(+m).toFixed(2)+'</span>';
}
function fmtCap(c){ if(c==null) return '-'; if(c>=10000) return (c/10000).toFixed(2)+'万亿'; return c.toFixed(2)+'亿'; }
function sortRows(rows){
  const v=document.getElementById('sort').value; const a=[...rows];
  if(v==='cap_desc') a.sort((x,y)=>(y.cap||0)-(x.cap||0));
  else if(v==='cap_asc') a.sort((x,y)=>(x.cap||0)-(y.cap||0));
  else if(v==='rev') a.sort((x,y)=>((y.latest_reversal?1:0)-(x.latest_reversal?1:0))||((y.last_date||'')<(x.last_date||'')?1:-1));
  else a.sort((x,y)=>x.code<y.code?-1:(x.code>y.code?1:0));
  return a;
}
function stateBadge(s){
  if(s==='K') return '<span class="badge bK">K</span>';
  if(s==='D') return '<span class="badge bD">D</span>';
  return '';
}
function tlText(tl){
  return tl.map(e=>{
    const m=e.date.slice(5);
    return (e.state==='K'?'<span class=sK>K</span>':'<span class=sD>D</span>')+'<span class=tl>'+m+'</span>';
  }).join(' → ');
}
function render(){
  // cards
  const c=document.getElementById('cards');
  c.innerHTML=
    card(DATA.total_stocks,'跟踪个股')+
    card(DATA.today_k.length,'今日 K', 'k')+
    card(DATA.today_d.length,'今日 D', 'd')+
    card(DATA.stocks_with_reversal,'反转个股');
  // warns
  const warns=DATA.rows.filter(r=>r.latest_reversal);
  const w=document.getElementById('warns');
  if(!warns.length){w.innerHTML='<div class="empty">暂无反转记录</div>';}
  else{
    w.innerHTML=warns.map(r=>{
      const lr=r.latest_reversal;
      const sd=lr.same_day?' <span class="badge bSame">同日</span>':'';
      const dir=(lr.from==='K'?'K→D':'D→K');
      return '<div class="witem">'+stateBadge(lr.from)+'<span class="arrow">→</span>'+
        stateBadge(lr.to)+sd+' '+
        '<b>'+esc(r.code)+'</b> '+esc(r.name)+' '+
        '<span class="tl">'+lr.from_date+' → '+lr.to_date+'</span>'+
        '<span class="tl">（累计反转 '+r.n_reversals+' 次）</span></div>';
    }).join('');
  }
  // table
  const q=document.getElementById('q').value.trim().toLowerCase();
  const tb=document.getElementById('tbody');
  const rows=DATA.rows.filter(r=>{
    if(FILTER==='rev' && r.reversals.length===0) return false;
    if(FILTER==='k' && r.current!=='K') return false;
    if(FILTER==='d' && r.current!=='D') return false;
    if(q && !(r.code.toLowerCase().includes(q)||(r.name||'').toLowerCase().includes(q))) return false;
    return true;
  });
  if(!rows.length){tb.innerHTML='<tr><td colspan="12" class="empty">无匹配</td></tr>';return;}
  tb.innerHTML=sortRows(rows).map(r=>{
    const cur=r.current==='K'?'<span class=sK>K</span>':(r.current==='D'?'<span class=sD>D</span>':'-');
    const rev=r.latest_reversal?('<b style="color:var(--warn)">'+(r.latest_reversal.from==='K'?'K→D':'D→K')+'</b>'):'-';
    return '<tr><td>'+esc(r.code)+'</td><td>'+esc(r.name)+'</td><td>'+fmtCap(r.cap)+'</td><td>'+cur+'</td><td>'+fmtMetric(r.metric)+'</td>'+
      '<td class="tl">'+esc(r.first_date)+'</td><td class="tl">'+esc(r.last_date)+'</td>'+
      '<td>'+r.nK+'</td><td>'+r.nD+'</td><td>'+rev+'</td><td>'+r.trailing+'</td>'+
      '<td class="wrap tl">'+tlText(r.timeline)+'</td></tr>';
  }).join('');
}
function card(n,l,cls){return '<div class="card '+(cls||'')+'"><div class="n">'+n+'</div><div class="l">'+l+'</div></div>';}
render();
</script>
</body>
</html>
"""


def render_html(analysis, config):
    return HTML_TEMPLATE.replace("__UPDATED__", analysis.get("updated") or "—").replace(
        "__DATA__", json.dumps(analysis, ensure_ascii=False)
    )


def write_html(html):
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print("  [生成] %s (%d 字节)" % (HTML_PATH, len(html)))


def generate_app(records, stocks, obs, config):
    """生成手机端上传应用（含观察池）。
    为彻底规避 GitHub 连接器对超长行/内容的截断与「手工转义引号」出错风险：
      - 种子数据 base64 切片为 seed_p1..N.js + seed_load.js
      - 整个 app HTML 也 base64 切片为 app_p1..M.js + app_load.js（app_load 用 document.write 注入）
      - 仓库里的 index.html 只做一个「壳」：顺序加载上述切片，自身几乎无引号、可安全推送
    所有切片文件均为 quote-light（base64），便于通过连接器逐文件可靠推送。"""
    import base64
    CHUNK = 5000

    def chunk_write(b64, prefix, var):
        parts = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]
        for i, ch in enumerate(parts):
            line = ('window.%s = "%s";\n' % (var, ch)) if i == 0 \
                else ('window.%s += "%s";\n' % (var, ch))
            with open(os.path.join(BASE, "%s%d.js" % (prefix, i + 1)), "w", encoding="utf-8") as f:
                f.write(line)
        return len(parts)

    # 1) 种子数据切片
    seed = {
        "updated": records.get("updated"),
        "days": records.get("days", []),
        "names": stocks,
        "obs": obs,
    }
    seed_b64 = base64.b64encode(json.dumps(seed, ensure_ascii=False).encode("utf-8")).decode("ascii")
    n = chunk_write(seed_b64, "seed_p", "__SEED_B64")
    with open(os.path.join(BASE, "seed_load.js"), "w", encoding="utf-8") as f:
        f.write('window.SEED_DATA = JSON.parse(new TextDecoder().decode('
                'Uint8Array.from(atob(window.__SEED_B64), c => c.charCodeAt(0))));\n')

    # 2) 整个 app HTML 切片（统一用绝对路径，根与 public 共用一套切片）
    tpl = open(APP_TEMPLATE_PATH, "r", encoding="utf-8").read()
    seed_tags = "".join('<script src="/seed_p%d.js"></script>' % (i + 1) for i in range(n)) \
        + '<script src="/seed_load.js"></script>'
    app_html = tpl.replace("__SEED__", "").replace("<!--SEED_SCRIPTS-->", seed_tags)
    app_b64 = base64.b64encode(app_html.encode("utf-8")).decode("ascii")
    m = chunk_write(app_b64, "app_p", "__APP_B64")
    with open(os.path.join(BASE, "app_load.js"), "w", encoding="utf-8") as f:
        f.write('document.open();document.write(new TextDecoder().decode('
                'Uint8Array.from(atob(window.__APP_B64), c => c.charCodeAt(0))));'
                'document.close();\n')

    # 3) 壳 index.html（quote-light）：只按顺序加载切片
    def shell():
        s = '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n' \
            '<title>DK 变化记录表</title>\n</head>\n<body>\n'
        for i in range(n):
            s += '<script src="/seed_p%d.js"></script>\n' % (i + 1)
        s += '<script src="/seed_load.js"></script>\n'
        for i in range(m):
            s += '<script src="/app_p%d.js"></script>\n' % (i + 1)
        s += '<script src="/app_load.js"></script>\n</body>\n</html>\n'
        return s
    for target in (os.path.join(BASE, "index.html"), os.path.join(BASE, "public", "index.html")):
        with open(target, "w", encoding="utf-8") as f:
            f.write(shell())

    # 清理旧的单文件
    for old in (os.path.join(BASE, "seed.js"), os.path.join(BASE, "public", "seed.js")):
        if os.path.exists(old):
            os.remove(old)
    print("  [生成] 壳 index.html + seed 切片 %d 个 + app 切片 %d 个 + 两个 loader" % (n, m))


def print_summary(analysis):
    print(
        "  [汇总] 更新=%s 跟踪=%d 今日K=%d 今日D=%d 反转个股=%d 累计反转=%d"
        % (
            analysis["updated"], analysis["total_stocks"], len(analysis["today_k"]),
            len(analysis["today_d"]), analysis["stocks_with_reversal"], analysis["total_reversals"],
        )
    )


# ---------- 自检 ----------
def run_selftest(config):
    stocks = {
        "600000": "浦发银行", "600001": "邯郸钢铁", "600002": "齐鲁石化",
        "600009": "上海机场", "600011": "华能国际",
    }
    days = [
        {"date": "2026-08-10", "k": ["600000", "600001", "600002"], "d": ["600009"], "metrics": {}},
        {"date": "2026-08-11", "k": ["600000", "600001"], "d": ["600002", "600009"], "metrics": {}},
        {"date": "2026-08-12", "k": ["600001"], "d": ["600000"], "metrics": {}},
        {"date": "2026-08-13", "k": ["600000"], "d": ["600001"], "metrics": {}},
        {"date": "2026-08-14", "k": [], "d": ["600000"], "metrics": {}},
    ]
    records = {"updated": "2026-08-14", "days": days}
    analysis = build_analysis(records, stocks, config)
    html = render_html(analysis, config)
    out = os.path.join(BASE, "_selftest.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    # 文本校验
    print("== 自检 ==")
    print_summary(analysis)
    for r in analysis["rows"]:
        if r["reversals"]:
            seq = " ".join(e["state"] for e in r["timeline"])
            print("  %s %s : %s  反转%d次 最近:%s->%s" % (
                r["code"], r["name"], seq, r["n_reversals"],
                r["latest_reversal"]["from"], r["latest_reversal"]["to"]))
    print("  自检网页已写: %s" % out)
    # 断言关键反转
    bycode = {r["code"]: r for r in analysis["rows"]}
    assert bycode["600002"]["n_reversals"] == 1, "600002 应 1 次反转 K->D"
    assert bycode["600000"]["n_reversals"] == 3, "600000 应 3 次反转"
    assert bycode["600001"]["n_reversals"] == 1, "600001 应 1 次反转 K->D"
    print("  断言通过 ✅")


# ---------- 入口 ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--k")
    ap.add_argument("--d")
    ap.add_argument("--date")
    ap.add_argument("--config", default=CONFIG_PATH)
    ap.add_argument("--regen", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()

    config = load_config(args.config)

    if args.selftest:
        run_selftest(config)
        return

    if args.regen:
        records = load_json(RECORDS_PATH, {"updated": None, "days": []})
        stocks = load_json(STOCKS_PATH, {})
        analysis = build_analysis(records, stocks, config)
        obs = build_obs(records, config)
        generate_app(records, stocks, obs, config)
        print_summary(analysis)
        return

    if args.k and args.d:
        date = args.date or parse_date_from_filename(args.k, config) or parse_date_from_filename(args.d, config)
        if not date:
            raise SystemExit("[错误] 文件名无日期且未用 --date 指定，无法定位当天。")
        process_day(args.k, args.d, date, config)
        records = load_json(RECORDS_PATH, {"updated": None, "days": []})
        stocks = load_json(STOCKS_PATH, {})
        analysis = build_analysis(records, stocks, config)
        obs = build_obs(records, config)
        generate_app(records, stocks, obs, config)
        print_summary(analysis)
        return

    raise SystemExit("用法:\n  python process.py --k X_K.xlsx --d X_D.xlsx [--date 2026-08-14]\n  python process.py --regen\n  python process.py --selftest")


if __name__ == "__main__":
    main()
