#!/usr/bin/env python3
"""仅重生成 AI 文字总结并打补丁进 seed 切片（不重抓 K 线、不重建 app）。

用于修正 AI 情绪结论与硬数据不一致：把观察池实况(胜率/涨跌家数/中位涨幅)
作为硬依据喂给智谱，并要求以价格动作为准下结论。
"""
import os, sys, glob, re, base64, json

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
import process as P

CHUNK = 5000

def decode_seed():
    parts = sorted(glob.glob(os.path.join(BASE, "seed_p*.js")),
                   key=lambda p: int(re.search(r"_p(\d+)\.js", p).group(1)))
    b = "".join(re.search(r'"(.*)"', open(p, encoding="utf-8").read()).group(1) for p in parts)
    return json.loads(base64.b64decode(b).decode("utf-8"))

def write_seed(seed):
    seed_b64 = base64.b64encode(json.dumps(seed, ensure_ascii=False).encode("utf-8")).decode("ascii")
    parts = [seed_b64[i:i + CHUNK] for i in range(0, len(seed_b64), CHUNK)]
    for i, ch in enumerate(parts):
        with open(os.path.join(BASE, "seed_p%d.js" % (i + 1)), "w", encoding="utf-8") as f:
            f.write(('window.__SEED_B64 = "%s";\n' % ch) if i == 0
                    else ('window.__SEED_B64 += "%s";\n' % ch))
    with open(os.path.join(BASE, "seed_load.js"), "w", encoding="utf-8") as f:
        f.write('window.SEED_DATA = JSON.parse(new TextDecoder().decode('
                'Uint8Array.from(atob(window.__SEED_B64), c => c.charCodeAt(0))));\n')
    print("  [补丁] 重写 seed 切片 %d 个 + seed_load.js" % len(parts))

def main():
    config = P.load_config(P.CONFIG_PATH)
    seed = decode_seed()
    flow = seed.get("flow", {})
    obs = seed.get("obs")
    er = seed.get("er")
    if not er or not er.get("available"):
        er = P.fetch_earnings_radar()
    print("  [补丁] 调用智谱重生成总结（含观察池实况）...")
    new_ai = P.fetch_ai_summary(flow, er, config, obs)
    if not new_ai:
        print("  [补丁] 智谱未返回，放弃打补丁")
        return
    flow["ai_summary"] = new_ai
    seed["flow"] = flow
    if er:
        seed["er"] = er
    write_seed(seed)
    print("  [补丁] 新总结:\n" + new_ai.get("text", ""))

if __name__ == "__main__":
    main()
