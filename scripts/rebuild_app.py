#!/usr/bin/env python3
"""仅重建 app 切片（app_p*.js / app_load.js / index.html），不重抓任何数据。

用途：只改了 app_template.html 的样式/结构，想重新烧录到静态站点，但保留
已烘焙好的种子数据（seed_p*.js / seed_load.js / data/*）。直接复用源码里
generate_app 的切片逻辑，但跳过 fetch_kline / earnings-radar / 智谱等联网步骤。
"""
import os, re, base64, glob

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_TEMPLATE_PATH = os.path.join(BASE, "app_template.html")
CHUNK = 5000


def chunk_write(b64, prefix, var):
    parts = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]
    for i, ch in enumerate(parts):
        line = ('window.%s = "%s";\n' % (var, ch)) if i == 0 \
            else ('window.%s += "%s";\n' % (var, ch))
        with open(os.path.join(BASE, "%s%d.js" % (prefix, i + 1)), "w", encoding="utf-8") as f:
            f.write(line)
    return len(parts)


def main():
    # 现有种子切片数量（决定要注入的 <script> 标签）
    seed_js = sorted(
        glob.glob(os.path.join(BASE, "seed_p*.js")),
        key=lambda p: int(re.search(r"_p(\d+)\.js$", p).group(1)),
    )
    n = len(seed_js)
    seed_tags = "".join('<script src="seed_p%d.js"></script>' % (i + 1) for i in range(n)) \
        + '<script src="seed_load.js"></script>'

    tpl = open(APP_TEMPLATE_PATH, "r", encoding="utf-8").read()
    app_html = tpl.replace("__SEED__", "").replace("<!--SEED_SCRIPTS-->", seed_tags)
    # 一致性校验：切片里必须包含模板特征串
    assert "DK 变化记录表" in app_html, "模板拼接异常"

    app_b64 = base64.b64encode(app_html.encode("utf-8")).decode("ascii")
    m = chunk_write(app_b64, "app_p", "__APP_B64")
    with open(os.path.join(BASE, "app_load.js"), "w", encoding="utf-8") as f:
        f.write('document.open();document.write(new TextDecoder().decode('
                'Uint8Array.from(atob(window.__APP_B64), c => c.charCodeAt(0))));'
                'document.close();\n')

    def shell(prefix=""):
        s = '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n' \
            '<title>DK 变化记录表</title>\n</head>\n<body>\n'
        for i in range(n):
            s += '<script src="%sseed_p%d.js"></script>\n' % (prefix, i + 1)
        s += '<script src="%sseed_load.js"></script>\n' % prefix
        for i in range(m):
            s += '<script src="%sapp_p%d.js"></script>\n' % (prefix, i + 1)
        s += '<script src="%sapp_load.js"></script>\n</body>\n</html>\n' % prefix
        return s

    for target, prefix in ((os.path.join(BASE, "index.html"), ""),
                           (os.path.join(BASE, "public", "index.html"), "../")):
        with open(target, "w", encoding="utf-8") as f:
            f.write(shell(prefix))

    print("  [重建] app 切片 %d 个 + app_load.js + index.html(根/public)，种子 %d 个未动" % (m, n))


if __name__ == "__main__":
    main()
