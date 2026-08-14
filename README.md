# DK 变化登记与反转预警看板

逐日登记出现 K 点 / D 点的个股，重点预警两类**反转事件**：

- `K → D`：先出现 K，之后某天出现 D
- `D → K`：先出现 D，之后某天出现 K

## 使用方法
1. 浏览器打开 `https://seonkoo.github.io/dk-tracker/`
2. 点「K表」「D表」选择当天的 Excel 文件（支持精简清单与全市场「选股」导出两种格式）
3. 点「解析并合并」→ 自动追加到历史、刷新看板
4. 反转信号（K→D / D→K）在顶部 ⚠ 区高亮

历史保存在浏览器本地，可用「导出历史 / 导入历史」备份迁移。

## 本地引擎（可选）
- `scripts/process.py`：读两份 xlsx → 合并进 `data/records.json` → 检测反转 → 生成 `index.html`
- `config.json`：列名映射、文件命名约定
- 用法：`python process.py --k X_K.xlsx --d X_D.xlsx --date YYYY-MM-DD`
