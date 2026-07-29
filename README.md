# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

只需一份 **`KEY<TAB>中文`** 表（如 `preview/function_key_zh.txt`）：

- 第一列英文 KEY → 匹配顺序（先命中生效）
- 第二列中文 → Excel A 列显示（直接改文件即可）

```bash
python3 compare_cells.py \
  --function-key-zh-file preview/function_key_zh.txt \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

| 参数 | 说明 |
|------|------|
| `--function-key-zh-file` | 必填，KEY↔中文表 |
| `--np-file` | NP1PP list（有默认路径） |
| `--c1-file` | C1Y list（有默认路径） |
| `--output` | 输出 Excel（有默认路径） |

### 匹配规则

- **先命中生效**：已匹配项不再被后续 key 改写
- KEY 后紧跟数字则不算命中（`OAI22` 不会误吃 `OAI2211`）
- 更长字母续接形态仍需排在前面：如 `OAI22OAI21` 在 `OAI22` 前

## 依赖

```bash
pip install -r requirements.txt
```

## 说明

- 读取两个 list 文件，跳过空行与 `rg:` 开头行
- A 列 = KEY + 第二列中文；未命中则在 stderr 打印完整列表（`source<TAB>cell`，来源为 NP1PP / C1Y / NP1PP+C1Y），便于复制排查
- 写出三列 Excel：`function_key | NP1PP | C1Y`；A 列按 key 合并
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行
- Excel：字号 16、首行/首列加粗、B=C 非空浅绿、首行冻结、分组边框
