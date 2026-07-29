# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

只需一份 **`KEY<TAB>中文`** 表（如 `preview/function_key_zh.txt`）：

- 第一列英文 KEY → 匹配顺序（先命中生效）
- 第二列中文 → Excel **A 列**（纯中文，同中文合并）及 **B 列**（KEY+中文）

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
- Excel 四列：`中文 | KEY(+中文) | NP1PP | C1Y`
  - A 列：纯中文，按相同中文合并
  - B 列：英文 KEY + 中文，按 KEY 合并（原 A 列逻辑）
  - C/D 列：库 cell 名，不按中文合并
- 行顺序：先按中文，再按 KEY
- 未命中则在 stderr 打印完整列表（`source<TAB>cell`）
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行
- Excel：字号 16、加粗、C=D 非空浅绿、首行冻结、分组边框
