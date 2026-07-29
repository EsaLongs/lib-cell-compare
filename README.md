# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

只需一份 **`KEY<TAB>中文`** 表（如 `preview/function_key_zh.txt`）：

- 第一列英文 KEY → 匹配顺序（先命中生效）
- 第二列中文 → Excel A 列显示（直接改文件即可）

```text
# key	chinese
BUF	缓冲器
INV	反相器
AOI21	与或非门
```

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_key_zh.txt \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

`--function-key-zh-file` 与上面等价，二选一即可，不必同时传两个文件。

### 匹配规则

- **先命中生效**：已匹配项不再被后续 key 改写
- KEY 后紧跟数字则不算命中（`OAI22` 不会误吃 `OAI2211`；`FILL1` 不会误吃 `FILL12`）
- 更长字母续接形态仍需排在前面：如 `OAI22OAI21` 在 `OAI22` 前
- 也可命令行临时给 key：`--function-keys OAI2211 OAI22 AN2`

### 生成建议表（只写一个文件）

```bash
python3 compare_cells.py \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --dump-suggested-keys preview/function_key_zh.txt
```

生成后改第二列中文，再带 `--function-keys-file` 跑比较即可。

### 可选删除 / 替换

在前缀匹配前执行（一般不必再用）：

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_key_zh.txt \
  --format-filter '"FOO"' \
  --format-replace '"X(\d{1,2}) : D\1"'
```

## Preview

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_key_zh.txt \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

## 依赖

```bash
pip install -r requirements.txt
```

## 说明

- 读取两个 list 文件，跳过空行与 `rg:` 开头行
- A 列 = 表中 KEY 的首次合格前缀命中 + 第二列中文；未命中则保留匹配前字符串并告警
- 写出三列 Excel：`function_key | NP1PP | C1Y`；A 列按 key 合并
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行
- Excel：字号 16、首行/首列加粗、B=C 非空浅绿、首行冻结（首行无下边框）、分组上下边框横跨 A–C
