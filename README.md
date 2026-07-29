# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

A 列由有序 **`--function-keys`** 决定：按给定顺序做前缀匹配，**先命中者生效**，已匹配的名字不再被后面的 key 改写。

规则要点：
- **先命中生效**：已匹配项不再被后续 key 改写
- 匹配后若下一位是数字则不算命中（`OAI22` 不会误吃 `OAI221` / `OAI2211` / `OAI2222`；`FILL1` 不会误吃 `FILL12`）
- 对非数字续接的更长形态，仍需把长的写在前面：`--function-keys OAI22OAI21 OAI22`、`XOR2AOI22 XOR2`
- 若只写 `XOR2`、不写 `XOR2AOI22`，则 `XOR2AOI22…` 会归到 `XOR2`（便于主动简化）

```bash
python3 compare_cells.py \
  --function-keys OAI2211 OAI22 AIOI21 AN2 AN3 AN4 \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

长列表可放文件（一行一个 key，`#` 开头为注释）：

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_keys.txt \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

CLI 与文件可同时给：先匹配 `--function-keys`，再匹配文件中的 key。

### 生成建议 key 列表

```bash
python3 compare_cells.py \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --dump-suggested-keys preview/function_keys.txt
```

按「长 key 优先」写出建议列表（可再手工删减/调整顺序）。尺寸类（`FILL12`、`DCAP10`）保留数字；复合门默认与基础门分开列出。

### 可选删除 / 替换

在前缀匹配前执行（一般不必再用）：

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_keys.txt \
  --format-filter '"FOO"' \
  --format-replace '"X(\d{1,2}) : D\1"'
```

## Preview

```bash
python3 compare_cells.py \
  --function-keys-file preview/function_keys.txt \
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
- A 列 = 有序 function key 的首次合格前缀命中，合并单元格内附中文简述（如 `AN2` + 与门）；未命中则保留匹配前字符串并告警
- 写出三列 Excel：`function_key | NP1PP | C1Y`；A 列按 key 合并
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行
- Excel：字号 16、首行/首列加粗、B=C 非空浅绿、首行冻结（首行无下边框）、分组上下边框横跨 A–C
