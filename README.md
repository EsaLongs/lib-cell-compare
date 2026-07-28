# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

默认**无需** `--format-filter`：直接从完整 cell 名提取逻辑功能根作为 A 列归类 key（如 `XOR2D1COT` → `XOR2`，`ND2X4APBCOTC` → `ND2`，`FILL12VGCOT` → `FILL12`）。工艺/优化段（`COT` / `OPT` / `EEQ*` / `SK*` / `TWA` 等）以及驱动强度、版图后缀会在提取时剥掉；尺寸数字保留。

```bash
python3 compare_cells.py \
  --np-file preview/NP1PP.list \
  --c1-file preview/C1Y.list \
  --output preview/NP1PP_vs_C1Y_preview.xlsx
```

### 可选删除（`--format-filter`）

在功能根提取**之前**执行；每个 token 用双引号包起来，空格分隔，按从左到右顺序剔除：

```bash
python3 compare_cells.py --format-filter '"FOO" "BAR"'
```

### 可选替换（`--format-replace`）

在删除之后、功能根提取之前执行；格式为 `"原字符串 : 新字符串"`（英文冒号 `:`）：

```bash
python3 compare_cells.py \
  --format-replace '"X(\d{1,2}) : D\1"'
```

原/新两侧都支持正则；分隔符必须是英文 `:`（两侧可有空格）。

可选路径参数：

```bash
python3 compare_cells.py \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

```bash
python3 compare_cells.py \
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
- A 列 = 从完整原名提取的功能根（可选先 filter / replace）
- 复合单元与基础门分开（如 `XOR2` ≠ `XOR2AOI22`；`XOR2` ≠ `XOR3`）
- 写出三列 Excel：`function_root | NP1PP | C1Y`；A 列按 key 合并
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行，不再按序号硬配
- Excel：字号 16、首行/首列加粗、B=C 非空浅绿、首行冻结（首行无下边框）、分组上下边框横跨 A–C
