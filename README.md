# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

### 删除（`--format-filter`）

每个 token 用双引号包起来，空格分隔，**按从左到右顺序**依次剔除：

```bash
python3 compare_cells.py --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"'
```

### 替换（`--format-replace`）

在删除之后执行；格式为 `"原字符串 : 新字符串"`（英文冒号 `:`）：

```bash
python3 compare_cells.py \
  --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"' \
  --format-replace '"X(\d{1,2}) : D\1"'
```

原/新两侧都支持正则；分隔符必须是英文 `:`（两侧可有空格）。

可选路径参数：

```bash
python3 compare_cells.py \
  --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"' \
  --format-replace '"FOO : BAR"' \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

```bash
python3 compare_cells.py \
  --format-filter '"COT.*" "EEQMBD" "EEQMBC" "OPT" "A.{2}$"' \
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
- 对完整 cell 名先 `--format-filter` 删除，再 `--format-replace` 替换
- token 用 `"` 引用并用空格分界（`shlex` 解析）
- 删除：每个正则反复剔除直到不再匹配；替换：每个规则按顺序 `re.sub` 一次
- 写出三列 Excel：`display_key | NP1PP | C1Y`；A 列按 display key 合并
- 组内仅全名完全相同才左右同行；NP-only / C1-only 各占一行，不再按序号硬配
- Excel：字号 16、首行/首列加粗、B=C 非空浅绿、首行冻结（首行无下边框）、分组上下边框横跨 A–C
