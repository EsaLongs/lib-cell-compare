# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

```bash
python3 compare_cells.py --format-filter TOKEN1 TOKEN2 ...
```

精确剔除 + 正则剔除示例：

```bash
python3 compare_cells.py \
  --format-filter EEQMBD EEQMBC OPT \
  --format-filter-regex 'A.{2}$'
```

可选参数：

```bash
python3 compare_cells.py \
  --format-filter EEQMBD EEQMBC OPT \
  --format-filter-regex 'A.{2}$' \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

仓库内 `preview/` 为部分 list 样本。生成预览 Excel：

```bash
python3 compare_cells.py \
  --format-filter EEQMBD EEQMBC OPT \
  --format-filter-regex 'A.{2}$' \
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
- 以 `COT` 前的前缀作为 base name
- `--format-filter`：精确字符串剔除（非正则）
- `--format-filter-regex`：对 display key 做 `re.sub` 剔除（如末尾 `A.{2}$`）；每个正则会反复应用直到不再匹配
- 先精确剔除，再按参数顺序应用正则
- 写出三列 Excel：`display_key | NP1PP | C1Y`，并对相同 display key 合并 A 列
