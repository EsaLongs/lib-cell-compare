# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

`--format-filter` 只接受正则；多个 token 用 `""` 分隔：

```bash
python3 compare_cells.py --format-filter 'EEQMBD""EEQMBC""OPT""A.{2}$'
```

可选参数：

```bash
python3 compare_cells.py \
  --format-filter 'EEQMBD""EEQMBC""OPT""A.{2}$' \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

```bash
python3 compare_cells.py \
  --format-filter 'EEQMBD""EEQMBC""OPT""A.{2}$' \
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
- `--format-filter`：按 `""` 拆成多个正则 token，对 display key 做 `re.sub`
- 每个正则按从左到右顺序应用，并反复剔除直到不再匹配
- 写出三列 Excel：`display_key | NP1PP | C1Y`，并对相同 display key 合并 A 列
