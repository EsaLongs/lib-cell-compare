# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

```bash
python3 compare_cells.py --format-filter TOKEN1 TOKEN2 ...
```

示例：

```bash
python3 compare_cells.py --format-filter EEQMBD EEQMBC OPT
```

可选参数：

```bash
python3 compare_cells.py --format-filter EEQMBD EEQMBC OPT \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

仓库内 `preview/` 为部分 list 样本。生成预览 Excel：

```bash
python3 compare_cells.py --format-filter EEQMBD EEQMBC OPT \
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
- 按 `--format-filter` token 做精确字符串剔除（非正则），再按 display key 分组
- 写出三列 Excel：`display_key | NP1PP | C1Y`，并对相同 display key 合并 A 列
