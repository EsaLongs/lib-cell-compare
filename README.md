# lib-cell-compare

比较两个工艺库（NP1PP / C1Y）的 cell 列表差异，并输出 Excel。

## 用法

`--format-filter` 只接受正则；多个 token 用 `""` 分隔，**按从左到右顺序**依次剔除：

```bash
python3 compare_cells.py --format-filter 'COT.*""EEQMBD""EEQMBC""OPT""A.{2}$'
```

含义：先删 `COT` 及之后内容，再删 `EEQMBD` / `EEQMBC` / `OPT`，最后删末尾 `A??`。

可选参数：

```bash
python3 compare_cells.py \
  --format-filter 'COT.*""EEQMBD""EEQMBC""OPT""A.{2}$' \
  --np-file /path/to/NP1PP.list \
  --c1-file /path/to/C1Y.list \
  --output /path/to/out.xlsx
```

## Preview

```bash
python3 compare_cells.py \
  --format-filter 'COT.*""EEQMBD""EEQMBC""OPT""A.{2}$' \
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
- 对完整 cell 名应用 `--format-filter`（不再硬编码截断 `COT`）
- token 按 `""` 拆分，按顺序 `re.sub`；每个正则反复剔除直到不再匹配，再进入下一个
- 写出三列 Excel：`display_key | NP1PP | C1Y`，并对相同 display key 合并 A 列
