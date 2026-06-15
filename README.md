# finance-excel-flow

一个可配置的 Excel 自动化处理项目。

适合这样的场景：
- 每个月收到一份源 Excel
- 需要按固定业务规则做清洗、筛选、计算、列重排
- 最终导出一份新的目标 Excel

你后续只需要替换源文件，并根据业务需求调整配置文件，不必再手工一列一列处理。

## 功能

- 读取 Excel 源文件
- 按配置重命名列
- 校验必填列
- 逐行执行筛选条件
- 逐行计算新列
- 指定导出列顺序
- 自动生成输出文件名
- 自动调整列宽和冻结首行
- 基于固定 Excel 模板复制导出，保留样式、合并单元格和版式

## 安装

建议先创建虚拟环境，再安装依赖：

```bash
pip install -r requirements.txt
```

如果你想本地安装成命令行工具，也可以：

```bash
pip install -e .
```

## 使用

准备好：
- 一个源 Excel 文件，例如 `inputs/source.xlsx`
- 一个规则配置文件，例如 `configs/example.yml`

然后执行：

```bash
finance-excel-flow --input inputs/source.xlsx --config configs/example.yml --output-dir output
```

或者：

```bash
python -m finance_excel_flow --input inputs/source.xlsx --config configs/example.yml --output-dir output
```

执行后会在 `output/` 下生成目标文件。

如果你要导出这份 KA 模板，可以直接在配置里指定模板路径：

```yaml
output:
  template_path: templates/ka_promo_and_expense_template.xlsx
  sheet_name: Sheet1
business_rules:
  - name: ace_ego_promo_accrual_1
    source_filters:
      - 'text("Customer number") == "0376"'
      - 'num("Due Amount") == 20'
      - 'text("Brand") == "000030"'
    sum_field: "Due Amount"
    percentage: 0.11
    name_template: "ACE EGO {source_yyyymm} Promo Accrual"
    template_cells:
      E5: 'lit("0376")'
      F5: 'lit("000030")'
      H5: 'round(accrual_amount, 2)'
      J5: 'name'
      L5: 'round(due_amount_total, 2)'
      M5: 'percentage'
```

这种模式下，程序会先复制模板，再按规则筛选明细、汇总金额并回填模板中的占位字段。`source_yyyymm` 会优先从源文件名里提取，比如 `Sales report_NA_202605.xlsx` 会得到 `202605`。

## 配置示例

见 [`configs/example.yml`](./configs/example.yml) 和 [`configs/template-example.yml`](./configs/template-example.yml)。

## 规则表达式写法

配置里的表达式按“每一行数据”执行，常用 helper：

- `col("列名")`：读取原始值
- `text("列名")`：读取文本
- `num("列名")`：读取数值，自动转成 `Decimal`
- `date("列名")`：解析日期
- `ifnull(a, b)`：空值兜底

示例：

```yaml
calculated_columns:
  毛利: 'num("销售额") - num("成本")'
filters:
  - 'num("销售额") > 0'
```

## 你后续如何扩展

如果你的业务规则更复杂，我们可以继续把它扩展成：
- 多 sheet 处理
- 汇总表 / 分组聚合
- 多个输出文件
- 模板 Excel 保留样式输出
- 自动化批量处理一个文件夹里的多份源文件
