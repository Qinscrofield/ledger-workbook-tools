# 台账工作簿工具集 / Ledger Workbook Tools

一组用 openpyxl 处理工程财务台账 Excel 的脚本。核心设计原则只有一条：

> **改格式，不改数据。** 每个会写文件的脚本都先对工作簿做快照（值、公式、合并单元格），
> 处理完再比对一遍，确认业务数据原样保留，只动样式和结构，并输出校验报告 JSON。

## 脚本

| 脚本 | 作用 |
| --- | --- |
| `inspect_ledger_workbook.py` | 只读检查：输出每个工作表的使用范围、非空单元格数、疑似表头行（按「日期/金额/备注…」等关键词打分）和前几行预览，先看清表长什么样 |
| `format_ledger_workbook.py` | 规范化排版：统一字体、边框、数字/日期格式、列宽、冻结窗格、打印设置；快照校验数据未变 |
| `fix_wps_view_ledger.py` | 修复 WPS 打开视图：清除冻结窗格和残留选区，让文件打开时回到 A1；快照校验数据未变 |
| `finance_standardize_ledger.py` | 把「一表多栏」的宽台账（每个往来对象占几列）拆解成一张规范流水账：序号、业务大类、明细分类、日期、往来对象、金额、税金、备注，并用公式回链原表单元格 |

## 用法

```bash
pip install openpyxl

python inspect_ledger_workbook.py 你的台账.xlsx
python format_ledger_workbook.py 你的台账.xlsx 输出目录
python fix_wps_view_ledger.py 已排版.xlsx
python finance_standardize_ledger.py 已排版_WPS兼容版.xlsx 输出目录
```

## 注意

- `format_ledger_workbook.py` 中的行列范围（84 行 × 52 列、表头行位置）是按示例台账的结构写的，换一张表需要先用 inspect 看清结构再调整常量。
- `finance_standardize_ledger.py` 里的 `MODULES` 是虚构示例映射（列号、往来对象名称均为示例），使用时按自己的宽表分栏替换。
- 所有脚本都不会修改输入文件，输出写到新文件/输出目录。
