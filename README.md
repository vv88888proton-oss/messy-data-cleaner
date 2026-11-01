# 05 脏数据清洗命令行工具(CSV / Excel)

## 解决什么问题
常见需求:“这份表是好几个人手填 / 从不同系统导出的,格式乱七八糟,帮我整理干净。”
一条命令读入 CSV 或 Excel,输出干净的数据和一份说明“改了什么”的清洗报告:

| 处理项 | 说明 |
|---|---|
| 列名标准化 | 去首尾/多余空格、全角括号转半角、中英文别名统一(`订单编号` / `Order No` → `order_id`),可用 `--column-map` 追加映射 |
| 空值处理 | `""`、`N/A`、`null`、`-`、`无`、`未知` 等统一为空;删除全空行;缺少必填列(默认 `order_id`)的行删除;金额空值可选不填 / 填 0 / 填中位数 |
| 日期统一 | `2026/3/5`、`2026.03.05`、`20260305`、`2026年3月5日`、`05.03.2026`、Excel 日期序列号 → `2026-03-05` |
| 金额统一 | `¥1,234.50`、`1234.5元`、`RMB 99.90`、`(100)`(负数)、`1.2万` → 数值,保留 2 位小数 |
| 手机号统一 | 去掉空格 / 横杠 / 括号 / `+86`、`0086` 前缀 → 11 位数字;校验规则可用 `--phone-pattern` 调整;Excel 中按文本保存,不会变成科学计数法 |
| 文本整理 | 去多余空格;`上海市` / ` 上海` → `上海`;订单号统一大写 |
| 去重 | 在**标准化之后**去重,`so-1` 与 `SO-1 ` 这类“看着不同其实一样”的行也能识别 |
| 清洗报告 | `*_report.md`(人读)+ `*_report.json`(程序读):输入/输出行数、每列空值/转换/无法识别数量;无法识别的原始值逐条列在 Excel 的“问题记录”工作表,**不静默丢数据** |

## 使用方法
```bash
pip install -r requirements.txt
python make_sample_data.py                                   # 生成虚构的脏数据
python cleaner.py sample_data/orders_messy.csv
python cleaner.py sample_data/orders_messy.xlsx --fill-amount median
python cleaner.py 你的文件.xlsx --sheet Sheet2 --required order_id,order_date \
    --phone-pattern '^1[3-9]\d{9}$' --column-map '{"客户编号": "customer_id"}' --out output/mine
python -m pytest -q   # 运行测试(38 个)
```

## 示例输出(真实运行结果)
```
orders_messy.csv:输入 44 行 → 输出 40 行;去重 2,全空行 1,缺必填 1,无法识别 3 个值
```
- `output/orders_messy_csv_clean.csv` / `.xlsx`:清洗结果(Excel 含“清洗结果”“问题记录”两个工作表)
- `output/orders_messy_csv_report.md` / `.json`:清洗报告
- `output/orders_messy_xlsx_*`:同一份数据的 Excel 版本(英文列名)+ 金额中位数填充

清洗前后对比(节选):
```
清洗前: so-2026037 | 03.01.2026 | " 2,227.8 " | (100)06201684 | " 上海市" | N/A
清洗后: SO-2026037 | 2026-01-03 | 2227.8      | 10006201684   | 上海      | (空)
```

## 关于示例数据
`sample_data/` 全部由 `make_sample_data.py` 随机生成:订单号是随机编号,手机号使用未分配的 `100` 号段,
不对应任何真实的人或号码,不包含姓名、地址等个人信息。因此默认手机号规则是“1 开头的 11 位数字”,
处理真实数据时建议加 `--phone-pattern '^1[3-9]\d{9}$'`。

## 技术栈
Python 3.10+、pandas、openpyxl、argparse、正则表达式、pytest(参数化测试)
