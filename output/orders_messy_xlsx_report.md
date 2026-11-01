# 清洗报告:orders_messy.xlsx

- 输入行数:44
- 输出行数:40
- 删除全空行:1
- 删除重复行:2
- 删除缺少必填字段的行:1

## 列名标准化

- `Order No` → `order_id`
- `Date` → `order_date`
- `Amount` → `amount`
- `Mobile` → `phone`
- `城市` → `city`
- `订单状态` → `status`

## 字段处理

| 列 | 空值标记转空 | 成功转换 | 无法识别(置空并记录) | 填充空值 |
|---|---|---|---|---|
| amount | 2 | 41 | 1 | 2 |
| city | 1 | — | 0 | 0 |
| order_date | 1 | 42 | 1 | 0 |
| order_id | 2 | — | 0 | 0 |
| phone | 1 | 42 | 1 | 0 |
| status | 6 | — | 0 | 0 |

无法识别的原始值见输出 Excel 的“问题记录”工作表。
