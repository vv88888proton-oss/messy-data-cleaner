"""生成故意弄脏的虚构订单数据(CSV + Excel)。

全部为随机生成的虚构数据:订单号是随机编号,手机号使用未分配的 100 号段(1000xxxxxxx),
不对应任何真实的人或号码;不包含姓名、地址等个人信息。
"""
import random
from pathlib import Path

import pandas as pd

random.seed(7)
OUT = Path(__file__).parent / "sample_data"
OUT.mkdir(exist_ok=True)

cities = ["上海", " 上海市", "北京", "北京市 ", "深圳", "杭州", "  成都"]
statuses = ["已完成", "已 完成", " 已发货", "待付款", "已取消", "N/A"]


def messy_date(y, m, d):
    fmt = random.choice(["{y}-{m:02d}-{d:02d}", "{y}/{m}/{d}", "{y}.{m:02d}.{d:02d}", "{y}{m:02d}{d:02d}",
                         "{y}年{m}月{d}日", "{d:02d}.{m:02d}.{y}"])
    return fmt.format(y=y, m=m, d=d)


def messy_amount(v):
    return random.choice([f"¥{v:,.2f}", f"{v}元", f" {v:,.1f} ", f"{v}", f"RMB {v:.2f}"])


def messy_phone():
    digits = "1000" + "".join(random.choice("0123456789") for _ in range(7))  # 未分配号段,虚构
    return random.choice([digits, f"+86 {digits[:3]}-{digits[3:7]}-{digits[7:]}", f"{digits[:3]} {digits[3:7]} {digits[7:]}",
                          f"86{digits}", f"({digits[:3]}){digits[3:]}"])


rows = []
for i in range(1, 41):
    rows.append({
        " 订单编号 ": f"so-{2026000 + i}" if i % 5 else f"SO-{2026000 + i}",
        "下单日期": messy_date(2026, random.randint(1, 10), random.randint(1, 28)),
        "金额(元)": messy_amount(round(random.uniform(20, 3000), 2)),
        "联系电话": messy_phone(),
        "城市": random.choice(cities),
        "订单状态": random.choice(statuses),
    })
# 故意制造的问题
rows[3]["金额(元)"] = "-"
rows[7]["下单日期"] = "下周一"
rows[9]["联系电话"] = "12345"
rows[12]["金额(元)"] = "待定"
rows.append(dict(rows[0]))                                   # 完全重复
rows.append({k: "" for k in rows[0]})                        # 全空行
rows.append({**rows[1], " 订单编号 ": ""})                   # 缺订单号
rows.append({**rows[2], " 订单编号 ": rows[2][" 订单编号 "].upper(), "城市": rows[2]["城市"] + " "})  # 标准化后重复
random.shuffle(rows)
df = pd.DataFrame(rows)
df.to_csv(OUT / "orders_messy.csv", index=False, encoding="utf-8-sig")
df.rename(columns={" 订单编号 ": "Order No", "下单日期": "Date", "金额(元)": "Amount", "联系电话": "Mobile"}) \
  .to_excel(OUT / "orders_messy.xlsx", index=False)
print(f"已生成 {OUT/'orders_messy.csv'}、{OUT/'orders_messy.xlsx'}({len(df)} 行,虚构数据)")
