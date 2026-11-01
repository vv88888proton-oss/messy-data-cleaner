from datetime import date

import pandas as pd
import pytest
from openpyxl import load_workbook

import cleaner as cl


@pytest.mark.parametrize("raw,expected", [
    (" 订单编号 ", "order_id"), ("Order No", "order_id"), ("金额(元)", "amount"),
    ("Mobile", "phone"), ("Customer  Type", "customer_type"), ("备注#1", "备注_1"),
])
def test_normalize_column_name(raw, expected):
    assert cl.normalize_column_name(raw) == expected


@pytest.mark.parametrize("raw,expected", [
    ("2026-03-05", date(2026, 3, 5)), ("2026/3/5", date(2026, 3, 5)), ("2026.03.05", date(2026, 3, 5)),
    ("20260305", date(2026, 3, 5)), ("2026年3月5日", date(2026, 3, 5)), ("05.03.2026", date(2026, 3, 5)),
    (46086, date(2026, 3, 5)), ("下周一", None), ("N/A", None),
])
def test_parse_date(raw, expected):
    assert cl.parse_date(raw) == expected


@pytest.mark.parametrize("raw,expected", [
    ("¥1,234.50", 1234.5), ("1234.5元", 1234.5), (" 1,234.0 ", 1234.0), ("RMB 99.90", 99.9),
    ("(100)", -100.0), ("1.2万", 12000.0), (88, 88.0), ("-", None), ("待定", None),
])
def test_parse_amount(raw, expected):
    assert cl.parse_amount(raw) == expected


@pytest.mark.parametrize("raw,expected", [
    ("10001234567", "10001234567"), ("+86 100-0123-4567", "10001234567"), ("100 0123 4567", "10001234567"),
    ("8610001234567", "10001234567"), ("(100)01234567", "10001234567"), (10001234567.0, "10001234567"),
    ("12345", None), ("", None),
])
def test_parse_phone(raw, expected):
    assert cl.parse_phone(raw) == expected


def test_phone_pattern_configurable():
    assert cl.parse_phone("10001234567", r"^1[3-9]\d{9}$") is None


def messy_df():
    return pd.DataFrame({
        " 订单编号 ": ["so-1", "SO-1", "so-2", "", "so-3", ""],
        "下单日期": ["2026/1/2", "2026-01-02", "下周一", "2026-01-05", "20260103", ""],
        "金额(元)": ["¥1,000", "1000元", "-", "50", "待定", ""],
        "联系电话": ["+86 100-0000-0001", "10000000001", "12345", "10000000002", "", ""],
        "城市": [" 上海市", "上海", "北京 ", "杭州", "N/A", ""],
    })


def test_clean_end_to_end_counts():
    df, issues, rep = cl.clean(messy_df(), source="t")
    assert rep.rows_in == 6
    assert rep.empty_rows_dropped == 1
    assert rep.missing_required_dropped == 1
    assert rep.duplicates_dropped == 1          # so-1 与 SO-1 标准化后完全一致
    assert rep.rows_out == 3
    assert list(df.columns) == ["order_id", "order_date", "amount", "phone", "city"]
    assert df["order_id"].tolist() == ["SO-1", "SO-3", "SO-2"]   # 按日期排序,无法识别的日期排最后
    assert df.loc[0, "city"] == "上海" and df.loc[0, "phone"] == "10000000001"
    assert set(issues["原始值"]) == {"下周一", "12345", "待定"}
    assert rep.nulls_normalized["amount"] == 2 and rep.nulls_normalized["city"] == 2  # 含全空行里的空串


@pytest.mark.parametrize("mode,expected", [("zero", 0.0), ("median", 1000.0)])
def test_fill_amount(mode, expected):
    df, _, rep = cl.clean(messy_df(), fill_amount=mode)
    assert df["amount"].isna().sum() == 0
    assert df.loc[df["order_id"] == "SO-2", "amount"].iloc[0] == expected
    assert rep.nulls_filled["amount"] == 2


def test_cli_writes_outputs(tmp_path):
    src = tmp_path / "in.csv"
    messy_df().to_csv(src, index=False, encoding="utf-8-sig")
    assert cl.main([str(src), "--out", str(tmp_path / "out")]) == 0
    out = tmp_path / "out"
    csv = pd.read_csv(out / "in_csv_clean.csv", dtype=str)
    assert csv.loc[0, "order_date"] == "2026-01-02"
    wb = load_workbook(out / "in_csv_clean.xlsx")
    assert wb.sheetnames == ["清洗结果", "问题记录"]
    assert wb["清洗结果"]["D2"].value == "10000000001"   # 手机号以文本保存
    report = (out / "in_csv_report.md").read_text(encoding="utf-8")
    assert "删除重复行:1" in report and "`联系电话` → `phone`" in report


def test_cli_missing_file(tmp_path):
    assert cl.main([str(tmp_path / "nope.csv")]) == 1
