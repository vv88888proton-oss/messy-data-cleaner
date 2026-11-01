"""05 脏数据清洗命令行工具:读 CSV/Excel → 列名标准化 → 空值/日期/金额/手机号统一 → 去重 → 输出清洗结果和清洗报告。"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

NULL_TOKENS = {"", "na", "n/a", "nan", "null", "none", "-", "--", "—", "无", "空", "未知", "#n/a"}

# 常见列名别名 → 标准列名(可用 --column-map 追加/覆盖)
COLUMN_ALIASES = {
    "订单编号": "order_id", "订单号": "order_id", "order_no": "order_id", "orderid": "order_id", "order_id": "order_id",
    "下单日期": "order_date", "日期": "order_date", "date": "order_date", "order_date": "order_date",
    "金额": "amount", "订单金额": "amount", "金额(元)": "amount", "amount": "amount", "price": "amount",
    "手机号": "phone", "手机": "phone", "联系电话": "phone", "mobile": "phone", "phone": "phone",
    "城市": "city", "city": "city",
    "状态": "status", "订单状态": "status", "status": "status",
}

DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%d.%m.%Y", "%Y年%m月%d日",
                "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%m/%d/%Y"]


@dataclass
class Report:
    source: str = ""
    rows_in: int = 0
    rows_out: int = 0
    column_renames: dict = field(default_factory=dict)
    empty_rows_dropped: int = 0
    duplicates_dropped: int = 0
    missing_required_dropped: int = 0
    nulls_normalized: dict = field(default_factory=dict)
    nulls_filled: dict = field(default_factory=dict)
    invalid: dict = field(default_factory=dict)
    converted: dict = field(default_factory=dict)

    def to_markdown(self) -> str:
        lines = [f"# 清洗报告:{self.source}", "",
                 f"- 输入行数:{self.rows_in}", f"- 输出行数:{self.rows_out}",
                 f"- 删除全空行:{self.empty_rows_dropped}", f"- 删除重复行:{self.duplicates_dropped}",
                 f"- 删除缺少必填字段的行:{self.missing_required_dropped}", "", "## 列名标准化", ""]
        lines += [f"- `{k}` → `{v}`" for k, v in self.column_renames.items()] or ["- 无变化"]
        lines += ["", "## 字段处理", "", "| 列 | 空值标记转空 | 成功转换 | 无法识别(置空并记录) | 填充空值 |", "|---|---|---|---|---|"]
        cols = sorted(set(self.nulls_normalized) | set(self.converted) | set(self.invalid) | set(self.nulls_filled))
        for c in cols:
            lines.append(f"| {c} | {self.nulls_normalized.get(c, 0)} | {self.converted.get(c, '—')} | "
                         f"{self.invalid.get(c, 0)} | {self.nulls_filled.get(c, 0)} |")
        lines += ["", "无法识别的原始值见输出 Excel 的“问题记录”工作表。", ""]
        return "\n".join(lines)


# ---------------------------------------------------------------- 单值解析
def normalize_column_name(name, aliases: dict = COLUMN_ALIASES) -> str:
    s = re.sub(r"\s+", " ", str(name)).strip()
    s = s.replace("(", "(").replace(")", ")")
    key = s.lower().replace(" ", "_")
    if s in aliases:
        return aliases[s]
    if key in aliases:
        return aliases[key]
    key = re.sub(r"[^\w]+", "_", key).strip("_")
    return aliases.get(key, key)


def is_null(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return isinstance(v, str) and v.strip().lower() in NULL_TOKENS


def parse_date(v):
    """多种写法 → datetime.date;无法识别返回 None。支持 Excel 序列号。"""
    if is_null(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.date()
    if isinstance(v, (int, float)) and 20000 < v < 80000:  # Excel 日期序列号
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    s = str(v).strip()
    if re.fullmatch(r"\d{5}(\.0)?", s):
        return parse_date(float(s))
    s = re.sub(r"\s+", " ", s)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_amount(v):
    """'¥1,234.50' / '1234.5元' / '(100)' / '1.2万' → float;无法识别返回 None。"""
    if is_null(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace(",", "").replace(" ", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = re.sub(r"^(¥|¥|RMB|CNY|\$)", "", s, flags=re.I)
    s = re.sub(r"(元|RMB|CNY)$", "", s, flags=re.I)
    mult = 1.0
    if s.endswith("万"):
        mult, s = 10000.0, s[:-1]
    try:
        val = float(s) * mult
    except ValueError:
        return None
    return -val if neg else val


def parse_phone(v, pattern: str = r"^1\d{10}$"):
    """去掉空格/横杠/括号/+86 前缀,统一成 11 位数字;不符合 pattern 返回 None。"""
    if is_null(v):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"[\s\-()()]", "", str(v))
    s = re.sub(r"^(\+?86|0086)(?=1\d{10}$)", "", s)
    return s if re.fullmatch(pattern, s) else None


def normalize_text(v):
    if is_null(v):
        return None
    return re.sub(r"\s+", " ", str(v)).strip()


# ---------------------------------------------------------------- 整表清洗
def clean(df: pd.DataFrame, *, required=("order_id",), fill_amount: str = "none", phone_pattern=r"^1\d{10}$",
          column_map: dict | None = None, source: str = "") -> tuple[pd.DataFrame, pd.DataFrame, Report]:
    rep = Report(source=source, rows_in=len(df))
    aliases = {**COLUMN_ALIASES, **(column_map or {})}
    new_cols = [normalize_column_name(c, aliases) for c in df.columns]
    rep.column_renames = {str(o): n for o, n in zip(df.columns, new_cols) if str(o) != n}
    df = df.copy()
    df.columns = new_cols
    df = df.loc[:, ~df.columns.duplicated()]
    issues = []

    # 空值标记统一为 None
    for c in df.columns:
        mask = df[c].map(lambda x: isinstance(x, str) and is_null(x))
        if mask.any():
            rep.nulls_normalized[c] = int(mask.sum())
        df[c] = df[c].map(lambda x: None if is_null(x) else x).astype(object)

    before = len(df)
    df = df.dropna(how="all")
    rep.empty_rows_dropped = before - len(df)

    parsers = {"order_date": parse_date, "amount": parse_amount,
               "phone": lambda x: parse_phone(x, phone_pattern)}
    for c in df.columns:
        fn = parsers.get(c, normalize_text)
        raw = df[c]
        parsed = raw.map(fn)
        bad = raw.notna() & parsed.isna()
        if c in parsers:
            rep.converted[c] = int((raw.notna() & parsed.notna()).sum())
        if bad.any():
            rep.invalid[c] = int(bad.sum())
            for idx in raw[bad].index:
                issues.append({"原始行号": int(idx) + 2, "列": c, "原始值": str(raw[idx]), "问题": "无法识别,已置空"})
        df[c] = parsed

    if "status" in df.columns:
        df["status"] = df["status"].map(lambda s: s.replace(" ", "") if isinstance(s, str) else None)
    if "city" in df.columns:
        df["city"] = df["city"].map(lambda s: s.rstrip("市") if isinstance(s, str) else None)

    if "order_id" in df.columns:
        df["order_id"] = df["order_id"].map(lambda s: None if is_null(s) else str(s).upper().replace(" ", ""))

    req = [c for c in required if c in df.columns]
    if req:
        before = len(df)
        df = df.dropna(subset=req)
        rep.missing_required_dropped = before - len(df)

    before = len(df)
    df = df.drop_duplicates()
    rep.duplicates_dropped = before - len(df)

    if "amount" in df.columns and fill_amount != "none":
        n = int(df["amount"].isna().sum())
        if n:
            val = 0.0 if fill_amount == "zero" else float(pd.to_numeric(df["amount"]).median())
            df["amount"] = df["amount"].fillna(val)
            rep.nulls_filled["amount"] = n
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"]).round(2)
    if "order_date" in df.columns:
        df["order_date"] = pd.to_datetime(df["order_date"])
        df = df.sort_values("order_date", kind="stable")

    df = df.reset_index(drop=True)
    rep.rows_out = len(df)
    return df, pd.DataFrame(issues, columns=["原始行号", "列", "原始值", "问题"]), rep


# ---------------------------------------------------------------- IO / 命令行
def read_any(path: Path, sheet=None) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, sheet_name=sheet or 0, dtype=object, keep_default_na=False)
    for enc in ("utf-8-sig", "gbk"):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"无法识别文件编码:{path}")


def write_outputs(df, issues, rep: Report, out_dir: Path, stem: str) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path, xlsx_path, md_path = out_dir / f"{stem}_clean.csv", out_dir / f"{stem}_clean.xlsx", out_dir / f"{stem}_report.md"
    csv_df = df.copy()
    if "order_date" in csv_df.columns:
        csv_df["order_date"] = csv_df["order_date"].dt.strftime("%Y-%m-%d")
    csv_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", datetime_format="YYYY-MM-DD", date_format="YYYY-MM-DD") as xw:
        df.to_excel(xw, sheet_name="清洗结果", index=False)
        issues.to_excel(xw, sheet_name="问题记录", index=False)
        ws = xw.sheets["清洗结果"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 16
        if "phone" in df.columns:  # 手机号按文本保存,避免 Excel 显示成科学计数法
            j = list(df.columns).index("phone") + 1
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                row[0].number_format = "@"
        if "order_date" in df.columns:
            j = list(df.columns).index("order_date") + 1
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                row[0].number_format = "yyyy-mm-dd"
    md_path.write_text(rep.to_markdown(), encoding="utf-8")
    (out_dir / f"{stem}_report.json").write_text(json.dumps(rep.__dict__, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"csv": csv_path, "xlsx": xlsx_path, "report": md_path}


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="清洗脏 CSV / Excel 并输出清洗报告")
    p.add_argument("input", help="输入文件(.csv / .xlsx)")
    p.add_argument("--sheet", help="Excel 工作表名,默认第一个")
    p.add_argument("--out", default="output", help="输出目录")
    p.add_argument("--required", default="order_id", help="必填列(标准化后的列名),逗号分隔;传空字符串表示不检查")
    p.add_argument("--fill-amount", choices=["none", "zero", "median"], default="none", help="金额空值处理")
    p.add_argument("--phone-pattern", default=r"^1\d{10}$", help="手机号校验正则(清洗后),如 ^1[3-9]\\d{9}$")
    p.add_argument("--column-map", help='额外列名映射 JSON,如 \'{"客户编号": "customer_id"}\'')
    return p.parse_args(argv)


def main(argv=None) -> int:
    a = parse_args(argv)
    path = Path(a.input)
    if not path.exists():
        print(f"文件不存在:{path}", file=sys.stderr)
        return 1
    raw = read_any(path, a.sheet)
    df, issues, rep = clean(raw, required=tuple(x for x in a.required.split(",") if x), fill_amount=a.fill_amount,
                            phone_pattern=a.phone_pattern,
                            column_map=json.loads(a.column_map) if a.column_map else None, source=path.name)
    paths = write_outputs(df, issues, rep, Path(a.out), f"{path.stem}_{path.suffix.lstrip('.').lower()}")
    print(f"{path.name}:输入 {rep.rows_in} 行 → 输出 {rep.rows_out} 行;去重 {rep.duplicates_dropped},"
          f"全空行 {rep.empty_rows_dropped},缺必填 {rep.missing_required_dropped},无法识别 {sum(rep.invalid.values())} 个值")
    print("已输出:" + "、".join(str(p) for p in paths.values()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
