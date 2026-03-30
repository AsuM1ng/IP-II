#!/usr/bin/env python3
"""按结局分组统计离散特征的样本数/比例，连续特征的均值/方差。"""

from __future__ import annotations

import argparse
import csv
import math
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

EPS = 1e-30


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "统计输入 CSV 中离散特征在两组结局下的样本数和比例，并统计连续特征的均值和方差。"
        )
    )
    parser.add_argument(
        "--input",
        default="data1sisclean.csv",
        help="输入 CSV 文件路径（默认 data1sisclean.csv）。",
    )
    parser.add_argument(
        "--output",
        default="feature_distribution_stats.csv",
        help="输出 CSV 文件路径。",
    )
    parser.add_argument(
        "--group-col",
        default="PulmonaryInfection",
        help="分组列名（默认 PulmonaryInfection）。",
    )
    parser.add_argument(
        "--positive-values",
        default="1,有,yes,true",
        help="判定为“肺部感染”组的取值（逗号分隔，忽略大小写）。",
    )
    parser.add_argument(
        "--positive-label",
        default="肺部感染",
        help="阳性分组展示名称。",
    )
    parser.add_argument(
        "--negative-label",
        default="非肺部感染",
        help="阴性分组展示名称。",
    )
    parser.add_argument(
        "--max-categories",
        type=int,
        default=20,
        help="特征唯一值数量超过该阈值时自动跳过（默认 20）。",
    )
    return parser.parse_args()


def normalize(v: str) -> str:
    return v.strip().lower()


def format_count_ratio(count: int, total: int) -> str:
    if total <= 0:
        return f"{count} (0.0%)"
    return f"{count} ({count / total * 100:.1f}%)"


def parse_float(value: str) -> float | None:
    text = value.strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def mean_variance(values: List[float]) -> Tuple[float, float]:
    if not values:
        return 0.0, 0.0
    mean = sum(values) / len(values)
    if len(values) == 1:
        return mean, 0.0
    variance = sum((x - mean) ** 2 for x in values) / (len(values) - 1)
    return mean, variance


def gammaincc(a: float, x: float) -> float:
    """Regularized upper incomplete gamma Q(a, x)."""
    if x < 0 or a <= 0:
        raise ValueError("invalid arguments for gammaincc")
    if x == 0:
        return 1.0
    if x < a + 1.0:
        # series for P(a, x)
        ap = a
        summ = 1.0 / a
        delta = summ
        for _ in range(10000):
            ap += 1.0
            delta *= x / ap
            summ += delta
            if abs(delta) < abs(summ) * 1e-14:
                break
        log_term = -x + a * math.log(x) - math.lgamma(a)
        p = summ * math.exp(log_term)
        return max(0.0, min(1.0, 1.0 - p))

    # continued fraction for Q(a, x)
    b = x + 1.0 - a
    c = 1.0 / EPS
    d = 1.0 / max(b, EPS)
    h = d
    for i in range(1, 10000):
        an = -i * (i - a)
        b += 2.0
        d = an * d + b
        if abs(d) < EPS:
            d = EPS
        c = b + an / c
        if abs(c) < EPS:
            c = EPS
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 1e-14:
            break
    log_term = -x + a * math.log(x) - math.lgamma(a)
    q = h * math.exp(log_term)
    return max(0.0, min(1.0, q))


def chi_square_p_value(contingency: Sequence[Sequence[int]]) -> float:
    rows = len(contingency)
    cols = len(contingency[0]) if rows else 0
    if rows < 2 or cols < 2:
        return 1.0

    row_totals = [sum(r) for r in contingency]
    col_totals = [sum(contingency[r][c] for r in range(rows)) for c in range(cols)]
    total = sum(row_totals)
    if total == 0:
        return 1.0

    chi2 = 0.0
    for r in range(rows):
        for c in range(cols):
            expected = row_totals[r] * col_totals[c] / total
            if expected <= 0:
                continue
            diff = contingency[r][c] - expected
            chi2 += (diff * diff) / expected

    dof = (rows - 1) * (cols - 1)
    if dof <= 0:
        return 1.0

    # chi-square survival function
    return gammaincc(dof / 2.0, chi2 / 2.0)


def format_p_value(p: float) -> str:
    if p < 0.001:
        return "<0.001"
    return f"{p:.3f}"


def read_csv_rows(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    last_exc: Exception | None = None
    for encoding in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            with path.open("r", encoding=encoding, newline="") as f:
                reader = csv.DictReader(f)
                if not reader.fieldnames:
                    raise ValueError("CSV header is empty")
                rows = [dict(row) for row in reader]
                return list(reader.fieldnames), rows
        except Exception as exc:
            last_exc = exc
    raise RuntimeError(f"读取文件失败: {path} ({last_exc})")


def choose_group_col(columns: Iterable[str], preferred: str) -> str:
    columns = list(columns)
    if preferred in columns:
        return preferred
    candidates = [
        "肺部感染（0=无、1=有）",
        "肺部感染",
        "PulmonaryInfection",
    ]
    for c in candidates:
        if c in columns:
            return c
    raise KeyError(f"未找到分组列: {preferred}")


def build_table(args: argparse.Namespace) -> List[Dict[str, str]]:
    input_path = Path(args.input)
    if not input_path.exists() and args.input == "data1sisclean.csv":
        fallback = Path("data1.csv")
        if fallback.exists():
            input_path = fallback
            print(f"[提示] 未找到 data1sisclean.csv，已自动使用 {fallback}。")

    columns, rows = read_csv_rows(input_path)
    group_col = choose_group_col(columns, args.group_col)

    positive_tokens = {normalize(v) for v in args.positive_values.split(",") if v.strip()}

    group_rows = {"pos": [], "neg": []}
    for row in rows:
        raw = (row.get(group_col) or "").strip()
        key = "pos" if normalize(raw) in positive_tokens else "neg"
        group_rows[key].append(row)

    pos_n = len(group_rows["pos"])
    neg_n = len(group_rows["neg"])

    out_rows: List[Dict[str, str]] = []

    for feature in columns:
        if feature == group_col:
            continue

        values = []
        numeric_values: List[float] = []
        is_numeric = True
        for row in rows:
            v = (row.get(feature) or "").strip()
            if v:
                values.append(v)
                parsed = parse_float(v)
                if parsed is None:
                    is_numeric = False
                else:
                    numeric_values.append(parsed)
        uniq = sorted(set(values))

        if len(uniq) == 0:
            continue

        # 连续变量：数值型且唯一值较多；离散变量按原逻辑处理不变。
        if is_numeric and len(uniq) > args.max_categories:
            pos_numeric = [
                parsed
                for r in group_rows["pos"]
                if (parsed := parse_float((r.get(feature) or "").strip())) is not None
            ]
            neg_numeric = [
                parsed
                for r in group_rows["neg"]
                if (parsed := parse_float((r.get(feature) or "").strip())) is not None
            ]
            pos_mean, pos_var = mean_variance(pos_numeric)
            neg_mean, neg_var = mean_variance(neg_numeric)
            out_rows.append(
                {
                    "特征": feature,
                    "分层": "均值",
                    f"{args.positive_label} N={pos_n}": f"{pos_mean:.4f}",
                    f"{args.negative_label} N={neg_n}": f"{neg_mean:.4f}",
                    "P-value": "",
                }
            )
            out_rows.append(
                {
                    "特征": "",
                    "分层": "方差",
                    f"{args.positive_label} N={pos_n}": f"{pos_var:.4f}",
                    f"{args.negative_label} N={neg_n}": f"{neg_var:.4f}",
                    "P-value": "",
                }
            )
            continue

        if len(uniq) > args.max_categories:
            continue

        pos_counter = Counter((r.get(feature) or "").strip() for r in group_rows["pos"])
        neg_counter = Counter((r.get(feature) or "").strip() for r in group_rows["neg"])

        category_counts = []
        for cat in uniq:
            category_counts.append([pos_counter.get(cat, 0), neg_counter.get(cat, 0)])

        p_value = chi_square_p_value(category_counts)
        first = True
        for cat in uniq:
            pos_count = pos_counter.get(cat, 0)
            neg_count = neg_counter.get(cat, 0)
            out_rows.append(
                {
                    "特征": feature if first else "",
                    "分层": cat,
                    f"{args.positive_label} N={pos_n}": format_count_ratio(pos_count, pos_n),
                    f"{args.negative_label} N={neg_n}": format_count_ratio(neg_count, neg_n),
                    "P-value": format_p_value(p_value) if first else "",
                }
            )
            first = False

    return out_rows


def write_csv(path: Path, rows: List[Dict[str, str]]) -> None:
    if not rows:
        raise ValueError("没有可写入的统计结果")
    fieldnames = list(rows[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    rows = build_table(args)
    output_path = Path(args.output)
    write_csv(output_path, rows)
    print(f"统计完成，输出文件：{output_path}")


if __name__ == "__main__":
    main()
