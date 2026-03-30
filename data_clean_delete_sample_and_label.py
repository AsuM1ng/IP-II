"""清洗头颈肿瘤围手术期原始数据。

功能概述：
1. 读取 `.xlsx` 原始数据。
2. 将中文字段名统一为简洁英文名。
3. 由入院日期、出院日期计算住院时长，并删除原日期列。
4. 删除不需要直接建模的字段（如身高、体重、手术日期）。
5. 数据处理前先统计各字段缺失率，并输出审阅摘要。
6. 缺失率高于 30% 的字段直接删除。
7. 以“肺部感染”为因变量：保留字段中，阳性样本缺失值补齐，阴性样本含缺失的记录直接删除。
8. 连续变量完成缺失值处理后进行 Z-score 标准化。
9. 分类变量缺失值以众数填充；二分类编码为 0/1，多分类编码为顺序整数。
10. BMI 按连续变量处理，完成缺失值填补后进行 Z-score 标准化。
11. 为类别很多、未来可能出现新类别的字段保存编码映射，便于复用。

说明：
- 本脚本优先依据“原始数据节选”的列名工作。
- 在完整数据中，即使分类变量出现更多新类别，脚本也会自动纳入编码映射。
- 输出三个文件：
  1) `data1.csv`：清洗后的数据；
  2) `category_mappings.json`：多分类/二分类字段编码字典；
  3) `data_review_summary.txt`：处理前数据审阅摘要。
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pandas as pd
from scipy.stats import chi2_contingency, ttest_ind


READ_PATH = Path("数据_清洗后_切口感染.xlsx")
WRITE_PATH = Path("data1.csv")
MAPPING_PATH = Path("category_mappings.json")
REVIEW_PATH = Path("data_review_summary.txt")
STAT_TABLE_PATH = Path("group_statistics.tsv")
TARGET_COLUMN = "IncisionInfection"
FEATURE_MISSING_THRESHOLD = 0.3
column_to_drop = []
# column_to_drop = ["HospitalizationCount", "DischargeDepartment", "PrimaryDiagnosis", "PrimarySurgeryName",
#                   "Height", "Weight", "OperationDate", "Hypertension", "Hyperlipidemia", "SteroidUse", "SurgeryName",
#                   "ImmuneDisease",  "UnplannedReoperation", "AnastomoticFistula", "DaysToFistulaConfirmation",
#                   "FatLiquefaction", "IncisionInfection", "Pathogen", "MultiDrugResistance", 'LesionSite', 'PreopAntibiotic', 'FlapType', 'PostopPathology']


# ---------------------------------------------------------------------------
# 1. 中文字段名 -> 英文字段名
# ---------------------------------------------------------------------------
# 仅保留核心业务含义；如果源数据中存在这些列，则会自动重命名。
COLUMN_RENAME_MAP = {
    "住院次数": "HospitalizationCount",
    "性别\n（1=男，2=女）": "Sex",
    "年龄": "Age",
    "入院日期": "AdmissionDate",
    "出院日期": "DischargeDate",
    "出院科室": "DischargeDepartment",
    "主要诊断": "PrimaryDiagnosis",
    "主要手术名称": "PrimarySurgeryName",
    "主要手术愈合等级": "PrimaryWoundHealingGrade",
    "身高": "Height",
    "体重": "Weight",
    "BMI": "BMI",
    "既往头颈部疾病手术外伤史（0=无，1=有）": "PriorHeadNeckHistory",
    "冠心病（0=无，1=有）": "CoronaryHeartDisease",
    "高血压（0=无，1=有）": "Hypertension",
    "外周血管病（0=无，1=有）": "PeripheralVascularDisease",
    "免疫性疾病（0=无，1=有）": "ImmuneDisease",
    "糖尿病（0=无，1=有）": "Diabetes",
    "高脂血症（0=无，1=有）": "Hyperlipidemia",
    "激素类（0=无，1=有）": "SteroidUse",
    "吸烟史（0=无，1=有）": "SmokingHistory",
    "饮酒史（0=无，1=有）": "AlcoholHistory",
    "术前前白蛋白PALB（0=未测）": "PreopPALB",
    "术前白蛋白ALB（0=未测）": "PreopALB",
    "术前血红蛋白HGB（0=未测）": "PreopHGB",
    "术前口咽拭子（未生长=0，阳性=1，未测=2）": "PreopOropharyngealSwab",
    "病变部位（喉=1、鼻=2、扁桃体=3、腮腺=4、咽部=5、唇=6、甲状腺=7、食管=8、舌=9、耳=10 、颌、面部=11、口底=12、气管=13、口腔、牙龈=14、梨状窝=15、上颌窦=16、皮肤肿物=17、腭=18、胸部=19、颈部=20）": "LesionSite",
    "ASA评分": "ASA",
    "术前抗菌药物：未用=0、头孢曲松钠=1、头孢呋辛钠=2、左奥硝唑氯化钠=3、吗啉硝唑氯化钠=4、甲磺酸左氧氟沙星氯化钠=5、克林霉素磷酸酯=6、盐酸莫西沙星=7、头孢噻肟钠=8、甲硝唑氯化钠=9、头孢米诺钠=10、奥硝唑=11、美洛西林钠舒巴坦钠=12、注射用哌拉西林钠舒巴坦钠=13、头孢哌酮钠舒巴坦钠=14、阿奇霉素=15、美罗培南=16、万古霉素=17": "PreopAntibiotic",
    "手术时长（min）（不详=0，其他具体写出）": "OperationDurationMin",
    "术中输血（有=1，无=0）": "IntraopTransfusion",
    "吻合方式（如有吻合口：手工=1，器械=2，不详=0）": "AnastomosisType",
    "颈清扫（0=无，1=有）": "NeckDissection",
    "术前放疗（0=无，1=有）": "PreopRadiotherapy",
    "术前化疗（0=无，1=有）": "PreopChemotherapy",
    "术前同步放化疗（0=无，1:<=60Gy,2:>60Gy，有，但具体剂量不详=3）": "PreopConcurrentCRT",
    "手术日期": "OperationDate",
    "腔镜（0=无，1=有）": "Endoscopy",
    "手术名称": "SurgeryName",
    "皮瓣（无=0、颏下皮瓣=1、股前外侧皮瓣=2、鼻唇沟=3、颈阔肌=4、锁骨=5、其他=6、前臂桡侧皮瓣=7、游离空肠瓣=8、游离腓骨瓣=9、胸大肌皮瓣=10、局部转移皮瓣=11）  待整理": "FlapType",
    "气管造瘘（无=0，有=1）": "Tracheostomy",
    "术后病理无=0、warthin瘤=1、鳞癌=2、乳头状癌=3、多形性腺瘤=4、基底细胞癌=5、腺癌=6、黑色素瘤=7、未见癌=8、肉瘤=9、良性=10、甲状腺髓样癌=11、粘液表皮样癌=12、分化差的癌=13、腺样囊性癌=14、淋巴细胞瘤=15、梭形细胞瘤=16、囊肿=17": "PostopPathology",
    "分化（无=0、低=1、中=2、高=3、未确定=4）": "Differentiation",
    "最新版pTNM（5=5期，1=I期，2=II期，3=III期，4=IV期，5=无）": "pTNMStage",
    "多原发（0=否、1=是）": "MultiplePrimary",
    "术后0-3天白蛋白ALB（选最低的）    （0=未测）": "PostopDay0to3ALB",
    "术后0-3天前白蛋白PALB（0=未测）": "PostopDay0to3PALB",
    "非计划二次手术(0=否，1=是)": "UnplannedReoperation",
    "肺部感染（0=无、1=有）": "PulmonaryInfection",
    "吻合口瘘（0=无、1=有）": "AnastomoticFistula",
    "吻合口瘘确认距术后天数（0=未发生，具体已写出）": "DaysToFistulaConfirmation",
    "脂肪液化（0=无，1=有）": "FatLiquefaction",
    "切口感染(0=无、1=有）": "IncisionInfection",
    "感染病原体": "Pathogen",
    "是否多重耐药（0=否，1＝是）": "MultiDrugResistance",
}


# ---------------------------------------------------------------------------
# 2. 工具函数
# ---------------------------------------------------------------------------
def normalize_missing(value: Any) -> Any:
    """统一识别空值表达。"""
    if pd.isna(value):
        return pd.NA
    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned == "" or cleaned.lower() in {"nan", "none", "null", "na", "n/a", "不详", "未知"}:
            return pd.NA
        return cleaned
    return value


FULLWIDTH_TRANSLATION = str.maketrans({
    "（": "(",
    "）": ")",
    "：": ":",
    "，": ",",
    "、": ",",
    "＝": "=",
    "－": "-",
    "≤": "<=",
    "≥": ">=",
    "　": " ",
})


def normalize_category_value(value: Any) -> Any:
    """标准化分类值，减少同义写法对编码的影响。"""
    value = normalize_missing(value)
    if pd.isna(value):
        return pd.NA

    if isinstance(value, str):
        text = value.translate(FULLWIDTH_TRANSLATION)
        text = re.sub(r"\s+", " ", text).strip()
        return text
    return value


ROMAN_TO_INT = {"Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5}


def convert_numeric_like(value: Any) -> Any:
    """尽量把字符串中的数值转成数字；无法转换则保留原值。"""
    value = normalize_missing(value)
    if pd.isna(value):
        return pd.NA

    if isinstance(value, str):
        text = normalize_category_value(value)
        if text in ROMAN_TO_INT:
            return ROMAN_TO_INT[text]

        matched = re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text)
        if matched:
            number = float(text)
            return int(number) if number.is_integer() else number
        return text
    return value


def safe_to_datetime(series: pd.Series) -> pd.Series:
    """把日期列转为 pandas datetime；无法识别的值转为 NaT。"""
    return pd.to_datetime(series.apply(normalize_missing), errors="coerce")


# ---------------------------------------------------------------------------
# 3. 日期与数值字段处理
# ---------------------------------------------------------------------------
def add_length_of_stay(df: pd.DataFrame) -> pd.DataFrame:
    """计算住院时长（天），并删除入院/出院日期列。"""
    if {"AdmissionDate", "DischargeDate"}.issubset(df.columns):
        admission = safe_to_datetime(df["AdmissionDate"])
        discharge = safe_to_datetime(df["DischargeDate"])
        length_of_stay = (discharge - admission).dt.days
        # 若数据中存在当天入/出院，希望保留为 0；若出院早于入院，则置为空待后续填补。
        length_of_stay = length_of_stay.where(length_of_stay >= 0, pd.NA)
        df["LengthOfStay"] = length_of_stay
        df = df.drop(columns=["AdmissionDate", "DischargeDate"])
    return df


def preprocess_numeric_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """识别并整理连续变量。"""
    numeric_candidates = [
        "HospitalizationCount",
        "Age",
        "BMI",
        "PreopPALB",
        "PreopALB",
        "PreopHGB",
        "OperationDurationMin",
        "PostopDay0to3ALB",
        "PostopDay0to3PALB",
        "DaysToFistulaConfirmation",
        "LengthOfStay",
    ]

    available_numeric_columns: list[str] = []
    for column in numeric_candidates:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column].apply(convert_numeric_like), errors="coerce")
            available_numeric_columns.append(column)

    return df, available_numeric_columns


# ---------------------------------------------------------------------------
# 4. 数据审阅与分组缺失值处理
# ---------------------------------------------------------------------------
def build_data_review_summary(df: pd.DataFrame, target_column: str = TARGET_COLUMN) -> str:
    """生成数据处理前的表格审阅摘要。"""
    total_rows, total_columns = df.shape
    lines = [
        "数据处理前审阅摘要",
        "=" * 24,
        f"样本量: {total_rows}",
        f"字段数: {total_columns}",
        "",
        "各字段缺失情况（按缺失率降序）:",
    ]

    missing_stats = (
        pd.DataFrame({
            "missing_count": df.isna().sum(),
            "missing_ratio": df.isna().mean(),
        })
        .sort_values(by=["missing_ratio", "missing_count"], ascending=False)
    )
    for column, row in missing_stats.iterrows():
        lines.append(f"- {column}: 缺失 {int(row['missing_count'])} / {total_rows} ({row['missing_ratio']:.1%})")

    if target_column in df.columns:
        lines.extend(["", f"按因变量 {target_column} 分组审阅:"])
        target_missing_mask = df[target_column].isna()
        if target_missing_mask.any():
            lines.append(f"- 因变量缺失样本数: {int(target_missing_mask.sum())}")

        for target_value, group_df in df.groupby(target_column, dropna=True):
            row_missing_ratio = group_df.isna().mean(axis=1)
            lines.append(
                f"- {target_column}={target_value}: {len(group_df)} 例，"
                f"行均缺失率 {row_missing_ratio.mean():.1%}，"
                f"中位缺失率 {row_missing_ratio.median():.1%}，"
                f"最大缺失率 {row_missing_ratio.max():.1%}"
            )
    else:
        lines.extend(["", f"未找到因变量列：{target_column}"])

    return "\n".join(lines)


def review_raw_data(df: pd.DataFrame, review_path: Path = REVIEW_PATH, target_column: str = TARGET_COLUMN) -> None:
    """保存并打印数据处理前的审阅结果。"""
    review_summary = build_data_review_summary(df, target_column=target_column)
    review_path.write_text(review_summary, encoding="utf-8")
    print(review_summary)


def report_and_drop_high_missing_features(
    df: pd.DataFrame,
    feature_missing_threshold: float = FEATURE_MISSING_THRESHOLD,
    target_column: str = TARGET_COLUMN,
) -> tuple[pd.DataFrame, list[str]]:
    """统计各字段缺失率，删除缺失率过高的字段，并输出结果。"""
    missing_ratio = df.isna().mean().sort_values(ascending=False)
    print("\n各字段缺失率统计：")
    for column, ratio in missing_ratio.items():
        print(f"- {column}: {ratio:.1%}")

    high_missing_features = [
        column
        for column, ratio in missing_ratio.items()
        if ratio > feature_missing_threshold and column != target_column
    ]

    if high_missing_features:
        print(f"\n缺失率高于 {feature_missing_threshold:.0%} 的字段将被删除：")
        for column in high_missing_features:
            print(f"- {column}")
        df = df.drop(columns=high_missing_features)
    else:
        print(f"\n没有字段的缺失率高于 {feature_missing_threshold:.0%}。")

    return df, high_missing_features


def split_by_target_and_handle_missing(
    df: pd.DataFrame,
    numeric_columns: list[str],
    target_column: str = TARGET_COLUMN,
) -> pd.DataFrame:
    """按肺部感染分组处理保留字段中的缺失值。

    - 阳性样本（1）：保留，并按原规则补齐缺失值；
    - 阴性样本（0）：只要存在缺失值则直接删除；
    - 其他样本：沿用原规则补齐连续变量。
    """
    if target_column not in df.columns:
        return df

    positive_mask = df[target_column] == 1
    negative_mask = df[target_column] == 0

    positive_df = df.loc[positive_mask].copy()
    negative_df = df.loc[negative_mask].copy()
    other_df = df.loc[~(positive_mask | negative_mask)].copy()

    if not positive_df.empty:
        positive_df = fill_numeric_missing_with_median(positive_df, numeric_columns)

    dropped_negative_count = 0
    if not negative_df.empty:
        keep_negative_mask = ~negative_df.isna().any(axis=1)
        dropped_negative_count = int((~keep_negative_mask).sum())
        negative_df = negative_df.loc[keep_negative_mask].copy()

    if not other_df.empty:
        other_df = fill_numeric_missing_with_median(other_df, numeric_columns)

    combined_df = pd.concat([positive_df, negative_df, other_df], axis=0).sort_index()
    print(
        "按因变量分组处理缺失值完成："
        f"阳性样本保留 {len(positive_df)} 例并完成填补；"
        f"阴性样本删除含缺失 {dropped_negative_count} 例，保留 {len(negative_df)} 例。"
    )
    return combined_df


# ---------------------------------------------------------------------------
# 5. 缺失值处理与统计
# ---------------------------------------------------------------------------
def fill_numeric_missing_with_median(df: pd.DataFrame, numeric_columns: list[str]) -> pd.DataFrame:
    """连续变量以中位数填补缺失值。"""
    for column in numeric_columns:
        median_value = df[column].median(skipna=True)
        if not pd.isna(median_value):
            df[column] = df[column].fillna(median_value)
    return df


def format_count_ratio(count: int, denominator: int) -> str:
    """把计数和占比格式化为 'n (xx.x%)'。"""
    if denominator <= 0:
        return "0 (0.0%)"
    ratio = count / denominator * 100
    return f"{count} ({ratio:.1f}%)"


def format_mean_std(series: pd.Series) -> str:
    """把连续变量格式化为 '均值 ± 标准差'。"""
    numeric_series = pd.to_numeric(series, errors="coerce").dropna()
    if numeric_series.empty:
        return "NA"
    mean_value = numeric_series.mean()
    std_value = numeric_series.std(ddof=1)
    if pd.isna(std_value):
        std_value = 0.0
    return f"{mean_value:.2f} ± {std_value:.2f}"


def compute_continuous_pvalue(infection: pd.Series, non_infection: pd.Series) -> str:
    """计算连续变量组间 p 值（Welch t 检验）。"""
    infection = pd.to_numeric(infection, errors="coerce").dropna()
    non_infection = pd.to_numeric(non_infection, errors="coerce").dropna()
    if len(infection) < 2 or len(non_infection) < 2:
        return ""
    pvalue = ttest_ind(infection, non_infection, equal_var=False, nan_policy="omit").pvalue
    if pd.isna(pvalue):
        return ""
    return f"{pvalue:.3f}"


def compute_categorical_pvalue(df: pd.DataFrame, column: str, target_column: str) -> str:
    """计算分类变量组间 p 值（卡方检验）。"""
    valid_df = df[[column, target_column]].dropna()
    if valid_df.empty:
        return ""

    contingency = pd.crosstab(valid_df[column], valid_df[target_column])
    if contingency.shape[0] < 2 or contingency.shape[1] < 2:
        return ""

    _, pvalue, _, _ = chi2_contingency(contingency)
    if pd.isna(pvalue):
        return ""
    return f"{pvalue:.3f}"


def build_group_statistics_table(
    df: pd.DataFrame,
    numeric_columns: list[str],
    target_column: str = TARGET_COLUMN,
) -> pd.DataFrame:
    """按肺部感染分组输出统计表格（分类：n/%；连续：均值±标准差）。"""
    if target_column not in df.columns:
        return pd.DataFrame()

    analysis_df = df[df[target_column].isin([0, 1])].copy()
    if analysis_df.empty:
        return pd.DataFrame()

    infection_df = analysis_df[analysis_df[target_column] == 1]
    non_infection_df = analysis_df[analysis_df[target_column] == 0]
    infection_n = len(infection_df)
    non_infection_n = len(non_infection_df)

    rows: list[dict[str, str]] = [
        {"变量": "", "肺部感染": f"N={infection_n}", "非肺部感染": f"N={non_infection_n}", "P-value": ""}
    ]

    for column in analysis_df.columns:
        if column == target_column:
            continue

        if column in numeric_columns:
            rows.append(
                {
                    "变量": column,
                    "肺部感染": format_mean_std(infection_df[column]),
                    "非肺部感染": format_mean_std(non_infection_df[column]),
                    "P-value": compute_continuous_pvalue(infection_df[column], non_infection_df[column]),
                }
            )
            continue

        pvalue = compute_categorical_pvalue(analysis_df, column, target_column)
        rows.append({"变量": column, "肺部感染": "", "非肺部感染": "", "P-value": pvalue})

        levels = (
            analysis_df[column]
            .dropna()
            .astype(str)
            .sort_values()
            .unique()
            .tolist()
        )
        for level in levels:
            infection_count = int((infection_df[column].astype(str) == level).sum())
            non_infection_count = int((non_infection_df[column].astype(str) == level).sum())
            rows.append(
                {
                    "变量": f"  {level}",
                    "肺部感染": format_count_ratio(infection_count, infection_n),
                    "非肺部感染": format_count_ratio(non_infection_count, non_infection_n),
                    "P-value": "",
                }
            )

    return pd.DataFrame(rows, columns=["变量", "肺部感染", "非肺部感染", "P-value"])


# ---------------------------------------------------------------------------
# 6. 分类变量编码
# ---------------------------------------------------------------------------
def encode_binary_series(series: pd.Series) -> tuple[pd.Series, dict[str, int]]:
    """二分类变量编码为 0/1。"""
    normalized = series.apply(normalize_category_value)
    mode = normalized.mode(dropna=True)
    fill_value = mode.iloc[0] if not mode.empty else "Missing"
    normalized = normalized.fillna(fill_value)

    unique_values = list(pd.unique(normalized))
    if len(unique_values) != 2:
        raise ValueError("当前列不是二分类变量，不能使用 0/1 编码。")

    # 排序保证编码结果稳定；同时尽量让 0/1 编码可复现。
    sorted_values = sorted(unique_values, key=lambda item: str(item))
    mapping = {str(value): index for index, value in enumerate(sorted_values)}
    encoded = normalized.map(lambda item: mapping[str(item)]).astype("Int64")
    return encoded, mapping


def encode_multiclass_series(series: pd.Series) -> tuple[pd.Series, dict[str, int]]:
    """多分类变量编码为顺序整数，并预留 0 作为未知类别编号。"""
    normalized = series.apply(normalize_category_value)
    mode = normalized.mode(dropna=True)
    fill_value = mode.iloc[0] if not mode.empty else "Missing"
    normalized = normalized.fillna(fill_value)

    unique_values = sorted({str(value) for value in pd.unique(normalized)})
    mapping = {value: index + 1 for index, value in enumerate(unique_values)}
    mapping["__UNKNOWN__"] = 0
    encoded = normalized.map(lambda item: mapping.get(str(item), 0)).astype("Int64")
    return encoded, mapping


def encode_categorical_columns(df: pd.DataFrame, numeric_columns: list[str]) -> tuple[pd.DataFrame, dict[str, dict[str, int]]]:
    """对所有非连续变量做编码，并输出编码字典。"""
    categorical_columns = [column for column in df.columns if column not in numeric_columns]
    category_mappings: dict[str, dict[str, int]] = {}

    for column in categorical_columns:
        non_null_count = df[column].dropna().nunique()
        if non_null_count <= 1:
            # 单值列直接填成 0，避免后续建模报错。
            df[column] = 0
            category_mappings[column] = {"SingleValue": 0}
            continue

        if non_null_count == 2:
            df[column], mapping = encode_binary_series(df[column])
        else:
            df[column], mapping = encode_multiclass_series(df[column])
        category_mappings[column] = mapping

    return df, category_mappings


# ---------------------------------------------------------------------------
# 7. 主流程
# ---------------------------------------------------------------------------
def clean_data(read_path: Path = READ_PATH) -> tuple[pd.DataFrame, dict[str, dict[str, int]]]:
    """执行完整清洗流程。"""
    # 读取 xlsx；完整数据中如果工作表名称不同，默认读取第一个工作表。
    df = pd.read_excel(read_path)

    # 统一原始空值表示。
    df = df.applymap(normalize_missing)

    # 字段改名：只对当前存在的列执行映射，避免因完整数据列略有变化而报错。
    rename_map = {column: COLUMN_RENAME_MAP[column] for column in df.columns if column in COLUMN_RENAME_MAP}
    df = df.rename(columns=rename_map)

    # 正式处理前先审阅表格数据。
    review_raw_data(df)

    # 统计各字段缺失率，并删除高缺失字段。
    df, _ = report_and_drop_high_missing_features(df)

    # 计算住院时长，并删除原日期列。
    df = add_length_of_stay(df)

    # 删除明确要求丢弃的字段。
    drop_columns = [column for column in column_to_drop if column in df.columns]
    if drop_columns:
        df = df.drop(columns=drop_columns)

    # 数值列预处理（BMI 也按连续变量纳入）。
    df, numeric_columns = preprocess_numeric_columns(df)

    # 删除高缺失与指定字段后，输出分组统计表。
    stats_table = build_group_statistics_table(df, numeric_columns, target_column=TARGET_COLUMN)
    if not stats_table.empty:
        print("\n分组统计表：")
        print(stats_table.to_string(index=False))
        stats_table.to_csv(STAT_TABLE_PATH, sep="\t", index=False, encoding="utf-8-sig")
        print(f"分组统计表已保存：{STAT_TABLE_PATH}")

    # 按肺部感染分组处理缺失值，不进行归一化。
    df = split_by_target_and_handle_missing(df, numeric_columns)

    # 分类变量众数填补 + 编码。
    df, category_mappings = encode_categorical_columns(df, numeric_columns)
    return df, category_mappings


def save_outputs(df: pd.DataFrame, category_mappings: dict[str, dict[str, int]]) -> None:
    """保存清洗结果和编码字典。"""
    df.to_csv(WRITE_PATH, index=False, encoding="utf-8-sig")
    MAPPING_PATH.write_text(json.dumps(category_mappings, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    cleaned_df, mappings = clean_data()
    save_outputs(cleaned_df, mappings)
    print(f"数据清洗完成，已输出：{WRITE_PATH}")
    print(f"分类映射已保存：{MAPPING_PATH}")
    print(f"数据审阅摘要已保存：{REVIEW_PATH}")
