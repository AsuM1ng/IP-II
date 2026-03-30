# IP-IncisionInfection

## 保留原始数据中与清洗表对应记录（切口感染）

新增脚本：`filter_raw_by_clean_and_build_label_dict.py`

### 功能
- 读取 `原始数据.xlsx` 与 `数据_清洗后_切口感染.xlsx`。
- 先做全字段精确匹配，再按 **BMI + 术前血红蛋白HGB + 术前前白蛋白PALB + 术前白蛋白ALB** 优先匹配。
- 仅保留在清洗表中存在的原始记录，并输出：
  - `原始数据_按清洗表筛选_切口感染.xlsx`（保留原始单元格样式/日期格式）
  - `原始数据_按清洗表筛选_切口感染.csv`
  - `未匹配_清洗数据_切口感染.xlsx`
- 统计各标签字段中，清洗表数值编码与原始文本的对应关系，输出：
  - `标签编号_原文映射字典.json`

### 运行
```bash
python filter_raw_by_clean_and_build_label_dict.py
```
