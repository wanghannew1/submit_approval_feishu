# 工资表验证逻辑

## 验证体系总览

验证分三大块，在 `main()` 中按顺序执行：

```
上传文件 → 数据解析(parse_excel)
         │
         ├─ 1. 签名栏检查 (check_signatures)
         │     关键字存在性 + 字段值非空 (check_signer_fields)
         │
         ├─ 2. 工资表内容校验 (validate_payroll)
         │     A. 纵向列加总
         │     B. 合计行横向公式
         │     C. 每行横向公式
         │     D. 表格格式验证
         │     E. 必需列头检查
         │
         ├─ 3. 零金额防护 (main 内联)
         ├─ 4. 重复文件检测 (main 内联)
         │
         └─ 结果 → 「验证结果」sheet (汇总表) + UI 预览
```

---

## 配置结构 (`config.json → validation`)

```json
{
  "validation": {
    "enabled": true,
    "strict": true,
    "tolerance": 0.0,
    "write_back_sheet": false,
    "write_back_sheet_name": "验证结果",
    "column_sum_checks": [...],
    "row_formulas": [...],
    "required_signatures": [...],
    "required_signer_fields": [...],
    "required_headers": [...]
  }
}
```

| 配置项 | 用途 | 默认 |
|--------|------|------|
| `enabled` | 总开关 | `false` |
| `strict` | 严格模式：失败时阻止提交 | `true` |
| `tolerance` | 金额比较容差（元） | `0.00` |
| `write_back_sheet` | 校验结果是否写入汇总表 sheet | `true` |

---

## 1. 签名栏检查 (`check_signatures`)

### 1a. 关键字存在性

搜索整个 Excel 所有单元格，判断签名栏关键字是否存在。

**配置格式：**

```json
"required_signatures": [
  ["总经理签字", "部长签字|部长、分管副总签字|分管领导审核", "财务审核", "业务审核"],
  ["总经理签字", "部长签字|部长、分管副总签字|分管领导审核", "财务审核"]
]
```

- 外层数组的每个元素是一个 **分组**（任一组全部匹配即通过）
- 分组内的 `|` 表示 **别名**（任一别名匹配即视为该字段存在）

**搜索方式：**

```
全部单元格文本 → 空格拼接 → 转小写 → 子串匹配
```

对 .xls 用 `xlrd`，对 .xlsx 用 `openpyxl`。

### 1b. 字段值非空 (`check_signer_fields`)

检查某些签名栏字段（如「制表人」）填写了**实际内容**而非留空。

**配置：**

```json
"required_signer_fields": ["制表人"]
```

**检测逻辑：**

```
每个单元格 → 字段名是否在文本中 → 提取字段后的内容 → strip + lstrip("：: ")
                                        → 余下非空 → 通过
                                        → 余下为空 → 失败
```

**失败示例：**
| 单元格内容 | 结果 |
|-----------|------|
| `制表人：张三` | 通过 |
| `制表人：` | 失败（冒号后无内容） |
| `制表人` | 失败（无冒号无姓名） |

---

## 2. 工资表内容校验 (`validate_payroll`)

### A. 纵向列加总

**目的**：检查 Excel 中各列的数据行累加值是否等于合计行对应值。

**配置：**

```json
"column_sum_checks": [
  {"column": "deduction_total"},
  {"column": "personal_tax"},
  {"column": "net_total"}
]
```

**算法：**

```
col_sum = sum(数据行[该列])
summary_val = 合计行[该列]
diff = col_sum - summary_val
passed = |diff| ≤ tolerance
```

**失败条件**：`|sum(数据行) - 合计行值| > tolerance`

---

### B. 合计行横向公式

**目的**：检查合计行是否满足 `lhs = sum(rhs_plus) − sum(rhs_minus)`。

**配置：**

```json
"row_formulas": [
  {
    "name": "转账合计 = 扣款合计 + ... + 实发合计 − 交纳现金",
    "lhs": "transfer_total",
    "rhs_plus": ["deduction_total", "personal_tax", "net_total"],
    "rhs_minus": ["pay_cash"]
  }
]
```

**算法：**

```
L = 合计行[lhs]
R = sum(合计行[rhs_plus]) − sum(合计行[rhs_minus])
diff = L − R
passed = |diff| ≤ tolerance
```

---

### C. 每行横向公式

**目的**：对每个数据行执行与 B 节相同的公式检查。

**算法：** 逐行计算 `lhs` 与 `rhs_plus − rhs_minus` 的差值，记录所有失败行。

**结果统计：**
- 全部通过 → `"N/N 行通过"`
- 有失败 → `"N/N 行通过；失败例: 第X行 差 +Y.YY; ..."`

---

### D. 表格格式验证

**目的**：捕获配置与实际 Excel 不匹配的场景（即使 A/B/C 因全零巧合通过）。

**第零步 — 合计行缺失（根因优先）：**

当 `parse_excel()` 未找到首列为「合计」的汇总行时，会标记 `no_summary_row=True` 并返回兜底结果（`column_indices` 为空、`all_header_texts` 为空集合）。此时 D 项**直接报合计行缺失**，不走下面的关键列存在性检查——因为此时关键列 index 全为 -1 是连锁反应，按「未找到关键列」报会把用户引向错误的表头配置方向。

失败提示：

```
未找到合计行：工资表中没有首列为「合计」的汇总行。
请确认最后一行明细数据之后、签字行之前存在合计行，
且合计行首列单元格内容为「合计」；若文件确实缺少合计行，请联系管理员调整配置。
```

（该提示使用配置的 `summary_row_marker`，如配置为「合计」则显示「合计」。）

**第一步 — 关键列存在性：**

检查 `transfer_total`、`deduction_total`、`net_total` 三列是否都在表头中找到。

```
missing_cols = [列 if column_indices[列] < 0 for 列 in critical_cols]
```

失败时提示用户检查 `header_start_row`/`header_row_count` 配置。

**第二步 — 合计行非零：**

如果三列都找到了但合计行三个值全是 0，且存在明细数据行：

```
has_real_data = any(合计行[列] != 0 for 列 in critical_cols)
passed = has_real_data or 无数据行
```

失败时提示「识别到的金额均为 0，但存在 N 行明细数据」。

---

### E. 必需列头检查

**目的**：确保工资表包含指定的表头列。

**配置：**

```json
"required_headers": [
  "序号", "姓名", "应发工资",
  "转账合计", "扣款合计", "实发工资", "实发合计"
]
```

**算法：**

```
all_header_texts = set(所有表头区域单元格文本)
for 必需列头 in required_headers:
    found = any(必需列头 in t for t in all_header_texts)
    未找到 → 加入 missing_required 列表

passed = len(missing_required) == 0
```

**搜索范围：** 配置的列头区域 + 宽范围兜底（标题行后到合计行前）。

**合计行缺失时（`no_summary_row=True`）：** 直接跳过不判失败。此时 `all_header_texts` 为空集合，所有必需列头必然「缺失」，但那只是合计行缺失的连锁反应（主因已由 D 项说明），跳过以免重复误导。

**失败示例：**
- 缺少「序号」→ `缺少必需列头：序号`
- 缺少「序号、姓名」→ `缺少必需列头：序号、姓名`
- 同时缺少全部 7 列 → `缺少必需列头：序号、姓名、应发工资、转账合计、扣款合计、实发工资、实发合计`

---

## 3. 零金额防护

**位置**：`main()` 中，提交按钮前。

**条件：**

```python
transfer_total in ("0.00", "0", "", None)
AND
net_total in ("0.00", "0", "", None)
```

任一文件的转账合计和实发合计同时为零 → 显示红色错误，禁止提交。

**合计行缺失（`no_summary_row=True`）时**：同样触发该防护，但提示改为说明真实原因——

```
未找到合计行——文件中未检测到首列为「合计」的汇总行，
因此无法读取转账合计、实发合计。
请确认最后一行明细数据之后、签字行之前存在合计行，且合计行首列单元格内容为「合计」；
若文件确实缺少合计行，请联系管理员调整配置。
```

---

## 4. 重复文件检测

**位置**：`main()` 中，零金额防护之后。

**条件：**

```python
same_unit      = unit_name 相同
same_ym        = year_month 相同
same_trans     = transfer_total 相同
same_deduct    = deduction_total 相同
same_net       = net_total 相同
```

五个条件**全部满足** → 视为重复文件，显示红色错误，禁止提交。

---

## 5. 控制参数

| 参数 | 用途 | 影响 |
|------|------|------|
| `strict: true` | 严格模式 | 校验失败 → 提交按钮 disabled |
| `strict: false` | 宽松模式 | 校验失败 → 黄色警告，仍可提交 |
| `tolerance: 0.00` | 容差（元） | A/B/C 节的 |diff| ≤ tolerance 算通过 |

---

## 6. 验证结果结构

`validate_payroll()` 返回：

```python
{
    "ok": True/False,          # 全部通过？
    "passed_count": N,         # 通过的检查项数
    "failed_count": N,         # 失败的检查项数
    "checks": [
        {"kind": "column_sum", "name": "...", "passed": True/False, "detail": "..."},
        {"kind": "row_formula_summary", ...},
        {"kind": "row_formula_rows", ...},
        {"kind": "table_format", ...},
        {"kind": "required_headers", ...},
    ]
}
```

`check_signatures()` 返回：

```python
{
    "ok": True/False,          # 签名全部存在？
    "found": ["总经理签字", ...],
    "missing": ["业务审核"],
}
```

`check_signer_fields()` 返回：

```python
{
    "ok": True/False,
    "empty_fields": ["制表人"],  # 未填写姓名字段
}
```

---

## 7. 用户界面提示

| 场景 | 提示类型 | 示例 |
|------|---------|------|
| A/B/C 校验失败 | `st.error` | `❌ 扣款合计列加总 123.45 与合计行 120.00 相差 +3.45` |
| D 关键列缺失 | `st.error` | `❌ 未在表头中找到关键列：转账合计（元）。请确认工资表表头起始行为第 3 行…` |
| D 合计全零 | `st.error` | `❌ 识别到的转账合计、扣款合计、实发合计均为 0，但存在 15 行明细数据…` |
| E 必需列缺失 | `st.error` | `❌ 缺少必需列头：序号、姓名` |
| 签名关键字缺失 | `st.error` | `❌ 缺少签名栏信息 — 业务审核` |
| 制表人未填姓名 | `st.error` | `❌ 签名栏「制表人」未填写姓名` |
| 零金额 | `st.error` | `❌ 识别到的转账合计和实发合计均为 0…` |
| 重复文件 | `st.error` | `❌ 文件A与文件B为重复上传…` |
| 全部通过 | 无 / 汇总表 | 汇总表「验证明细」sheet 中逐项标注 ✅ |
