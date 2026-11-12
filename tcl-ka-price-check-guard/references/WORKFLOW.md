# 执行层流程（Guard 之外的部分）

> 本 skill 不负责插列/刷值，但门禁要校验执行层的产物，所以这里把口径写清楚。
> 列位一律**按表头名定位**，不硬编码列号；插列后**重新读表头**。

## 1. 五步核对

| 步 | 动作 | 取数 | 落地 |
|---|---|---|---|
| ① | H 渠道供价后插 3 列 | 大区表 G 连锁供价 | I 大区供价 / J 差额 / K 备注 |
| ② | 付款数量核对 | 明细 sheet（型A 键 = C 列单号 + 型号） | 只报差异，不覆盖原值 |
| ③ | 无函 = 确认费用 + 争议费用 | 表内恒等式 | 校验，不落列 |
| ④ | R 公司授权底价后插 3 列 | 大区表 K 当月零售价 | S 大区零售价 / T 差额 / U 备注 |
| ⑤ | AD 总部最低零售限价刷值 | 总部表 G → 淘汰机现净供价兜底 | AD + AE 备注 |

- 差额用**真公式** + 零基准保护：`=IF(S=0,0,R-S)`；不许写死值。
- 填色：新插列淡蓝 `#DDEBF7`，刷值列淡黄 `#FFF2CC`；COM `Interior.Color` 是 **BGR**。
- 已匹配的值不覆盖；未匹配填 0 + 备注，不跨渠道硬凑。
- 多台拆单按单号汇总；同单拆多行别逐行报错。
- 第三步恒等式 `X+Y=Q+P` 恒成立，**先验恒等式再造异常**；只比 Q 必出假异常（差恰 = P）。

## 2. 金额链

```
AD → AG(33) → AH(34) → 汇总表 K / L / C27 / C28
C26 = 孰低值
```

- 改 AD 后**手动计算模式下唯一有效的定向重算** = 逐格重写原 `.Formula`：
  ```python
  for r, c in RECALC:
      f = ws.Cells(r, c).Formula
      if isinstance(f, str) and f.startswith("="):
          ws.Cells(r, c).Formula = f
  ```
  `Range.Calculate` / `Dirty=True` / 切自动再切回，**全部无效**；空格跳过不硬断言。
- `excel.Calculation = -4135` **必须在 `Workbooks.Open` 之后**；禁 `CalculateFullRebuild()`。
- 公式文本零改动，只有缓存值随 AD 联动；门禁 G3/G6 会查。

## 3. 核对情况表结构

- 位置：紧邻主表之后，sheet 标签染红；标准样式 `style_check_sheet.apply_standard(cs)`（幂等）。
- 8 列：`序号 | 异常类型 | 行号 | 门店 | 型号 | 异常说明 | 业务填写原因 | OA号或价格依据`。
- 末两列黄底留空给业务；F 列异常说明写清对账价。
- 标题/表头红底白字或按现行定版样式；结论行 A:H 合并。
- 情况表 G/H 在业务填写前**保持为空**。

**整表重建**（需增删行时唯一合法做法）：
```
读条目 → Delete → Add() → 单参数 Move(Before=主表后首个sheet) → 回填 → apply_standard
```
禁表内删行 / 改序号 / 改结论；`Move(After=)` 不可靠。

## 4. 逐格写 + 读回（COM）

```python
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False; excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(src, UpdateLinks=0, ReadOnly=False)
    excel.Calculation = -4135
    if wb.ReadOnly:
        wb.Close(SaveChanges=False)
        raise RuntimeError("文件被 Excel 占用，请关闭后重试")
    ws = wb.Worksheets("制表")
    for r, c, val in plan:
        ws.Cells(r, c).Value = val
        assert str(ws.Cells(r, c).Value or "").strip() == str(val)   # 逐格读回
    wb.SaveAs(out, FileFormat=51)      # 另存新版本，不覆盖
    wb.Close(SaveChanges=False)
finally:
    excel.Quit()                       # 失败路径也必须 Quit，否则残留进程锁文件
```

- 🔴 `Range.Value = 二维数组` 会**静默只写首行**，必须逐格写。
- `.Value = None` 报错；清空用 `.ClearContents()`。
- COM 行号是 float，用 `int()`。
- `RPC_E_CALL_REJECTED` → `sleep(1.2)` 重试。
- `~$` 锁文件 / ReadOnly = 用户正开着 Excel，**不杀窗口**，提示关闭后重试。

## 5. plan.json 契约

门禁 G6 用 plan 断言「差异格集合 == 预期」：

```json
{
  "plan": [
    {"sheet": "制表", "r": 26, "c": 11, "old": "...", "new": "供价不对"},
    {"sheet": "核对情况表", "r": 104, "c": 1, "old": "核对结果：", "new": ""}
  ]
}
```

- 用 openpyxl **只读**建 plan，先断言旧值；建完关闭 workbook 再开 COM（两者不能同时持有，含图片写回会崩）。
- 改完由**独立脚本**做全簿逐格 diff，值层差异格数 = `--expect-diff`，且集合 == plan。

## 6. 回填（业务填完 G/H 之后）

```
预览： python scripts/backfill_reason_to_tz.py --file "<源>"
落盘： ... --out "<新版本>" --apply
```

- 供价核对 → K、公司授权底价核对 → U、总部底价核对 → AE。
- 追加 `｜业务反馈：{G}；依据：{H}`，先剥旧段再写，**幂等**。
- 目标格是公式 → 阻断，不覆盖。
- 另存新文件，落盘后自动复跑 10 项一致性；不过就报错。
- G/H 全空 → 明确输出「无可回填内容」，不生成空转文件。

## 7. 交付前自检顺序

1. `model_key.py` 自测 PASS
2. `identify_table.py` 无冲突
3. 独立 verify 全簿 diff（值层 + 公式层）
4. `check_tz_cs_consistency.py` 10 项全过
5. `delivery_gate.py --prev ... --expect-diff N --plan ...` 21 项全过
6. 拿上一版做**反向验证**，确认门禁能拦住已知缺陷
7. 记忆/日志 → 交付
