# Excel / COM / openpyxl 硬坑清单

> 每一条都是实测踩出来的。写脚本前扫一遍，别重复交学费。

## 求末行 / 末列

- 🔴 禁 `End(-4162)`（`xlUp`）：AutoFilter 打开时会**漏 83 行**。
- ✅ `UsedRange.Row + UsedRange.Rows.Count - 1`，并与 openpyxl `max_row` 交叉校验。
- 脏 sheet 可能爆到 16000+ 列 / 9000+ 行 → 扫描限幅 `min(max_row, 3000)` / `min(max_column, 80)`。

## 写入

- 🔴 `Range.Value = 二维数组` **静默只写首行**（不报错）→ 必须**逐格写 + 逐格读回**。
- `.Value = None` 抛 "Property can not be set" → 清空用 `.ClearContents()`。
- COM 返回的行号是 `float` → `int()`。
- `RPC_E_CALL_REJECTED` → `sleep(1.2)` 后重试。
- `~$` 锁文件 / `wb.ReadOnly` = 用户正开着 Excel → **不杀窗口**，提示关闭后重试；有可见窗口的 EXCEL.EXE 一律保留。
- 失败路径**必须 `excel.Quit()`**，否则残留无窗口进程反而锁住文件，形成「只读打开」死循环。
- `DispatchEx` 开新实例，`Workbooks.Open(P, UpdateLinks=0)` 避免「是否更新链接」模态框。
- `wb.Save()` 可能**静默失败**（时间戳不变）→ 比对 `os.path.getmtime` before/after；改完用 openpyxl 重开读回。

## 计算模式

- `excel.Calculation = -4135` **必须在 `Workbooks.Open` 之后**（放前面报错）。
- 禁 `CalculateFullRebuild()`（会全簿重算、破坏基线）。
- 手动模式下**唯一有效的定向重算** = 逐格重写原 `.Formula`：
  ```python
  f = ws.Cells(r, c).Formula
  if isinstance(f, str) and f.startswith("="):
      ws.Cells(r, c).Formula = f
  ```
  `Range.Calculate` / `Dirty=True` / 切自动再切回**全部无效**；空格跳过，不硬断言为公式。
- 公式文本零改动，只让缓存值联动；门禁 G6 会逐格查。

## Sheet 操作

- `Sheets.Add(After=...)` 在 win32com 下**不可靠**（会落到倒数第几位）。
  ✅ `Add()` 后单参数 `cs.Move(Before=wb.Sheets(mainIdx+1))`。
- 任何 sheet 不能删行 / 改序号；需增删行才整表重建。
- 插列后**列号整体平移** → 按表头名重新定位；重复表头（多个「备注」）按出现次序取。
- 型号/金额列定位按**表头文字**，不按列号；找不到立刻报错停下，不猜、不退回默认值。

## openpyxl

- 🔴 **openpyxl 与 COM 不能同时持有同一文件**（含图片写回会崩）→ 只读建 plan，关闭后再开 COM。
- `read_only=True, data_only=True` 读值层，`data_only=False` 读公式层；两者配合做全簿 diff。
- 错误值判定：值层 `str` 以 `#` 开头且**不是** `####`（列宽不足）；公式层若以 `=` 开头说明是公式结果，不是被覆盖的文本。
- 文本格在公式层会随值呈现 → diff 判「公式层改动格 == 值层改动格 且 无真公式」才算纯文案改动。
- `enumerate(iter_rows(min_row=3), 3)` 才能对齐绝对行号。
- 日期比对：COM 返回 `pywintypes` 时区差 → 统一 `strftime` 归一，否则全是假差异。

## 样式

- COM `Interior.Color` 是 **BGR**：淡蓝 `#DDEBF7` → `0xF7EBDD`，淡黄 `#FFF2CC` → `0xCCF2FF`。
- 整块 `Range.Interior.Color = X` 在 COM 下可靠生效，但**写完必须读回校验**。
- 用 openpyxl 重写颜色会引入 theme 色歧义 → 样式统一由 `style_check_sheet.apply_standard(cs)` 在 COM 内套用。
- `Range.Characters(start, length)` 在动态绑定下报「找不到成员」→ 需手动 DISPID `Invoke` + `win32.Dispatch` 包装；`Borders(edge)` 可直接调。

## 公式与错误值

- 全簿凡含 `/` 或 `VLOOKUP` / `XLOOKUP` 且未包 `IFERROR` 的，按现行口径包 `=IFERROR(原式,0)`；
  **断链首选修法 = 重定向到本簿内嵌同名表**，不要无脑归 0。
- 错误值扫描**不能靠 COM**（错误格返回错误码、列宽不足返回 `####`）→ SaveAs 后用 openpyxl `data_only=True` 扫缓存值。
- 差额定值保护：`=IF(基准=0,0,被减数-基准)`；超授权列同步保护数量 ≤ 0 与基准 = 0。

## 环境

- Git Bash 的 `PATH` 会突然丢（`dirname/cd/grep/tail: command not found`）→ 每条命令开头
  `export PATH="/usr/bin:/bin:/c/Windows/System32"`，或**全程用绝对路径解释器**。
- 别把 PortableGit 的 bin 塞进 PATH（会让 `cp`/`tail` 解析到不可执行副本）。
- 别用 `| tail` 截断输出（可能不可执行），直接全量输出。
- Windows 下含中文的 `.vbs` 必须 UTF-16 LE with BOM + CRLF，否则 WSH 按 GBK 解码会吞掉换行。

## 校验器自身的坑

- audit 脚本返回 `(sheet,(r,c))` 还是 `(sheet,r,c)`，别混用。
- 回扫日期列 / 错误值列会报假异常 → 收窄到数值列，全簿 diff 交给独立终验。
- `t.index()` 定位段落必须在**所有替换完成后**重算，不能用替换前的偏移。
- 三元表达式优先级、字典嵌套层级、`rec` 的 msg 未 `str()` —— 都是实测炸过的点。
- 别为让校验变绿而放宽判据；**先坐实事实，再改判据，并由另一项校验兜住豁免**。
