# -*- coding: utf-8 -*-
"""业务在「核对情况表」G(业务填写原因)/H(OA号或价格依据) 填完后，回填到「制表」对应行。

规则：制表里的不一致（没找到价格 / 价格有差异）必须和核对情况表一致；
业务填写原因后，要回填到制表里面 —— 但**不覆盖原核对标注**，幂等可重复跑。

落位：
  供价核对         -> 制表 K 备注
  公司授权底价核对 -> 制表 U 备注
  总部底价核对     -> 制表 AE 备注
追加段格式：｜业务反馈：{G}；依据：{H}   （重复执行自动替换旧段，不叠加）

用法：
  python backfill_reason_to_tz.py --file "<源xlsx>"                       # 预览（不改文件）
  python backfill_reason_to_tz.py --file "<源xlsx>" --out "<新xlsx>" --apply
"""
import sys, os, argparse
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook
from profile import load as load_profile

if "--profile" in sys.argv:
    _i = sys.argv.index("--profile")
    if _i + 1 < len(sys.argv):
        os.environ["KA_PROFILE"] = sys.argv[_i + 1]

P = load_profile()
MAIN, CHECK = P["sheets"]["main"], P["sheets"]["check"]
MARK = P["backfill_mark"]
TYPE2COL = P["type2remark_col"]


def build_plan(src):
    wb = load_workbook(src, read_only=True, data_only=True)
    TZ = list(wb[MAIN].iter_rows(min_row=1, max_row=P["read_max_row"],
                                 max_col=P["read_max_col"], values_only=True))
    CS = list(wb[CHECK].iter_rows(min_row=1, max_row=P["check_max_row"],
                                  max_col=P["check_max_col"], values_only=True))
    def t(r, c):
        return TZ[r - 1][c - 1]
    filled = []
    for i, row in enumerate(CS, 1):
        if isinstance(row[0], (int, float)) and row[1]:
            g, h = str(row[6] or "").strip(), str(row[7] or "").strip()
            if g or h:
                filled.append(dict(cs_row=i, no=int(row[0]), typ=str(row[1]),
                                   tz_row=int(row[2]), g=g, h=h))
    plan, errors = [], []
    for x in filled:
        r, c = x["tz_row"], TYPE2COL[x["typ"]]
        old = str(t(r, c) or "").strip()
        if old.startswith("="):
            errors.append("制表 r%d c%d 是公式，禁止覆盖" % (r, c))
            continue
        base = old.split("｜" + MARK)[0].strip()
        seg = "｜%s：%s" % (MARK, x["g"] or "未填写")
        if x["h"]:
            seg += "；依据：" + x["h"]
        new = (base + " " + seg).strip()
        plan.append(dict(cs_row=x["cs_row"], no=x["no"], typ=x["typ"], r=r, c=c,
                         old=old, new=new, g=x["g"], h=x["h"]))
    wb.close()
    # 公式层再确认目标格不是公式
    wbf = load_workbook(src, read_only=True, data_only=False)
    FZ = list(wbf[MAIN].iter_rows(min_row=1, max_row=P["read_max_row"],
                                  max_col=P["read_max_col"], values_only=True))
    for p in plan[:]:
        v = FZ[p["r"] - 1][p["c"] - 1]
        if isinstance(v, str) and v.startswith("="):
            errors.append("制表 r%d c%d 公式层为公式，禁止覆盖" % (p["r"], p["c"]))
    wbf.close()
    return filled, plan, errors


def apply(src, out, plan):
    import win32com.client as win32
    if os.path.exists(out):
        raise SystemExit("输出文件已存在，不覆盖：" + out)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(src, UpdateLinks=0, ReadOnly=False)
        excel.Calculation = -4135  # 必须在 Open 之后；禁 FullRebuild
        if wb.ReadOnly:
            wb.Close(SaveChanges=False)
            raise SystemExit("文件被 Excel 占用（ReadOnly），请先关闭再重试")
        ws = wb.Worksheets(MAIN)
        for p in plan:
            ws.Cells(p["r"], p["c"]).Value = p["new"]
            got = ws.Cells(p["r"], p["c"]).Value
            if str(got or "").strip() != p["new"]:
                raise RuntimeError("读回不一致 r%d c%d: %r" % (p["r"], p["c"], got))
        wb.SaveAs(out, FileFormat=51)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    from check_tz_cs_consistency import run as consistency
    rc = consistency(out, profile=P["_profile_file"])
    if rc != 0:
        raise SystemExit("回填后一致性校验未通过，请检查：" + out)
    print("回填完成并复验通过:", out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--out")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--profile")
    a = ap.parse_args()
    filled, plan, errors = build_plan(a.file)
    print("已填 G/H 的条目:", len(filled), " 待回填格:", len(plan), " 阻断错误:", len(errors))
    for e in errors:
        print("  [ERR]", e)
    for p in plan[:15]:
        print("  情况表r%-3d 序%-3d %-12s -> 制表 r%-4d c%-2d\n      原: %s\n      新: %s" %
              (p["cs_row"], p["no"], p["typ"], p["r"], p["c"], p["old"][:60], p["new"][:110]))
    if len(plan) > 15:
        print("  ... 其余 %d 格" % (len(plan) - 15))
    if not plan:
        print("G/H 全空，无可回填内容（规则已固化，业务填完再跑本脚本即可）")
    if errors:
        raise SystemExit(1)
    if a.apply:
        if not a.out:
            raise SystemExit("--apply 必须指定 --out")
        if not plan:
            raise SystemExit("无回填内容，不生成空转文件")
        apply(a.file, a.out, plan)
    else:
        print("（预览模式，未改文件；加 --out <新文件> --apply 才落盘）")
