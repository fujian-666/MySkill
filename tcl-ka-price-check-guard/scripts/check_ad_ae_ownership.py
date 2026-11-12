# -*- coding: utf-8 -*-
"""第五步列归属硬断言（用户 11-12 14:44 纠正：「AD不是你自己要填的吗？不是业务填的，
也不知道总部价格」「AD填完，AE备注」，别再把 AD 说成业务列）。

只读，绝不改文件。核心口径：
  AD(总部最低零售限价) = 核对方按【总部线下渠道G → OA到手价 → 淘汰机现净供价】自行填列，
     业务原表该列只是空壳：358 行 `=O{r}`(直接引用公司授权底价) + 354 行空白 + 10 行外链
     XLOOKUP，业务既不填、也不掌握总部价。
  AE(备注) = AD 填完之后紧跟其右的备注列，写取数依据/证据/短标注；先有 AD 才有 AE。

断言（任何一条不过即退出码 1）：
  1 AD 全部为数值常数、真公式 0 行（业务原表公式已被核对方刷数覆盖）
  2 AD 与业务「公司授权底价」R 不全等（不许照抄业务底价）
  3 AD 与业务自带「核对」列(外链 VLOOKUP) 不全等（不许盲从业务外链）
  4 AE 列位 == AD+1 且表头行写「备注」
  5 AE 零真公式（全为文本/空白）
  6 AE 非空行的 AD 必须全部有值（AD 先填、AE 后备注）
  7 AE 空白行的 AD 也必须有值（总部线下G直取按口径留空，不是漏填）
  8 AD 无空白、AE 非空行文案类型只落在 {OA证据, 短标注}

用法: python check_ad_ae_ownership.py "<xlsx>" [--profile <profile.json>]
"""
import sys, json, os, collections
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook

DEF = dict(main="制表", header_row=2, first=3, last=810,
           ad=30, ae=31, r=18, biz_check=32, read_max_row=812, read_max_col=36)
R = []


def rec(name, ok, detail=""):
    R.append((name, bool(ok), detail))
    print("  [%s] %-46s %s" % ("OK" if ok else "FAIL", name, detail if detail else ""))


def grab(ws, maxrow, maxcol):
    return list(ws.iter_rows(min_row=1, max_row=maxrow, max_col=maxcol, values_only=True))


def main(path, profile=None):
    P = dict(DEF)
    # 通用 skill 包有 profile.py（模板+本地配置深合并）；项目工作区无此模块，走裸 json
    try:
        from profile import load as load_profile
        J = load_profile(profile)
    except ImportError:
        J = json.load(open(profile, encoding="utf-8")) if profile else {}
    if J:
        C = J.get("cols", {})
        P.update(main=J.get("sheets", {}).get("main", DEF["main"]),
                 header_row=J.get("header_row", 2),
                 first=J.get("first_row", 3), last=J.get("last_row", 810),
                 ad=C.get("AD", 30), ae=C.get("AE", 31), r=C.get("R", 18),
                 biz_check=C.get("biz_check", C.get("AF", C.get("AE", 31) + 1)),
                 read_max_row=J.get("read_max_row", 812),
                 read_max_col=J.get("read_max_col", 36))
    wv = load_workbook(path, read_only=True, data_only=True)
    wf = load_workbook(path, read_only=True, data_only=False)
    try:
        V = grab(wv[P["main"]], P["read_max_row"], P["read_max_col"])
        F = grab(wf[P["main"]], P["read_max_row"], P["read_max_col"])
    finally:
        wv.close(); wf.close()

    def v(r, c): return V[r - 1][c - 1]
    def f(r, c): return F[r - 1][c - 1]
    def isf(x): return isinstance(x, str) and x.startswith("=")
    def num(x):
        try: return float(x)
        except Exception: return 0.0
    ROWS = range(P["first"], P["last"] + 1)
    ad, ae, rcol, bcol = P["ad"], P["ae"], P["r"], P["biz_check"]

    print("文件:", os.path.basename(path))
    print("列位: AD=%d AE=%d 公司授权底价=%d 业务自带核对=%d  数据行 %d~%d"
          % (ad, ae, rcol, bcol, P["first"], P["last"]))

    ad_num = [r for r in ROWS if isinstance(v(r, ad), (int, float))]
    ad_formula = [r for r in ROWS if isf(f(r, ad))]
    ad_blank = [r for r in ROWS if v(r, ad) in (None, "")
                or (isinstance(v(r, ad), str) and not v(r, ad).strip())]
    rec("1 AD为核对方刷入的数值常数(808行/零公式)",
        len(ad_num) == len(ROWS) and not ad_formula and not ad_blank,
        dict(数值=len(ad_num), 公式=ad_formula[:5], 空白=ad_blank[:5]))

    diff_r = [r for r in ROWS if abs(num(v(r, ad)) - num(v(r, rcol))) > 0.005]
    rec("2 AD不照抄业务「公司授权底价」(不同行数>0)", len(diff_r) > 0,
        "与R不同 %d 行 / 共 %d 行" % (len(diff_r), len(ROWS)))

    diff_b = [r for r in ROWS if abs(num(v(r, ad)) - num(v(r, bcol))) > 0.005]
    rec("3 AD不盲从业务自带「核对」外链(不同行数>0)", len(diff_b) > 0,
        "与业务核对列不同 %d 行" % len(diff_b))

    head_ae = v(P["header_row"], ae)
    head_ad = v(P["header_row"], ad)
    rec("4 AE紧邻AD右侧且表头为「备注」",
        ae == ad + 1 and str(head_ae or "").strip() == "备注"
        and "总部最低零售限价" in str(head_ad or ""),
        "AD表头=%r AE表头=%r 列位差=%d" % (head_ad, head_ae, ae - ad))

    ae_formula = [r for r in ROWS if isf(f(r, ae))]
    ae_num = [r for r in ROWS if isinstance(v(r, ae), (int, float))]
    rec("5 AE为备注文本(零公式/零数值)", not ae_formula and not ae_num,
        dict(公式=ae_formula[:5], 数值=ae_num[:5]))

    ae_nonblank = [r for r in ROWS if str(v(r, ae) or "").strip()]
    ae_blank = [r for r in ROWS if not str(v(r, ae) or "").strip()]
    rec("6 AE非空行的AD全部已填(先AD后AE)",
        all(v(r, ad) not in (None, "") for r in ae_nonblank),
        "AE非空 %d 行, 其中AD空 %d 行" %
        (len(ae_nonblank), sum(1 for r in ae_nonblank if v(r, ad) in (None, ""))))
    rec("7 AE空白行的AD也全部已填(直取留空非漏填)",
        all(v(r, ad) not in (None, "") for r in ae_blank),
        "AE空白 %d 行, 其中AD空 %d 行" %
        (len(ae_blank), sum(1 for r in ae_blank if v(r, ad) in (None, ""))))

    kind = collections.Counter()
    for r in ae_nonblank:
        s = str(v(r, ae)).strip()
        kind["OA证据" if s.startswith(("OA《", "按OA《", "(OA号")) else "短标注"] += 1
    rec("8 AE备注只落在OA证据/短标注两类", set(kind) <= {"OA证据", "短标注"},
        dict(分类=dict(kind), 合计=len(ae_nonblank)))

    print("-" * 82)
    print("口径: AD 是【核对方】按总部线下G→OA到手价→淘汰机现净供价自行填列，业务不填、"
          "也不掌握总部价；AE 是 AD 填完后紧随其右的备注列。")
    bad = [x[0] for x in R if not x[1]]
    print("断言 %d 项：PASS %d / FAIL %d" % (len(R), len(R) - len(bad), len(bad)))
    if bad: print("未通过:", bad)
    print("=" * 82)
    return 1 if bad else 0


if __name__ == "__main__":
    if len(sys.argv) < 2: raise SystemExit(__doc__)
    pf = sys.argv[sys.argv.index("--profile") + 1] if "--profile" in sys.argv else None
    sys.exit(main(sys.argv[1], pf))
