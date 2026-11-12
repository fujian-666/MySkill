# -*- coding: utf-8 -*-
"""KA 价差表「表型自动识别器」——跑任何核对/修复前必须先跑这个，杜绝战区/渠道/月份搞混。

用法:
    python identify_table.py "<xlsx或xls完整路径>"
    python identify_table.py "<路径>" --json

判定维度（四元组，缺一不可）:
    战区(长沙/南宁/...) × 渠道(五星/苏宁) × 月份(1-12) × 版本(初版/审核版/...)

输出: 表型代号、主表 sheet、表头行、数据行范围、型号列、关键列位、应选用的源表路径。
源表根目录与月份路径全部来自 config/profile.json，**绝不凭记忆拼路径**。
"""
import os, sys, re, json, collections
import openpyxl
from profile import load as load_profile

L = openpyxl.utils.get_column_letter
P = load_profile()
ROOT_PRICE = P.get("price_root") or ""
SRC_PATHS = {int(k): tuple(v) for k, v in (P.get("source_paths") or {}).items()}


def find_header_row(ws, max_scan=6):
    """找表头行：候选行里'非空单元格最多'且含关键词的那一行"""
    KEY = ("供价", "型号", "名称", "门店", "数量", "金额")
    best, bestscore = 1, -1
    for hr in range(1, max_scan + 1):
        n = 0
        hit = 0
        for c in range(1, min(ws.max_column, 60) + 1):
            v = ws.cell(hr, c).value
            if v is not None and str(v).strip():
                n += 1
                if any(k in str(v) for k in KEY):
                    hit += 1
        score = hit * 10 + n
        if score > bestscore:
            bestscore, best = score, hr
    return best


def find_col(ws, hrow, names, max_c=80):
    """按表头名定位列（精确优先，再按最短表头模糊）"""
    if isinstance(names, str):
        names = [names]
    for nm in names:
        for c in range(1, min(ws.max_column, max_c) + 1):
            v = ws.cell(hrow, c).value
            if v is not None and str(v).strip().replace("\n", "") == nm:
                return c
    for nm in names:
        best, blen = None, 10 ** 9
        for c in range(1, min(ws.max_column, max_c) + 1):
            v = ws.cell(hrow, c).value
            if v is None:
                continue
            s = str(v).replace("\n", "")
            if nm in s and len(s) < blen:
                best, blen = c, len(s)
        if best:
            return best
    return None


def identify(path):
    path = os.path.abspath(path)
    fn = os.path.basename(path)
    res = {"文件": path, "文件名": fn, "警告": []}

    # 1. 月份
    mon = None
    m = re.search(r"(\d{1,2})\s*月", fn)
    if m:
        mon = int(m.group(1))
    else:
        m = re.search(r"26年(\d{1,2})月", fn) or re.search(r"2026-(\d{1,2})", fn)
        if m:
            mon = int(m.group(1))
    res["月份"] = mon
    if mon is None:
        res["警告"].append("❌ 文件名里认不出月份，必须人工确认")

    # 2. 渠道
    if "五星" in fn:
        ch = "五星"
    elif "苏宁" in fn:
        ch = "苏宁"
    else:
        ch = None
        res["警告"].append("❌ 文件名里认不出渠道（五星/苏宁），必须人工确认")
    res["渠道"] = ch

    # 3. 战区（文件名优先，回退路径）
    fn_up = fn.upper()
    if "广西" in fn_up:
        zone = "广西"
    elif "南宁" in fn_up:
        zone = "南宁"
    elif "长沙" in fn_up:
        zone = "长沙"
    else:
        up = path.upper()
        if "广西" in up:
            zone = "广西"
        elif "南宁" in up:
            zone = "南宁"
        elif "长沙" in up:
            zone = "长沙"
        else:
            zone = None
            res["警告"].append("⚠️ 文件名/路径里没有'长沙/南宁/广西'，战区需人工确认")
    res["战区"] = zone

    # 4. 打开
    if not os.path.exists(path):
        res["警告"].append("❌ 文件不存在")
        return res
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        if path.lower().endswith(".xls"):
            res["警告"].append("🔴 是 .xls 老格式 → 先用 COM SaveAs(路径,51) 转 xlsx")
            res["sheets"] = "未知(需转换)"
            return res
        res["警告"].append(f"❌ 打开失败: {e}")
        return res

    res["sheets"] = wb.sheetnames

    # 5. 主表 / 表头 / 型号列 / 数据行
    main = None
    cands = [s for s in wb.sheetnames if "制表" in s]
    if cands:
        cands2 = [s for s in cands if not re.search(r"\(\s*2\s*\)", s)]
        main = (cands2 or cands)[0]
    if main is None:
        best, bn = None, 0
        for s in wb.sheetnames:
            try:
                n = wb[s].max_row
                if n > bn:
                    bn, best = n, s
            except Exception:
                pass
        main = best
        res["警告"].append("⚠️ 没找到名字含'制表'的 sheet，回退取行数最多者，请人工确认")
    res["主表"] = main
    ws = wb[main]

    hrow = find_header_row(ws)
    res["表头行"] = hrow
    mcol = find_col(ws, hrow, ["商品_名称", "商品名称", "名称", "型号"])
    res["型号列"] = L(mcol) if mcol else None
    res["型号列号"] = mcol

    last = hrow
    for r in range(hrow + 1, ws.max_row + 3):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, max(mcol or 1, 8) + 1)):
            last = r
    res["数据行"] = (hrow + 1, last)
    res["数据条数"] = last - hrow

    code = f"{zone or '?'}·{ch or '?'}·{mon or '?'}月"
    res["表型"] = code

    # 7. 关键列位
    keycols = ["渠道供价", "供价", "大区供价", "公司授权底价", "总部版授权底价",
               "总部最低零售限价", "付款金额", "数量", "无函", "确认费用", "争议费用",
               "战区版费用\n超授权", "总部版\n超授权", "最终超授权", "报批大区依据"]
    found = {}
    seen = set()
    for k in keycols:
        c = find_col(ws, hrow, [k, k.replace("\n", "")])
        if c and c not in seen:
            seen.add(c)
            found[k.replace("\n", "")] = L(c)
    res["关键列位"] = found

    # 8. 该用的源表（月份必须一致）
    if mon in SRC_PATHS:
        dq_rel, zb_rel = SRC_PATHS[mon]
        res["应选用源表"] = {
            "大区表": os.path.join(ROOT_PRICE, dq_rel),
            "总部表": os.path.join(ROOT_PRICE, zb_rel),
        }
        for k, p in res["应选用源表"].items():
            if not os.path.exists(p):
                res["警告"].append(f"❌ {k}路径不存在: {p}")
        res["应选用源表sheet"] = {"大区表": "连锁+区域连锁价格表", "总部表": "线下渠道"}
        res["源表列位"] = {"大区表": "F型号 / G连锁供价 / K当月零售价（双层表头）",
                           "总部表": "D型号 / F连锁供价 / G最低零售价（表头行4）"}
    else:
        res["警告"].append(f"⚠️ 没有登记 {mon} 月的源表路径，必须人工指定")

    # 9. 错误值快扫
    err = collections.Counter()
    for c in range(1, min(ws.max_column, 60) + 1):
        for r in range(1, min(ws.max_row, 3000) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("#"):
                err[L(c)] += 1
    res["主表错误值"] = dict(err) if err else "0"

    # 10. 交叉校验：文件名月份 vs 数据里的日期
    dc = find_col(ws, hrow, ["销售日期", "主时间", "日期"])
    if dc:
        months = collections.Counter()
        for r in range(hrow + 1, min(last, hrow + 500) + 1):
            v = ws.cell(r, dc).value
            s = str(v)
            m2 = re.search(r"(\d{4})-(\d{2})-", s) or re.search(r"(\d{4})/(\d{2})/", s)
            if m2:
                months[m2.group(2)] += 1
        if months:
            top = months.most_common(1)[0]
            res["数据实际月份"] = top[0]
            if mon and int(top[0]) != mon:
                res["警告"].append(
                    f"🔴 月份冲突！文件名说 {mon} 月，数据日期主要是 {top[0]} 月 —— 可能用错价格表")
    wb.close()
    return res


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    r = identify(sys.argv[1])
    if "--json" in sys.argv:
        print(json.dumps(r, ensure_ascii=False, indent=2))
        return
    print("=" * 72)
    print(f"文件: {r['文件名']}")
    print(f"表型: 【{r.get('表型')}】   主表: {r.get('主表')}")
    print(f"表头行: {r.get('表头行')}   数据行: {r.get('数据行')}   型号列: {r.get('型号列')}")
    print(f"主表错误值: {r.get('主表错误值')}")
    if r.get("数据实际月份"):
        print(f"数据实际月份: {r['数据实际月份']} 月")
    print("\n关键列位:")
    for k, v in (r.get("关键列位") or {}).items():
        print(f"   {v:<4} {k}")
    if r.get("应选用源表"):
        print("\n应选用源表（月份必须一致）:")
        for k, v in r["应选用源表"].items():
            print(f"   {k}: {v}")
        for k, v in (r.get("源表列位") or {}).items():
            print(f"   {k}列位: {v}")
    if r.get("警告"):
        print("\n⚠️ 警告:")
        for w in r["警告"]:
            print("   " + w)
    else:
        print("\n✅ 无警告，表型判定通过")
    print("=" * 72)


if __name__ == "__main__":
    main()
