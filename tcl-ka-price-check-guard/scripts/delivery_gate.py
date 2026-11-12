# -*- coding: utf-8 -*-
"""一键交付门禁（把散落在记忆里的硬门槛串成一条命令，不全过禁止交付）。

用法:
  python delivery_gate.py "<xlsx>"                                    # 单文件全量门禁
  python delivery_gate.py "<xlsx>" --prev "<上一版>" --expect-diff 86 [--plan <plan.json>]
  python delivery_gate.py --next "<任一现有xlsx或所在目录>"           # 只输出当天下一版应叫的名字
  python delivery_gate.py "<xlsx>" --profile "<profile.json>"        # 指定另一四元组

检查项:
  G0 型号归一化模块自测（后缀/连字符/等级机三轮漏配坑）
  G1 命名合规：当天第N版从V1连续、无缺口、不沿用前一天链、存盘目录正确
  G2 表型识别：identify_table.py 确认四元组，冲突即停
  G3 基线保真：sheet/图片/错误值/真公式/行数/金额链/AH=C27/C28/C26
  G4 13项双向一致性（check_tz_cs_consistency.py）
  G5 业务视图：结论区空、K/U/AE 内部出处归零、三列专属词、U一致短词、条目数、G/H空、OA证据行、AD无空白
  G5 第五步列归属（check_ad_ae_ownership.py）：AD 是【核对方】自行填列（业务不填、也不掌握总部价），
     AE 是 AD 填完后的备注列；严禁把 AD 说成业务填的列
  G6 全簿diff（给 --prev 时）：格数=plan、数值公式列零改动、真公式文本零改动

🔴 四元组不同 = 另一张表。所有基线/列位/口径都来自 config/profile.json，绝不套用别的表。
"""
import sys
import os
import re
import json
import zipfile
import subprocess
import argparse
import collections
import datetime

sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook
from profile import load as load_profile, rows as profile_rows

# --profile 必须在模块级派生常量（NAME_RE / ROWS / IDENTIFY ...）之前生效
if "--profile" in sys.argv:
    _i = sys.argv.index("--profile")
    if _i + 1 < len(sys.argv):
        os.environ["KA_PROFILE"] = sys.argv[_i + 1]

P = load_profile()
B = P["baseline"]
C, W = P["cols"], P["words"]
MAIN, CHECK, SUMMARY = P["sheets"]["main"], P["sheets"]["check"], P["sheets"]["summary"]
ROWS = profile_rows(P)
IDENTIFY = P["identify_script"]
NAME_RE = re.compile(r"^(?P<stem>.+?)—%s(?P<m>\d{1,2})-(?P<d>\d{1,2})核对V(?P<n>\d+)\.xlsx$"
                     % re.escape(P["owner"]))
KW = P["internal_keywords"]
NUMCOLS = P["num_cols"]
CONCL_FROM, CONCL_TO = P["conclusion_rows"]

R = []


def rec(name, ok, detail=""):
    R.append((name, bool(ok), detail))
    print("  [%s] %-38s %s" % ("OK" if ok else "FAIL", name, detail if detail else ""))


def snap(path, data_only):
    w = load_workbook(path, read_only=True, data_only=data_only)
    D = {}
    for sn in w.sheetnames:
        for i, row in enumerate(w[sn].iter_rows(values_only=True), 1):
            for c, v in enumerate(row, 1):
                if v is not None:
                    D[(sn, i, c)] = v
    sh = w.sheetnames
    w.close()
    return D, sh


def media_count(path):
    with zipfile.ZipFile(path) as z:
        return len([n for n in z.namelist() if n.startswith("xl/media/")])


def versions(directory, stem, md):
    out = {}
    for f in os.listdir(directory):
        if f.startswith("~$"):        # Excel 锁文件不能算版本 / 不能当 stem
            continue
        m = NAME_RE.match(f)
        if m and m.group("stem") == stem and (int(m.group("m")), int(m.group("d"))) == md:
            out[int(m.group("n"))] = f
    return out


def suggest_next(path):
    d = path if os.path.isdir(path) else os.path.dirname(os.path.abspath(path))
    stems = set()
    for f in os.listdir(d):
        if f.startswith("~$"):        # 目录里有 ~$ 锁文件时，别把 "~$原名" 当成表名
            continue
        m = NAME_RE.match(f)
        if m:
            stems.add(m.group("stem"))
    if not stems:
        raise SystemExit("目录里没有符合「原名—%sM-D核对VN.xlsx」的文件，无法推断表名：%s"
                         % (P["owner"], d))
    stem = sorted(stems)[0]
    today = datetime.date.today()
    md = (today.month, today.day)
    nums = versions(d, stem, md)
    n = max(nums) + 1 if nums else 1
    name = "%s—%s%d-%d核对V%d.xlsx" % (stem, P["owner"], md[0], md[1], n)
    print("当天下一版应为:", os.path.join(d, name))
    print("依据: 今天(%d-%d)已有 %s → V%d"
          % (md[0], md[1], sorted(nums) or "无（跨日重新从V1起算）", n))
    return os.path.join(d, name)


def gate(args):
    path = os.path.abspath(args.xlsx)
    if not os.path.exists(path):
        raise SystemExit("文件不存在: " + path)
    print("=" * 82)
    print("交付门禁:", os.path.basename(path))
    print("四元组 / 配置:", P["quad"], "/", P["_profile_file"])
    print("=" * 82)

    # G0 归一化模块
    try:
        from model_key import selftest
        rec("G0 型号归一化自测", selftest(), "后缀/连字符/等级机")
    except Exception as e:
        rec("G0 型号归一化自测", False, repr(e))

    # G1 命名合规
    m = NAME_RE.match(os.path.basename(path))
    if not m:
        rec("G1 命名格式", False, "不符合 原名—%sM-D核对VN.xlsx" % P["owner"])
    else:
        d = os.path.dirname(path)
        md = (int(m.group("m")), int(m.group("d")))
        vs = versions(d, m.group("stem"), md)
        nums = sorted(vs)
        continuous = nums == list(range(1, max(nums) + 1)) if nums else False
        rec("G1 同日版本从V1连续无缺口", continuous, "同日版本=%s" % nums)
        rec("G1 存盘目录", ("价差表" in path) and os.path.basename(d).endswith("月"), d)
        today = datetime.date.today()
        if md != (today.month, today.day):
            print("       [WARN] 文件名日期 %d-%d ≠ 今天 %d-%d（跨零点执行请确认归属日）"
                  % (md[0], md[1], today.month, today.day))
        if int(m.group("n")) != max(nums):
            print("       [WARN] 本次复核的不是当天最新版（当前当天最大 V%d）" % max(nums))
        rec("G1 未沿用前一天版本链", int(m.group("n")) <= len(nums) and continuous,
            "V%s / 当天共%d版" % (m.group("n"), len(nums)))

    # G2 表型识别
    if os.path.exists(IDENTIFY):
        r = subprocess.run([sys.executable, IDENTIFY, path], capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=300)
        out = r.stdout or ""
        ok = (r.returncode == 0 and "表型:" in out and P["quad"] in out
              and ("主表: %s" % MAIN) in out
              and ("表头行: %d" % P["header_row"]) in out
              and ("数据行: (%d, %d)" % (P["first_row"], P["last_row"])) in out)
        rec("G2 表型识别(四元组)", ok, P["quad"] if ok else (out[-300:] or r.stderr[-300:]))
    else:
        rec("G2 表型识别(四元组)", False, "缺 identify_table.py: " + IDENTIFY)

    # G3 基线
    V, sheets = snap(path, True)
    F, _ = snap(path, False)
    err = sum(1 for v in V.values() if isinstance(v, str) and v.startswith("#") and not v.startswith("####"))
    fml = sum(1 for v in F.values() if isinstance(v, str) and v.startswith("="))
    med = media_count(path)
    rec("G3 sheet/图片/错误值/真公式 基线",
        len(sheets) == B["sheets"] and med == B["media"] and err == B["errors"] and fml == B["formulas"],
        "sheet%d 图%d 错%d 公式%d（期望 %d/%d/%d/%d）"
        % (len(sheets), med, err, fml, B["sheets"], B["media"], B["errors"], B["formulas"]))
    rows_n = sum(1 for r in ROWS if V.get((MAIN, r, C["K"])) is not None)
    rec("G3 制表数据行数", rows_n == B["rows"], "%d 行（期望 %d）" % (rows_n, B["rows"]))
    AH = sum(float(V[(MAIN, r, C["AH"])]) for r in ROWS
             if isinstance(V.get((MAIN, r, C["AH"])), (int, float)))
    sc = B["summary_col"]
    C27 = V.get((SUMMARY, B["summary_row_c27"], sc))
    C28 = V.get((SUMMARY, B["summary_row_c28"], sc))
    C26 = V.get((SUMMARY, B["summary_row_c26"], sc))
    rec("G3 金额链 AH=C27 / C28 / C26",
        abs(AH - B["ah"]) < 0.01 and isinstance(C27, (int, float)) and abs(float(C27) - AH) < 0.01
        and isinstance(C28, (int, float)) and abs(float(C28) - B["c28"]) < 1e-5
        and isinstance(C26, (int, float)) and abs(float(C26) - B["c26"]) < 1e-5,
        "AH=%.2f C27=%s C28=%s C26=%s" % (AH, C27, C28, C26))

    # G4 13项一致性
    from check_tz_cs_consistency import run as consistency
    rc = consistency(path, profile=P["_profile_file"])
    rec("G4 双向一致性13项", rc == 0, "退出码 %d" % rc)

    # G5 业务视图
    concl = [(r, V.get((CHECK, r, 1))) for r in range(CONCL_FROM, CONCL_TO + 1)
             if V.get((CHECK, r, 1)) not in (None, "")]
    rec("G5 结论区已清空", not concl, concl[:5])
    hits = [(r, c) for r in ROWS for c in (C["K"], C["U"])
            if any(k in str(V.get((MAIN, r, c)) or "") for k in KW)]
    rec("G5 K/U内部取数出处归零", not hits, hits[:5])
    hits_ae = [(r, str(V.get((MAIN, r, C["AE"])) or "")[:60]) for r in ROWS
               if any(k in str(V.get((MAIN, r, C["AE"])) or "") for k in KW)]
    rec("G5 AE内部取数出处归零", not hits_ae, hits_ae[:5])
    U_OK = P.get("u_ok_word", "授权底价一致")
    u_ok = {str(V.get((MAIN, r, C["U"])) or "").strip() for r in ROWS
            if "一致" in str(V.get((MAIN, r, C["U"])) or "")}
    rec("G5 U列一致文案统一", u_ok == {U_OK}, sorted(u_ok)[:5])
    for key, col in (("K", C["K"]), ("U", C["U"]), ("AE", C["AE"])):
        got = sum(1 for r in ROWS if str(V.get((MAIN, r, col)) or "").strip() == W[key])
        want = P["word_counts"][key]
        rec("G5 第%d列「%s」%d行" % (col, W[key], want), got == want, "实际 %d" % got)
    bad_words = []
    for r in ROWS:
        k = str(V.get((MAIN, r, C["K"])) or "").strip()
        u = str(V.get((MAIN, r, C["U"])) or "").strip()
        a = str(V.get((MAIN, r, C["AE"])) or "").strip()
        if "不对" in k and k != W["K"]:
            bad_words.append(("K", r, k))
        if "不对" in u and u != W["U"]:
            bad_words.append(("U", r, u))
        if "不对" in a and a != W["AE"]:
            bad_words.append(("AE", r, a))
    rec("G5 三列零串词", not bad_words, bad_words[:5])
    oa = sum(1 for r in ROWS
             if str(V.get((MAIN, r, C["AE"])) or "")
             .startswith(tuple(P.get("oa_evidence_prefix", ["OA《", "按OA《", "(OA号"]))))
    rec("G5 OA证据行保留(含简化格式)", oa == P["oa_evidence_rows"], "%d 行" % oa)
    blank_ad = [r for r in ROWS
                if V.get((MAIN, r, C["AD"])) is None
                or (isinstance(V.get((MAIN, r, C["AD"])), str)
                    and not V.get((MAIN, r, C["AD"])).strip())]
    zero_ad = [r for r in ROWS
               if isinstance(V.get((MAIN, r, C["AD"])), (int, float))
               and abs(float(V.get((MAIN, r, C["AD"])))) < 0.005]
    rec("G5 AD无空白/0值行数符合配置且全是退货",
        not blank_ad and len(zero_ad) == P.get("ad_zero_expect_rows", 0)
        and all(str(V.get((MAIN, r, C["K"])) or "").strip() == "退货" for r in zero_ad),
        dict(空白=blank_ad[:5], 零值=zero_ad))
    items = [r for r in range(1, P["check_max_row"] + 1)
             if isinstance(V.get((CHECK, r, 1)), (int, float)) and V.get((CHECK, r, 2))]
    gh = [r for r in items if str(V.get((CHECK, r, 7)) or "").strip()
          or str(V.get((CHECK, r, 8)) or "").strip()]
    rec("G5 情况表条目数 / G-H留空", len(items) == P["items"] and not gh,
        "条目%d 已填G/H %d" % (len(items), len(gh)))

    # G5 第五步列归属（用户 11-12 14:44 纠正）：AD 是核对方自行填列，业务不填也不掌握总部价；
    # AE 是 AD 填完后的备注列。只读断言，绝不改文件。
    owner = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "check_ad_ae_ownership.py")
    if os.path.exists(owner):
        cmd = [sys.executable, owner, os.path.abspath(path)]
        if P.get("_profile_file"):
            cmd += ["--profile", P["_profile_file"]]
        q = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        rec("G5 AD/AE第五步列归属8项断言", q.returncode == 0,
            "退出码%d %s" % (q.returncode, (q.stdout or "").strip().splitlines()[-2]
                            if q.returncode == 0 else (q.stdout or q.stderr or "")[-160:]))
    else:
        rec("G5 AD/AE第五步列归属8项断言", False, "缺 check_ad_ae_ownership.py")

    # G6 全簿 diff
    if args.prev:
        A, _ = snap(os.path.abspath(args.prev), True)
        AF, _ = snap(os.path.abspath(args.prev), False)
        diff = {k for k in set(A) | set(V) if A.get(k) != V.get(k)}
        by = collections.Counter(k[0] for k in diff)
        rec("G6 值层差异格数", args.expect_diff is None or len(diff) == args.expect_diff,
            "实际%d 期望%s 分布%s" % (len(diff), args.expect_diff, dict(by)))
        numbad = [(r, c) for r in ROWS for c in NUMCOLS
                  if A.get((MAIN, r, c)) != V.get((MAIN, r, c))]
        rec("G6 数值/公式列零改动", not numbad, numbad[:5])
        fbad = [k for k in set(AF) | set(F)
                if AF.get(k) != F.get(k) and ((isinstance(AF.get(k), str) and AF[k].startswith("="))
                                              or (isinstance(F.get(k), str) and F[k].startswith("=")))]
        rec("G6 真公式文本零改动", not fbad, fbad[:5])
        if args.plan:
            with open(args.plan, encoding="utf-8") as f:
                plan = json.load(f)
            exp = {(x["sheet"], x["r"], x["c"]) for x in plan["plan"]}
            rec("G6 差异格集合==plan", diff == exp,
                dict(多=sorted(diff - exp)[:5], 少=sorted(exp - diff)[:5]))

    bad = [x for x in R if not x[1]]
    print("=" * 82)
    print("门禁结论: %s（%d项，%d项不通过）"
          % ("全部通过 ✅ 允许交付" if not bad else "禁止交付 ❌", len(R), len(bad)))
    if bad:
        print("不通过:", [x[0] for x in bad])
    print("=" * 82)
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", help="待交付 xlsx（--next 时可传目录或任一现有 xlsx）")
    ap.add_argument("--prev", help="上一版 xlsx，用于全簿 diff")
    ap.add_argument("--expect-diff", type=int, help="期望的值层差异格数")
    ap.add_argument("--plan", help="plan.json，用于断言差异格集合")
    ap.add_argument("--next", action="store_true", help="只输出当天下一版应叫的名字")
    ap.add_argument("--profile", help="指定 profile.json（也可用环境变量 KA_PROFILE）")
    a = ap.parse_args()
    if not a.xlsx:
        raise SystemExit("必须给 xlsx 路径")
    if a.next:
        suggest_next(a.xlsx)
        sys.exit(0)
    sys.exit(gate(a))
