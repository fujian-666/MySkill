# -*- coding: utf-8 -*-
"""制表 <-> 核对情况表 双向一致性校验（每版交付前必跑，13 项）。

规则：
  1 制表不一致正向行 == 情况表行集（双向，按行集合，不靠文案分类映射）
  2 制表 U 不一致正向行均在情况表（只列正向，差集只能是退货行）
  3 情况表三类条目在制表侧有对应标注
  4 情况表每一条，制表对应列必须确实有「不一致」标注（不许情况表列了、制表写一致）
  5 情况表异常说明里的说明价，必须与制表 H / R / AD 对应单元格数值一致
  6 标「一致」的行，J/T 差额必须真的为 0（不许假一致）
  7 标「不对」的行，J/T 差额必须非零；I=0/S=0 的零基准保护行豁免
  8 大区供价 I=0 / 大区零售价 S=0 的行，不得标「一致」
  9 情况表 G/H 已填的条目，制表对应备注必须已回填写入（未填则跳过）
 10 三个价格各是各的描述：K/U/AE 只能用各自的专属词，且总部条目在 AE 必须有标注
 11 U 列「一致」文案统一为 profile.u_ok_word（不许再写来源长句）
 12 AE 列内部取数出处归零，只允许短标注白名单；OA 证据行仍 oa_evidence_rows 行
 13 总部最低零售限价 AD 无空白；0 值行数 == ad_zero_expect_rows 且全是退货行

用法: python check_tz_cs_consistency.py "<xlsx路径>" [--json] [--profile <profile.json>]
退出码 0=全过, 1=有不一致
"""
import sys, json, re
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook
from profile import load as load_profile, rows as profile_rows


def num(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def run(path, as_json=False, profile=None):
    P = load_profile(profile)
    C, W = P["cols"], P["words"]
    ROWS = profile_rows(P)
    MAIN, CHECK = P["sheets"]["main"], P["sheets"]["check"]
    TY_GJ, TY_ZB, TY_HQ = P["types"]
    WHITE = P["white"]
    MARK = P["backfill_mark"]

    def clean(v):
        """剥掉回填段，避免回填后误判为长文案"""
        s = str(v or "").strip()
        return s.split("｜" + MARK)[0].strip()

    def plain(v):
        s = clean(v)
        return any(w in s for w in WHITE)

    wb = load_workbook(path, read_only=True, data_only=True)
    TZ = list(wb[MAIN].iter_rows(min_row=1, max_row=P["read_max_row"],
                                 max_col=P["read_max_col"], values_only=True))
    CS = list(wb[CHECK].iter_rows(min_row=1, max_row=P["check_max_row"],
                                  max_col=P["check_max_col"], values_only=True))
    def t(r, c):
        return TZ[r - 1][c - 1]

    items = []
    for i, row in enumerate(CS, 1):
        if isinstance(row[0], (int, float)) and row[1]:
            items.append(dict(r=i, no=int(row[0]), typ=str(row[1]), row=int(row[2]),
                              desc=row[5], g=row[6], h=row[7]))

    K_bad = {r for r in ROWS if not plain(t(r, C["K"]))}
    U_bad = {r for r in ROWS if not plain(t(r, C["U"]))}
    K_gj = {r for r in ROWS if clean(t(r, C["K"])) == W["K"]}
    U_dj = {r for r in ROWS if clean(t(r, C["U"])) in (W["U"], "底价不对")}
    HQ_bad = {r for r in ROWS if clean(t(r, C["AE"])) == W["AE"]}
    RET = {r for r in ROWS if clean(t(r, C["K"])) == "退货"}
    cs_gj = {x["row"] for x in items if x["typ"] == TY_GJ}
    cs_zb = {x["row"] for x in items if x["typ"] == TY_ZB}
    cs_hq = {x["row"] for x in items if x["typ"] == TY_HQ}
    cs_all = {x["row"] for x in items}

    res = []
    def chk(name, ok, detail):
        res.append(dict(name=name, ok=bool(ok), detail=detail))

    # 文案简化后不能靠文案分类映射，一律按「行集合」双向一致
    all_bad_pos = {r for r in ROWS if clean(t(r, C["K"])) != "退货"
                   and (r in K_bad or r in U_bad)}
    chk("1 制表不一致正向行 == 情况表行集(双向)", all_bad_pos == cs_all,
        dict(制表多=sorted(all_bad_pos - cs_all), 情况表多=sorted(cs_all - all_bad_pos)))
    extra = U_bad - cs_all
    chk("2 U列不一致正向行均在情况表(只列正向)", extra <= RET,
        dict(非退货例外=sorted(extra - RET), 退货豁免=sorted(extra & RET)))
    t1, t2, t3 = cs_gj - K_bad, cs_zb - U_bad, cs_hq - (K_bad | U_bad | HQ_bad)
    chk("3 情况表三类条目在制表侧有对应标注", not t1 and not t2 and not t3,
        dict(供价核对无K标注=sorted(t1), 底价核对无U标注=sorted(t2), 总部底价无标注=sorted(t3)))
    no_mark = []
    for x in items:
        r = x["row"]
        if x["typ"] == TY_GJ and r not in K_bad:
            no_mark.append((x["no"], x["typ"], r))
        if x["typ"] == TY_ZB and r not in U_bad:
            no_mark.append((x["no"], x["typ"], r))
        if x["typ"] == TY_HQ and r not in (K_bad | U_bad):
            no_mark.append((x["no"], x["typ"], r))
    chk("4 情况表每条制表侧均有不一致标注", not no_mark, dict(无标注=no_mark[:30]))

    bad = []
    for x in items:
        r = x["row"]
        v = t(r, P["type2price_col"][x["typ"]])
        nums = re.findall(r"\d+(?:\.\d+)?", str(x["desc"] or ""))
        if not (isinstance(v, (int, float)) and any(abs(float(n) - float(v)) < 0.01 for n in nums)):
            bad.append((x["no"], x["typ"], r, v, str(x["desc"])[:60]))
    chk("5 说明价 == 制表 H/R/AD", not bad, dict(未对上=bad[:30]))

    # 6~8：「两边一致」不等于「标一致的行真的一致」
    k_fake = [r for r in ROWS if clean(t(r, C["K"])) == "供价一致" and abs(num(t(r, C["J"]))) > 0.01]
    u_fake = [r for r in ROWS if "一致" in clean(t(r, C["U"])) and abs(num(t(r, C["T"]))) > 0.01]
    chk("6 标供价一致/底价一致的行差额确为0", not k_fake and not u_fake,
        dict(供价假一致=k_fake[:20], 底价假一致=u_fake[:20]))
    # I=0 / S=0 时 J/T 由零基准保护恒为 0，不是假差异；第8项保证它们不标「一致」
    k_zero = [r for r in ROWS if clean(t(r, C["K"])) == W["K"]
              and abs(num(t(r, C["J"]))) < 0.01 and num(t(r, C["I"])) != 0]
    u_zero = [r for r in ROWS if clean(t(r, C["U"])) in (W["U"], "底价不对")
              and abs(num(t(r, C["T"]))) < 0.01 and num(t(r, C["S"])) != 0]
    zero_base = [r for r in ROWS if num(t(r, C["I"])) == 0 and num(t(r, C["S"])) == 0]
    chk("7 标不对的行差额非零(零基准保护除外)", not k_zero and not u_zero,
        dict(供价零差=k_zero[:20], 底价零差=u_zero[:20], 零基准豁免=len(zero_base)))
    fake_zero = [r for r in ROWS if clean(t(r, C["K"])) != "退货" and
                 ((num(t(r, C["I"])) == 0 and "一致" in clean(t(r, C["K"]))) or
                  (num(t(r, C["S"])) == 0 and "一致" in clean(t(r, C["U"]))))]
    chk("8 大区供价/零售价为0的行不标一致", not fake_zero, dict(假一致=fake_zero[:20]))

    filled = [x for x in items if str(x["g"] or "").strip() or str(x["h"] or "").strip()]
    not_back = []
    for x in filled:
        tgt = P["type2remark_col"][x["typ"]]
        if MARK not in str(t(x["row"], tgt) or ""):
            not_back.append((x["no"], x["typ"], x["row"], tgt))
    chk("9 已填G/H 均已回填制表", not not_back,
        dict(已填条数=len(filled), 未回填=not_back[:30]))

    # 10 三个价格是三个定义，备注各是各的描述，不许串词
    k_txt = {clean(t(r, C["K"])) for r in ROWS if not plain(t(r, C["K"]))}
    u_txt = {clean(t(r, C["U"])) for r in ROWS if not plain(t(r, C["U"]))}
    ae_txt = {clean(t(r, C["AE"])) for r in ROWS if "不对" in clean(t(r, C["AE"]))}
    chk("10 三列描述各是各的(K/U/AE)", k_txt == {W["K"]} and u_txt == {W["U"]}
        and ae_txt <= {W["AE"]} and HQ_bad == cs_hq,
        dict(K非标准=sorted(k_txt - {W["K"]}), U非标准=sorted(u_txt - {W["U"]}),
             AE串词=sorted(ae_txt - {W["AE"]}),
             总部条目缺标注=sorted(cs_hq - HQ_bad), 总部多标=sorted(HQ_bad - cs_hq)))

    # 11~13 用户纠正即固化断言：U 一致短词 / AE 短标注白名单 / AD 零值闭环
    U_OK = P.get("u_ok_word", "授权底价一致")
    u_ok_txt = {clean(t(r, C["U"])) for r in ROWS if "一致" in clean(t(r, C["U"]))}
    chk("11 U列一致文案统一", u_ok_txt == {U_OK}, dict(实际=sorted(u_ok_txt)))
    AE_KW = P["internal_keywords"]
    AE_SHORT = set(P.get("ae_short_labels", [])) | {W["AE"]}
    OA_PREFIX = P.get("oa_evidence_prefix", ["OA《", "按OA《", "(OA号"])
    ae_all = {clean(t(r, C["AE"])) for r in ROWS if clean(t(r, C["AE"]))}
    ae_long = {v for v in ae_all if v.startswith(("OA《", "按OA《"))}
    bad_ae = sorted(v for v in ae_all if any(k in v for k in AE_KW))
    oa_cnt = sum(1 for r in ROWS if clean(t(r, C["AE"])).startswith(tuple(OA_PREFIX)))
    chk("12 AE列出处归零/短标注白名单/OA证据行数",
        not bad_ae and (ae_all - ae_long) <= AE_SHORT and oa_cnt == P["oa_evidence_rows"],
        dict(残留出处=bad_ae[:5], 非白名单=sorted((ae_all - ae_long) - AE_SHORT)[:5],
             OA证据=oa_cnt, 期望=P["oa_evidence_rows"]))
    blank_ad = [r for r in ROWS if t(r, C["AD"]) is None
                or (isinstance(t(r, C["AD"]), str) and not t(r, C["AD"]).strip())]
    zero_ad = [r for r in ROWS if isinstance(t(r, C["AD"]), (int, float))
               and abs(float(t(r, C["AD"]))) < 0.005]
    chk("13 AD无空白/0值行数符合配置且全是退货",
        not blank_ad and len(zero_ad) == P.get("ad_zero_expect_rows", 0)
        and all(r in RET for r in zero_ad),
        dict(空白=blank_ad[:20], 零值行=sorted(zero_ad), 期望=P.get("ad_zero_expect_rows", 0)))

    wb.close()
    if as_json:
        print(json.dumps(dict(file=path, pass_=all(x["ok"] for x in res), checks=res),
                         ensure_ascii=False, indent=1))
    else:
        print("文件:", path)
        for x in res:
            print("  [%s] %s %s" % ("OK" if x["ok"] else "FAIL", x["name"],
                                    "" if x["ok"] else json.dumps(x["detail"], ensure_ascii=False)))
        print("结论:", "全部通过" if all(x["ok"] for x in res) else "存在不一致，禁止交付")
    return 0 if all(x["ok"] for x in res) else 1


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--profile")
    a = ap.parse_args()
    sys.exit(run(a.xlsx, a.json, a.profile))
