# -*- coding: utf-8 -*-
"""型号归一化唯一公共模块（把三轮「假未匹配」固化成可跑断言）。

血泪：只用 upper + 去空格的精确匹配会漏
  85Z11L-A <-> 85Z11L、65Z9H-JN <-> 65Z9H、75A300-Pro <-> 75A300 Pro
必须用 norm() 统一键 + candidates() 分级去有限后缀；精确失败先跑变体再下「没有」结论。
等级机行（特一/特二/特三）必须判【原始串】，归一化 [^A-Z0-9] 会吃掉中文括号。

自测: python model_key.py        （PASS 才可参与匹配）
"""
import re
import sys

# 长后缀优先，避免 "A" 抢先截断 "...PRO"
SUFFIXES = ("PRO", "JN", "SE", "A", "C")
GRADE_MARKS = ("特一", "特二", "特三", "等级机")


def norm(s):
    """统一键：去所有非字母数字 + 大写。连字符/空格/中文括号一律吃掉。"""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def candidates(model):
    """匹配候选键，按优先级：原键 → 逐级去有限后缀。精确失败必须逐个再查。"""
    k = norm(model)
    out = [k]
    base = k
    changed = True
    while changed:
        changed = False
        for suf in SUFFIXES:
            if base.endswith(suf) and len(base) > len(suf) + 1:
                base = base[:-len(suf)]
                out.append(base)
                changed = True
                break
    return list(dict.fromkeys(x for x in out if x))


def is_grade(text):
    """等级机行判原始串；只有等级机价时不得硬凑标准价填 AD。"""
    s = str(text or "")
    return any(m in s for m in GRADE_MARKS)


def split_rows(rows):
    """rows: [(原始型号, ...), ...] → (标准行, 等级机行)"""
    std, grade = [], []
    for row in rows:
        (grade if is_grade(row[0]) else std).append(row)
    return std, grade


def lookup(table, model, default=None):
    """table: {norm键: 值}；返回 (命中的键, 值)，全未命中返回 (None, default)。"""
    for k in candidates(model):
        if k in table:
            return k, table[k]
    return None, default


def selftest():
    assert norm("85Z11L-A") == "85Z11LA"
    assert norm("75A300-Pro") == norm("75A300 Pro") == "75A300PRO"
    assert "85Z11L" in candidates("85Z11L-A"), "-A 后缀坑"
    assert "65Z9H" in candidates("65Z9H-JN"), "-JN 后缀坑"
    assert "75A300" in candidates("75A300-Pro"), "连字符坑"
    assert candidates("85Z11L-A")[0] == "85Z11LA", "精确键必须排第一"
    assert is_grade("85Z11K(特一)") and not is_grade("85Z11K")
    std, grade = split_rows([("85Z11K", 9000), ("85Z11K(特一)", 4790)])
    assert len(std) == 1 and len(grade) == 1
    k, v = lookup({"85Z11L": 10999}, "85Z11L-A")
    assert k == "85Z11L" and v == 10999
    k, v = lookup({}, "不存在的型号")
    assert k is None and v is None
    return True


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    print("model_key 自测:", "PASS" if selftest() else "FAIL")
    for s in ("85Z11L-A", "65Z9H-JN", "75A300-Pro", "75A300 Pro"):
        print("  %-14s -> %s" % (s, candidates(s)))
