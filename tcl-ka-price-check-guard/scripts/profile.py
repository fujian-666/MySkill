# -*- coding: utf-8 -*-
"""四元组配置的唯一入口。

铁律：战区 × 渠道 × 月份 × 版本 任一不同 = 另一张表，
不套列位 / 口径 / 源表 / 基线。换表 = 换 profile，绝不改脚本逻辑。

优先级：
  1. load(path) 显式传入
  2. 环境变量 KA_PROFILE
  3. config/profile.json（本地真实配置，gitignore）
  4. config/profile.example.json（模板，基线为 0，跑门禁会失败 —— 有意的防呆）

用法：
    from profile import load
    P = load()
    rows = range(P["first_row"], P["last_row"] + 1)
"""
import os
import json
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_FILE = os.path.join(ROOT, "config", "profile.example.json")
LOCAL_FILE = os.path.join(ROOT, "config", "profile.json")


def _expand(v):
    if isinstance(v, str):
        return (v.replace("${SKILL_DIR}", ROOT)
                 .replace("${HOME}", os.path.expanduser("~")))
    if isinstance(v, list):
        return [_expand(x) for x in v]
    if isinstance(v, dict):
        return {k: _expand(x) for k, x in v.items()}
    return v


def _deep_update(base, override):
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_update(base[k], v)
        else:
            base[k] = v
    return base


def load(path=None):
    """返回配置 dict。模板与本地配置深合并，缺键自动回落到模板默认值。"""
    chosen = path or os.environ.get("KA_PROFILE") or (
        LOCAL_FILE if os.path.exists(LOCAL_FILE) else DEFAULT_FILE)
    with open(DEFAULT_FILE, encoding="utf-8") as f:
        cfg = json.load(f)
    if os.path.abspath(chosen) != os.path.abspath(DEFAULT_FILE):
        with open(chosen, encoding="utf-8") as f:
            _deep_update(cfg, json.load(f))
    cfg["_profile_file"] = chosen
    return _expand(cfg)


def rows(P):
    return range(P["first_row"], P["last_row"] + 1)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    P = load()
    print("profile:", P["_profile_file"])
    print("quad:", P["quad"], " owner:", P["owner"])
    print("rows:", P["first_row"], "-", P["last_row"], " sheets:", P["sheets"])
    print("words:", P["words"])
