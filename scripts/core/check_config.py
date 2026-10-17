# -*- coding: utf-8 -*-
"""
配置体检 —— 判断这套系统是「未配置 / 有问题 / 可以直接跑」。

用途：
  1. 初次使用时，AI 先用它确认现状，再决定要不要走引导流程
  2. AI 生成/改完 config 后，跑一次自检，避免带着错误配置去跑引擎

用法：
    python core/check_config.py [icost-system目录]
    ./check.sh                    # 包装脚本，自动找 Python

退出码：
    0  就绪，可以直接跑 run.sh
    1  已配置但有问题（会列出具体问题）
    2  未配置（模板状态），需要先按 初次引导.md 配置
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SYSTEM = os.path.dirname(HERE)

CFG_FILES = ['categories.json', 'accounts.json', 'pay_mapping.json',
             'rules.json', 'merchant_map.json', 'import_fix.json']
NEED_SETUP_MARK = 'SETUP_PENDING.md'
SETUP_TPL = 'icost_template.xlsx'


def jload(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


class Report(object):
    def __init__(self):
        self.err = []      # 阻断：跑不了
        self.warn = []     # 警告：能跑但结果会不对
        self.info = []     # 统计信息

    def e(self, where, msg):
        self.err.append('%s: %s' % (where, msg))

    def w(self, where, msg):
        self.warn.append('%s: %s' % (where, msg))

    def i(self, msg):
        self.info.append(msg)


def check_env(rep):
    try:
        import openpyxl                      # noqa: F401
        rep.i('Python %s · openpyxl 已安装' % sys.version.split()[0])
    except ImportError:
        rep.e('环境', '缺少 openpyxl，请先 pip install openpyxl')


def check_files(cfg, rep):
    missing = [f for f in CFG_FILES if not os.path.exists(os.path.join(cfg, f))]
    if missing:
        rep.e('config/', '缺少文件：' + '、'.join(missing))
    if not os.path.exists(os.path.join(cfg, SETUP_TPL)):
        rep.e('config/', '缺少官方模板 %s（它是导入文件的骨架，不能删）' % SETUP_TPL)
    return set(CFG_FILES) - set(missing)


def check_categories(cfg, rep):
    p = os.path.join(cfg, 'categories.json')
    if not os.path.exists(p):
        return None
    try:
        t = jload(p)
    except Exception as ex:
        rep.e('categories.json', 'JSON 解析失败：%s' % ex)
        return None
    if not isinstance(t, dict) or not t:
        rep.e('categories.json', '内容为空，格式应为 {账本: {方向: {一级: [二级]}}}')
        return None
    n1 = n2 = 0
    for bk, dirs in t.items():
        if not isinstance(dirs, dict):
            rep.e('categories.json', '账本「%s」下应为 {方向: {...}}，实际是 %s' % (bk, type(dirs).__name__))
            continue
        for sd in ('支出', '收入'):
            if sd not in dirs:
                rep.w('categories.json', '账本「%s」没有「%s」方向，该方向无法归类' % (bk, sd))
        for sd, c1s in dirs.items():
            if sd not in ('支出', '收入'):
                rep.w('categories.json', '账本「%s」出现非标准方向「%s」（只能是 支出/收入）' % (bk, sd))
            if not isinstance(c1s, dict):
                rep.e('categories.json', '%s/%s 下应为 {一级: [二级]}' % (bk, sd))
                continue
            for c1, c2s in c1s.items():
                n1 += 1
                if c2s is None:
                    continue
                if not isinstance(c2s, list):
                    rep.e('categories.json', '%s/%s/%s 的二级应为数组' % (bk, sd, c1))
                    continue
                n2 += len(c2s)
    rep.i('分类树：%d 个账本 · %d 个一级 · %d 个二级' % (len(t), n1, n2))
    return t


def check_accounts(cfg, rep):
    p = os.path.join(cfg, 'accounts.json')
    if not os.path.exists(p):
        return set()
    try:
        a = jload(p)
    except Exception as ex:
        rep.e('accounts.json', 'JSON 解析失败：%s' % ex)
        return set()
    names = {}
    total = 0
    for grp, val in a.items():
        # ⚠️ 顶层不能有 _说明 之类的键：引擎会遍历 value 调 .get('accounts')
        if grp.startswith('_'):
            rep.e('accounts.json', '顶层不能出现「%s」这类下划线键（引擎会崩溃），说明请写到 config/README.md' % grp)
            continue
        if not isinstance(val, dict) or 'accounts' not in val:
            rep.e('accounts.json', '组「%s」格式应为 {"accounts": [...]}，实际缺 accounts 字段' % grp)
            continue
        for acc in val['accounts']:
            nm = acc.get('name') if isinstance(acc, dict) else acc
            if not nm:
                rep.w('accounts.json', '组「%s」里有缺 name 的账户' % grp)
                continue
            total += 1
            names.setdefault(nm, []).append(grp)
            if isinstance(acc, dict) and acc.get('bal'):
                pass
    dup = {k: v for k, v in names.items() if len(v) > 1}
    if dup:
        rep.w('accounts.json', '账户名重复（导入会对不上）：' +
              '、'.join('%s（在 %s）' % (k, '/'.join(v)) for k, v in list(dup.items())[:5]))
    rep.i('账户：%d 个，分 %d 组' % (total, len(a)))
    return set(names.keys())


def _path_ok(tree, bk, sd, c1, c2):
    if tree is None:
        return True
    if bk not in tree:
        return False
    if sd not in tree[bk]:
        return False
    if c1 not in tree[bk][sd]:
        return False
    if c2:
        subs = tree[bk][sd][c1] or []
        if c2 not in subs:
            return False
    return True


def check_rule_paths(cfg, rep, tree):
    """rules / merchant_map 里的分类路径必须存在于分类树"""
    p = os.path.join(cfg, 'rules.json')
    if os.path.exists(p):
        try:
            r = jload(p)
        except Exception as ex:
            rep.e('rules.json', 'JSON 解析失败：%s' % ex)
            r = {}
        bad = []
        for section in ('规则', '品类词', '平台兜底'):
            items = r.get(section)
            if not isinstance(items, list):
                continue
            for it in items:
                if not isinstance(it, dict):
                    continue
                if not _path_ok(tree, it.get('账本'), it.get('方向'),
                                it.get('一级'), it.get('二级')):
                    bad.append('%s/%s/%s/%s（%s）' % (it.get('账本'), it.get('方向'),
                                                  it.get('一级'), it.get('二级'), section))
        if bad:
            rep.w('rules.json', '%d 条规则的分类不在分类树里，例：%s' % (len(bad), '；'.join(bad[:3])))
        else:
            cnt = sum(len(r.get(k, []) or []) for k in ('规则', '品类词', '平台兜底'))
            rep.i('归类规则：%d 条（分类路径全部有效）' % cnt)

    p = os.path.join(cfg, 'merchant_map.json')
    if os.path.exists(p):
        try:
            m = jload(p)
        except Exception as ex:
            rep.e('merchant_map.json', 'JSON 解析失败：%s' % ex)
            return
        bad = []
        real = 0
        for k, v in m.items():
            if k.startswith('_'):
                continue
            real += 1
            if isinstance(v, list) and len(v) >= 3:
                bk, sd, c1 = v[0], v[1], v[2]
                c2 = v[3] if len(v) > 3 else ''
                if not _path_ok(tree, bk, sd, c1, c2):
                    bad.append('%s → %s/%s/%s' % (k, bk, c1, c2))
        if bad:
            rep.w('merchant_map.json', '%d 条映射的分类不在分类树里，例：%s' % (len(bad), '；'.join(bad[:3])))
        rep.i('商家映射：%d 条%s' % (real, '' if real else '（空 —— 首次使用可以为空，引擎会从历史里学）'))


def check_pay_mapping(cfg, rep, acc_names):
    p = os.path.join(cfg, 'pay_mapping.json')
    if not os.path.exists(p):
        return
    try:
        pm = jload(p)
    except Exception as ex:
        rep.e('pay_mapping.json', 'JSON 解析失败：%s' % ex)
        return
    rules = pm.get('规则')
    if not isinstance(rules, list) or not rules:
        rep.w('pay_mapping.json', '没有规则，扣款账户识别不到')
        return
    if not acc_names:
        return
    bad = [r.get('账户') for r in rules
           if isinstance(r, dict) and r.get('账户') and r['账户'] not in acc_names]
    if bad:
        rep.w('pay_mapping.json', '%d 条规则指向不存在的账户：%s' % (len(set(bad)), '、'.join(list(set(bad))[:3])))
    else:
        rep.i('支付映射：%d 条规则，账户全部有效' % len(rules))


def check_import_fix(cfg, rep):
    p = os.path.join(cfg, 'import_fix.json')
    if not os.path.exists(p):
        return
    try:
        f = jload(p)
    except Exception as ex:
        rep.e('import_fix.json', 'JSON 解析失败：%s' % ex)
        return
    bad = [it.get('key') for it in f.get('patch', [])
           if isinstance(it, dict) and '|' not in str(it.get('key', ''))]
    if bad:
        rep.w('import_fix.json', 'patch 的 key 必须形如「YYYY/MM/DD HH:MM:SS|金额」，有 %d 条不符' % len(bad))
    n = len(f.get('patch', [])) + len(f.get('drop', []))
    rep.i('人工修正表：%d 条' % n)


def check_template(cfg, rep):
    p = os.path.join(cfg, SETUP_TPL)
    if not os.path.exists(p):
        return
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p)
        ws = wb.active
        hdr = [c.value for c in ws[1]]
        if ws.title != 'icost_template':
            rep.w(SETUP_TPL, 'sheet 名是「%s」，官方模板应为 icost_template' % ws.title)
        if len([h for h in hdr if h]) != 10:
            rep.w(SETUP_TPL, '表头应有 10 列，实际 %d 列' % len([h for h in hdr if h]))
        else:
            rep.i('官方模板：%s · 10 列 · 正常' % ws.title)
    except ImportError:
        pass
    except Exception as ex:
        rep.e(SETUP_TPL, '读取失败：%s' % ex)


def main():
    system = sys.argv[1] if len(sys.argv) > 1 else SYSTEM
    cfg = os.path.join(system, 'config')
    rep = Report()

    print('=' * 58)
    print('  iCost 记账系统 · 配置体检')
    print('=' * 58)
    print('系统目录：%s\n' % system)

    if not os.path.isdir(cfg):
        print('❌ 找不到 config/ 目录，检查路径是否正确')
        return 2

    # 0) 未配置标记（脱敏包里会有这个文件）
    pending = os.path.join(cfg, NEED_SETUP_MARK)
    if os.path.exists(pending):
        print('🟡 状态：未配置（模板状态）\n')
        for line in rep_info_pending():
            print('  ' + line)
        print('\n请按 icost-system/初次引导.md 完成首次配置。')
        return 2

    check_env(rep)
    have = check_files(cfg, rep)
    tree = check_categories(cfg, rep) if 'categories.json' in have else None
    acc_names = check_accounts(cfg, rep) if 'accounts.json' in have else set()
    check_rule_paths(cfg, rep, tree)
    check_pay_mapping(cfg, rep, acc_names)
    check_import_fix(cfg, rep)
    check_template(cfg, rep)

    print('--- 概览 ---')
    for x in rep.info:
        print('  · ' + x)
    if rep.warn:
        print('\n--- ⚠️ 警告（能跑，但结果可能不对）---')
        for x in rep.warn:
            print('  · ' + x)
    if rep.err:
        print('\n--- ❌ 错误（必须修，否则跑不了）---')
        for x in rep.err:
            print('  · ' + x)

    print('')
    if rep.err:
        print('🔴 状态：有问题 —— 先修上面的「错误」')
        return 1
    if rep.warn:
        print('🟡 状态：可以跑，但有 %d 条警告，建议先处理' % len(rep.warn))
        return 0
    print('🟢 状态：就绪 —— 可以直接跑 ./run.sh')
    return 0


def rep_info_pending():
    return [
        '这个包里的 config/ 是脱敏模板，还没有填入你的真实配置。',
        '',
        '需要填的三个文件（按优先级）：',
        '  1. config/categories.json   从 iCost「分类管理」截图录入',
        '  2. config/accounts.json     从 iCost「账户管理」截图录入',
        '  3. config/pay_mapping.json  支付渠道 → 你的账户',
        '',
        '填完后删除 config/SETUP_PENDING.md，再跑一次本脚本。',
    ]


if __name__ == '__main__':
    sys.exit(main())
