#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
icost 智能记账引擎
==================

从微信/支付宝账单文件生成 iCost 导入模板，并处理跨账单去重。

设计原则：
1. 配置与代码分离 —— 分类树、账户、支付映射全在 config/ 里，换人用改配置即可
2. 三层归类策略 —— 语义规则 > 历史学习 > 平台兜底
3. 复用历史决策 —— 从已有账单里提取判别词，纯度 ≥80% 才采纳
4. 永远去重   —— 以「日期+金额」为键，避免重复导入

用法：
    python3 icost.py --wechat 微信账单.xlsx --alipay 支付宝账单.csv \
                     --history iCost历史导出.xlsx --out ./output
"""

import os
import re
import csv
import json
import sys
import glob
import argparse
import datetime
import collections

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG = os.path.join(BASE, 'config')


# ══════════════════════════════════════════════════════════
# 配置加载
# ══════════════════════════════════════════════════════════

def load_config():
    def rd(n):
        with open(os.path.join(CONFIG, n), encoding='utf-8') as f:
            return json.load(f)
    return {
        'cats': rd('categories.json'),
        'accts': rd('accounts.json'),
        'pay': rd('pay_mapping.json'),
        'rules': rd('rules.json'),
        'merch': rd('merchant_map.json'),
    }


def account_names(cfg):
    """全部主力账户名，按历史笔数降序（用于下拉框）"""
    out = []
    for grp, v in cfg['accts'].items():
        if grp.startswith('历史'):      # 归档组不作为可选项
            continue
        for a in v['accounts']:
            if a.get('main'):
                out.append({'name': a['name'], 'grp': grp, 'hist': a.get('hist', 0)})
    out.sort(key=lambda x: -x['hist'])
    return out


# ══════════════════════════════════════════════════════════
# 账单解析
# ══════════════════════════════════════════════════════════

def _read_any(path):
    """按后缀选择解析器"""
    if path.lower().endswith(('.xlsx', '.xlsm')):
        return _parse_xlsx(path)
    if path.lower().endswith('.csv'):
        return _parse_csv(path)
    raise ValueError('不支持的格式: %s' % path)


def _parse_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    # 定位表头
    hi = None
    for i, r in enumerate(rows[:60]):
        if r and r[0] and str(r[0]).strip() in ('交易时间', '日期'):
            hi = i
            break
    if hi is None:
        raise ValueError('%s 未找到表头' % path)
    recs = []
    for r in rows[hi + 1:]:
        if not r or not r[0]:
            continue
        recs.append({
            'dt': str(r[0]).strip(),
            'io': str(r[4] or '').strip() if len(r) > 4 else '',
            'amt': r[5] if len(r) > 5 else 0,
            'party': str(r[2] or '').strip() if len(r) > 2 else '',
            'goods': str(r[3] or '').strip() if len(r) > 3 else '',
            'pay': str(r[6] or '').strip() if len(r) > 6 else '',
            'typ': str(r[1] or '').strip() if len(r) > 1 else '',
            'st': str(r[7] or '').strip() if len(r) > 7 else '',
        })
    return recs


def _parse_csv(path):
    raw = None
    for enc in ('gbk', 'gb18030', 'utf-8-sig', 'utf-8'):
        try:
            with open(path, encoding=enc) as f:
                raw = f.read()
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise ValueError('%s 编码无法识别' % path)
    lines = raw.splitlines()
    hi = next((i for i, l in enumerate(lines) if '交易时间' in l), None)
    if hi is None:
        raise ValueError('%s 未找到表头' % path)
    rows = list(csv.reader(lines[hi:]))
    recs = []
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        recs.append({
            'dt': r[0].strip(), 'io': r[5].strip() if len(r) > 5 else '',
            'amt': r[6].strip() if len(r) > 6 else '0',
            'party': r[2].strip() if len(r) > 2 else '',
            'goods': r[4].strip() if len(r) > 4 else '',
            'pay': r[7].strip() if len(r) > 7 else '',
            'typ': r[1].strip() if len(r) > 1 else '',
            'st': r[8].strip() if len(r) > 8 else '',
        })
    return recs


# ══════════════════════════════════════════════════════════
# 支付方式 → 账户
# ══════════════════════════════════════════════════════════

def pay_to_account(pay, cfg):
    p = (pay or '').strip()
    if not p or p == '/':
        return ''
    # 组合支付拆分
    parts = [x.strip() for x in p.split('&') if x.strip()]
    rules = sorted(cfg['pay']['规则'], key=lambda r: r['优先级'])
    for part in parts:
        for rule in rules:
            if any(k in part for k in rule['匹配']):
                return rule['账户']
    # 整串再试一次（不拆）
    for rule in rules:
        if any(k in p for k in rule['匹配']):
            return rule['账户']
    return ''


# ══════════════════════════════════════════════════════════
# 历史去重
# ══════════════════════════════════════════════════════════

def _pdt(v):
    if isinstance(v, datetime.datetime):
        return v
    s = str(v).strip()
    for f in ('%Y/%m/%d %H:%M:%S', '%Y/%m/%d %H:%M', '%Y/%m/%d',
              '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return datetime.datetime.strptime(s, f)
        except ValueError:
            continue
    return None


def dedupe(recs, history_path):
    """以（日期, 金额）为键剔除已在 iCost 中的记录"""
    if not history_path or not os.path.exists(history_path):
        return recs, []
    import openpyxl
    wb = openpyxl.load_workbook(history_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(x).strip() if x else '' for x in rows[0]]
    try:
        i_dt, i_amt, i_note = hdr.index('日期'), hdr.index('金额'), hdr.index('备注')
    except ValueError:
        return recs, []
    idx = collections.defaultdict(list)
    for r in rows[1:]:
        t = _pdt(r[i_dt])
        if not t:
            continue
        amt = round(abs(r[i_amt]) if isinstance(r[i_amt], (int, float)) else 0, 2)
        idx[(t.strftime('%Y-%m-%d'), amt)].append(
            (t, r[i_amt], r[i_note] if i_note < len(r) else ''))
    keep, dup = [], []
    for r in recs:
        t = _pdt(r['dt'])
        if not t:
            keep.append(r)
            continue
        try:
            amt = round(abs(float(r['amt'])), 2)
        except (TypeError, ValueError):
            keep.append(r)
            continue
        k = (t.strftime('%Y-%m-%d'), amt)
        if k in idx:
            dup.append((t, amt, r, idx[k][0]))
        else:
            keep.append(r)
    return keep, dup


# ══════════════════════════════════════════════════════════
# 历史学习：提取判别词
# ══════════════════════════════════════════════════════════

def learn_keywords(history_path, cfg):
    """从历史账单提取判别词。

    策略：
    1. 优先取完整词组（备注中连续 2-6 字的语义片段），避免滑动窗口产生碎片
    2. 要求出现次数 ≥ min_n（默认 3），纯度 ≥ purity（默认 0.85）
    3. 去掉被更长词条包含的冗余短词
    """
    if not history_path or not os.path.exists(history_path):
        return []
    import openpyxl
    wb = openpyxl.load_workbook(history_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(x).strip() if x else '' for x in rows[0]]
    try:
        i_amt, i_c1, i_c2, i_note, i_bk, i_type = (
            hdr.index('金额'), hdr.index('一级分类'), hdr.index('二级分类'),
            hdr.index('备注'), hdr.index('账本'), hdr.index('类型'))
    except ValueError:
        return []
    valid = set()
    for bk, sides in cfg['cats'].items():
        for sd, c1s in sides.items():
            for c1, c2s in c1s.items():
                valid.add((bk, sd, c1, ''))
                for c2 in c2s:
                    valid.add((bk, sd, c1, c2))
    EXP = {'支出', '还款', '待报销'}
    kw = collections.defaultdict(collections.Counter)
    for r in rows[1:]:
        if not r or r[i_type] not in EXP or not r[i_note]:
            continue
        key = (r[i_bk], '支出', r[i_c1], r[i_c2] or '')
        if key not in valid:
            continue
        note = str(r[i_note]).strip()
        # 切出 2-6 字的候选片段（含标点边界）
        for seg in re.split(r'[\s\-_/｜|、,，。()（）【】\[\]「」]', note):
            seg = seg.strip()
            if len(seg) < 2:
                continue
            if re.search(r'[a-zA-Z0-9]', seg):
                # 含英文数字的整段也作为一个候选（如 AppleID、TokenPlanPlus）
                kw[seg][:0]
            L = len(seg)
            for n in range(2, min(L, 6) + 1):
                for i in range(L - n + 1):
                    g = seg[i:i + n]
                    if re.search(r'[a-zA-Z0-9]', g):
                        continue
                    kw[g][key] += 1
            if L <= 6 and not re.search(r'[a-zA-Z0-9]', seg):
                kw[seg][key] += 1
    cands = []
    for g, c in kw.items():
        tot = sum(c.values())
        if tot < 3 or len(g) < 2:
            continue
        (key, ct), = c.most_common(1)
        if ct / tot >= 0.85:
            cands.append({'w': g, 'bk': key[0], 'sd': key[1], 'c1': key[2],
                          'c2': key[3], 'n': tot, 'p': round(ct / tot, 2)})
    cands.sort(key=lambda x: (-x['n'], -len(x['w'])))
    # 去冗余：若长词与短词同频同分类，丢掉短词
    keep = []
    seen = set()
    for k in cands:
        sig = (k['bk'], k['c1'], k['c2'])
        dup = False
        for j in keep:
            if k['w'] in j['w'] and (j['bk'], j['c1'], j['c2']) == sig and j['n'] >= k['n']:
                dup = True
                break
        if not dup:
            keep.append(k)
    return keep


# ══════════════════════════════════════════════════════════
# 归类：三层策略
# ══════════════════════════════════════════════════════════

def classify(r, cfg, learned):
    """返回 (账本, 方向, 一级, 二级, 依据, 置信度)"""
    p = r.get('party') or ''
    g = r.get('goods') or ''
    t = r.get('typ') or ''
    s = ' '.join([p, g, t])
    io = r.get('io') or ''

    # 0) 中性交易
    if io in ('不计收支', '/', ''):
        if '花呗' in p or '白条' in p:
            return ('日常/收入', '支出', '其他支出', '代买', '还款', '高')
        return ('日常/收入', '收入', '转账', '', '内部划转', '高')

    side = '支出' if io == '支出' else '收入'

    # 1) 收入侧特殊
    if side == '收入':
        if not g or g == '/':
            return ('日常/收入', '收入', '群收款', '', 'AA收款', '中')
        return ('日常/收入', '收入', '其他', '转账红包', '朋友往来', '中')

    # 2) 转账给个人 → 社交账本
    if ('转账' in g or '转账' in t) and p:
        for c1, c2 in (('朋友社交', '礼金'), ('亲情往来', '孝敬父母')):
            pass
        return ('情感社交', '支出', '朋友社交', '礼物', '转账待定', '需确认')

    # 3) 语义规则
    for rule in cfg['rules']['规则']:
        if re.search(rule['匹配'], s, re.I):
            return (rule['账本'], rule['方向'], rule['一级'], rule['二级'],
                    rule.get('依据', '规则'), '高')

    # 3.5) 美团订单：按商品名细分
    if '美团' in s or '美团App' in s:
        mt = cfg['rules'].get('美团特殊处理', {})
        for cat, words in (('美业', mt.get('美业', [])), ('运动', mt.get('运动', []))):
            if any(w in g for w in words):
                tgt = (('日常/收入', '支出', '个人形象', '护肤用品') if cat == '美业'
                       else ('日常/收入', '支出', '医疗健康', '健身保健'))
                return tgt + ('美团-%s' % cat, '高')
        if any(w in g for w in mt.get('餐饮', [])):
            return ('日常/收入', '支出', '日常餐饮', '日常三餐', '美团点单-餐饮', '高')

    # 3.6) 人工商家映射（品牌名，无品类词）
    for k, v in cfg['merch'].items():
        if k.startswith('_'):
            continue
        if k in s:
            return (v[0], v[1], v[2], v[3], '商家映射「%s」' % k, '高')

    # 4) 历史学习（长词优先，取首个命中）
    for k in learned:
        if re.search(re.escape(k['w']), s, re.I):
            return (k['bk'], k['sd'], k['c1'], k['c2'],
                    '历史学习「%s」(%d次/纯度%.0f%%)' % (k['w'], k['n'], k['p'] * 100), '高')

    # 4.5) 通用品类词兜底（覆盖历史未见的新商户）
    mt = cfg['rules'].get('品类词')
    if mt:
        for item in mt:
            if re.search(item['匹配'], s, re.I):
                return (item['账本'], item['方向'], item['一级'], item['二级'],
                        item.get('依据', '品类词'), item.get('置信', '中'))

    # 5) 平台兜底
    for rule in cfg['rules']['平台兜底']:
        if re.search(rule['匹配'], s, re.I):
            return (rule['账本'], rule['方向'], rule['一级'], rule['二级'],
                    '平台兜底', '中')

    # 6) 兜底
    return ('日常/收入', '支出', '其他支出', '代买', '待确认', '需确认')


# ══════════════════════════════════════════════════════════
# 输出
# ══════════════════════════════════════════════════════════

ICO_HDR = ['日期', '类型', '金额', '一级分类', '二级分类', '账户1', '账户2',
           '备注', '货币', '标签', '账本', '位置', '退款', '优惠', '手续费',
           '图片链接1', '图片链接2', '图片链接3', '图片附件1', '图片附件2', '图片附件3']


def write_template(rows, path):
    """中间文件（21 列，带账本）。

    注意：这是 iCost 的「导出」格式，App 明确标注「仅供查看，不用于数据恢复」，
    不能直接拿去导入。真正能导入的是 export_split.py 产出的 10 列 / 按账本拆分的文件。
    """
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '收支账单'
    ws.append(ICO_HDR)
    for x in rows:
        amt = -x['amt'] if x['类型'] == '支出' else x['amt']
        ws.append([x['日期'], x['类型'], float(amt), x['c1'], x['c2'],
                   x['acc'], '', x['rem'], 'CNY', '', x['bk'],
                   '', '', '', '', '', '', '', '', '', ''])
    for col, w in zip('ABCDEFGHIJKLMNOPQRSTU',
                      [19, 10, 12, 12, 12, 28, 10, 34, 8, 8, 12, 8, 8, 8, 8] + [10] * 6):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(path)


def write_csv(path, header, rows):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


# ══════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════

def run(wechat, alipay, history, outdir, learn_n=200, purity=0.8):
    cfg = load_config()
    os.makedirs(outdir, exist_ok=True)

    recs = []
    for label, path in (('微信', wechat), ('支付宝', alipay)):
        if path and os.path.exists(path):
            got = _read_any(path)
            for r in got:
                r['src'] = label
            recs += got
            print('  %s %s → %d 笔' % (label, os.path.basename(path), len(got)))
    if not recs:
        print('没有可解析的账单文件')
        return

    # 跨账单合并：同刻同额的多笔收款
    recs = _merge_aa(recs)
    print('  合并 AA 收款后 → %d 笔' % len(recs))

    # 去重
    keep, dup = dedupe(recs, history)
    if dup:
        print('  剔除已入库 %d 笔（%.2f 元）' % (
            len(dup), sum(d[1] for d in dup)))
    recs = keep

    # 历史学习
    learned = learn_keywords(history, cfg) if history else []
    learned = [k for k in learned if k['p'] >= purity][:learn_n]
    print('  提取判别词 %d 条（出现≥%d次，纯度≥%.0f%%）' % (
        len(learned), 10, purity * 100))

    # 归类
    rows = []
    for r in recs:
        bk, sd, c1, c2, why, conf = classify(r, cfg, learned)
        amt = float(r['amt'])
        if r['io'] in ('不计收支', '/', ''):
            typ = '转账' if c1 == '转账' else '还款'
        else:
            typ = '支出' if r['io'] == '支出' else '收入'
        rem = (r.get('goods') or '').strip()
        if not rem or rem == '/' or '二维码' in rem or '转账备注' in rem or '收钱码' in rem:
            rem = (r.get('party') or '').strip()[:24] or '(空)'
        rows.append({
            '日期': str(r['dt'])[:19].replace('-', '/'),
            '类型': typ, 'amt': amt, 'bk': bk, 'sd': sd, 'c1': c1, 'c2': c2,
            'acc': pay_to_account(r.get('pay'), cfg), 'rem': rem[:34],
            'why': why, 'conf': conf, 'src': r.get('src', ''),
            'party': r.get('party', ''), 'goods': r.get('goods', ''),
        })
    rows.sort(key=lambda x: x['日期'])

    # 输出
    p1 = os.path.join(outdir, 'iCost明细_中间文件.xlsx')
    write_template(rows, p1)
    write_csv(os.path.join(outdir, '记账明细.csv'),
              ['日期', '类型', '金额', '账本', '方向', '一级分类', '二级分类',
               '账户', '备注', '归类依据', '置信度', '交易对方'],
              [[x['日期'][:16], x['类型'], x['amt'], x['bk'], x['sd'], x['c1'],
                x['c2'], x['acc'], x['rem'], x['why'], x['conf'], x['party']]
               for x in rows])
    if dup:
        write_csv(os.path.join(outdir, '已剔除-重复记录.csv'),
                  ['日期', '金额', '账单商户', 'iCost已有备注', '处理'],
                  [[d[0].strftime('%Y/%m/%d %H:%M'), d[1], d[2].get('party', ''),
                    d[3][2], '已剔除-不重复导入'] for d in dup])

    # 注意：这里【不】产出导入文件。
    # iCost明细_中间文件.xlsx 是未确认的草稿，必须经过确认页 → split.sh 之后才能导入。

    need = [x for x in rows if x['conf'] == '需确认']
    exp = sum(x['amt'] for x in rows if x['类型'] == '支出')
    inc = sum(x['amt'] for x in rows if x['类型'] == '收入')
    print()
    print('=' * 56)
    print('  总笔数 %d  |  支出 %d 笔 %.2f  |  收入 %d 笔 %.2f'
          % (len(rows), sum(1 for x in rows if x['类型'] == '支出'), exp,
             sum(1 for x in rows if x['类型'] == '收入'), inc))
    print('  自动完成 %d 笔 (%.0f%%)  |  需人工确认 %d 笔'
          % (len(rows) - len(need), (len(rows) - len(need)) / len(rows) * 100, len(need)))
    print('  账本分布: %s' % dict(collections.Counter(x['bk'] for x in rows)))
    print('  输出目录: %s' % outdir)
    print('=' * 56)
    if need:
        print('\n  需确认（金额降序）:')
        for x in sorted(need, key=lambda o: -o['amt']):
            print('    %9.2f  %-18s %s' % (x['amt'], x['party'][:16], x['why']))
    return rows


def _merge_aa(recs):
    """同一天内金额相同、时间相近的多笔小额收入 → 合并为一笔 AA 收款"""
    income = [r for r in recs if r.get('io') == '收入']
    grp = collections.defaultdict(list)
    for r in income:
        try:
            amt = round(abs(float(r['amt'])), 2)
        except (TypeError, ValueError):
            continue
        t = _pdt(r['dt'])
        if t and 0 < amt <= 200:
            grp[(t.strftime('%Y-%m-%d'), amt)].append(r)
    merged, drop = [], set()
    for (day, amt), items in grp.items():
        if len(items) < 3:
            continue
        ts = sorted(_pdt(r['dt']) for r in items)
        if (ts[-1] - ts[0]).total_seconds() > 3600:   # 1 小时内
            continue
        base = dict(items[0])
        base['amt'] = amt * len(items)
        base['party'] = 'AA收款:' + '、'.join(
            (r.get('party') or '').split('(')[0].strip() for r in items)
        base['goods'] = '%d 人各 %.2f' % (len(items), amt)
        merged.append(base)
        drop.update(id(r) for r in items)
    return merged + [r for r in recs if id(r) not in drop]


def main():
    ap = argparse.ArgumentParser(description='iCost 智能记账引擎')
    ap.add_argument('--wechat', help='微信支付账单 xlsx')
    ap.add_argument('--alipay', help='支付宝交易明细 csv')
    ap.add_argument('--history', help='iCost 已导出的历史 xlsx（用于去重与学习）')
    ap.add_argument('--out', default='./output', help='输出目录')
    ap.add_argument('--learn-n', type=int, default=200, help='最多学习多少条判别词')
    ap.add_argument('--purity', type=float, default=0.8, help='判别词纯度阈值')
    a = ap.parse_args()
    if not a.wechat and not a.alipay:
        ap.error('至少要提供 --wechat 或 --alipay')
    print('iCost 智能记账引擎')
    print('-' * 56)
    run(a.wechat, a.alipay, a.history, a.out, a.learn_n, a.purity)


if __name__ == '__main__':
    main()
