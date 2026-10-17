#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iCost 导入文件生成器（按账本拆分 + 严格对齐官方模板）

官方导入模板（App 内导出，icost_template.xlsx）规范：
  sheet 名 : icost_template
  列数     : 10
  表头     : 日期 类型 金额 一级分类 二级分类 账户1 账户2 备注 货币 标签
  日期格式 : 2011年01月11日 12:00:00
  金额     : 正数（正负由「类型」决定，不靠符号）
  类型     : 支出 / 收入 / 转账（还款等由 App 内部类型承担）
  标签     : #标签1#标签2

注意：iCost 的「导出」文件是 21 列（带账本/位置/退款等），文件名写着
「仅供查看，不用于数据恢复」，即导出格式 ≠ 导入格式。本脚本只按官方
导入模板的 10 列输出，并按账本拆分（iCost 是按账本导入的）。
"""

import os
import re
import csv
import json
import sys
import datetime
import collections

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 常量
SHEET_NAME = 'icost_template'
ICO_HDR = ['日期', '类型', '金额', '一级分类', '二级分类',
           '账户1', '账户2', '备注', '货币', '标签']

VALID_TYPES = {'支出', '收入', '转账', '还款',
               '退款入账', '报销入账', '待报销', '已报销', '退款收入',
               '收款', '借出', '借入'}

# 冲减类：金额是负向抵消，分类必须跟「原订单」走，因此按支出侧分类树校验。
# （历史 66 笔退款入账没有一笔落在「其他」上，全是 交通出行/火车飞机、
#   购物/服饰、住房支出/房租 这类原订单分类）
OFFSET_TYPES = {'退款入账', '报销入账', '待报销', '已报销', '退款收入'}

NEED_ACC2 = {'转账', '还款', '收款', '借出', '借入'}

BOOK_ORDER = ['日常/收入', '情感社交', '成长提升', '大额固定', '储蓄理财']
BOOK_FNAME = {'日常/收入': '日常-收入', '情感社交': '情感社交',
              '成长提升': '成长提升', '大额固定': '大额固定',
              '储蓄理财': '储蓄理财'}

# 配置目录定位：脚本可能在项目里（icost-system/core/），也可能在技能包里（scripts/core/）。
# 统一按「脚本上一级目录的 config/」找，和 engine.py 保持一致，这样整包可以随便挪位置。
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG = os.path.join(BASE, 'config')


def load_cfg():
    with open(os.path.join(CFG, 'categories.json'), encoding='utf-8') as f:
        tree = json.load(f)
    with open(os.path.join(CFG, 'accounts.json'), encoding='utf-8') as f:
        acc_cfg = json.load(f)
    accs = set()
    for grp in acc_cfg.values():
        for a in grp.get('accounts', []):
            accs.add(a['name'])
    fix = {}
    p = os.path.join(CFG, 'import_fix.json')
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            fix = json.load(f)
    return tree, accs, fix


# ---------------------------------------------------------------- 工具
def fmt_date(v):
    """把任意常见日期写法统一成 2011年01月11日 12:00:00"""
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y年%m月%d日 %H:%M:%S')
    s = str(v).strip()
    m = re.match(r'^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?[ T]?(\d{1,2}:\d{2}(:\d{2})?)?', s)
    if not m:
        return s
    y, mo, d = m.group(1), int(m.group(2)), int(m.group(3))
    t = m.group(4) or '00:00:00'
    if len(t.split(':')) == 2:
        t += ':00'
    return f'{y}年{mo:02d}月{d:02d}日 {t}'


def cell(v):
    """空值统一成空字符串，绝不写 None"""
    if v is None:
        return ''
    return str(v).strip()


def num(v):
    try:
        return round(abs(float(v)), 2)
    except Exception:
        return 0.0


# ---------------------------------------------------------------- xlsx 写出
OFFICIAL_TPL = os.path.join(CFG, 'icost_template.xlsx')

NS = ('http://schemas.openxmlformats.org/spreadsheetml/2006/main')


def _esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


class _SST:
    """共享字符串表：官方模板用的是 <si><t>xxx</t></si>，
    而 openpyxl 会写成 <si><r><t>xxx</t></r></si> 甚至 t="inlineStr"+数字实体，
    这几种写法很多 iOS 端的 xlsx 解析器不认，必须自己拼。"""

    def __init__(self):
        self.items = []
        self.idx = {}

    def add(self, v):
        if v not in self.idx:
            self.idx[v] = len(self.items)
            self.items.append(v)
        return self.idx[v]

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               f'<sst xmlns="{NS}" count="{len(self.items)}" '
               f'uniqueCount="{len(self.items)}">']
        for s in self.items:
            if s != s.strip():
                out.append(f'<si><t xml:space="preserve">{_esc(s)}</t></si>')
            else:
                out.append(f'<si><t>{_esc(s)}</t></si>')
        out.append('</sst>')
        return ''.join(out)


def _sheet_xml(rows, sst):
    n = len(rows) + 1
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           f'<worksheet xmlns="{NS}"><sheetPr/>',
           f'<dimension ref="A1:J{n}"/>',
           '<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
           '<selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>',
           '<sheetFormatPr defaultColWidth="9.64285714285714" '
           'defaultRowHeight="13.1" outlineLevelRow="3"/>',
           '<sheetData>']
    for ri, r in enumerate([ICO_HDR] + rows, start=1):
        cells = []
        for ci, col in enumerate('ABCDEFGHIJ'):
            v = r[ci] if isinstance(r, (list, tuple)) else r[ICO_HDR[ci]]
            if v is None or v == '':
                continue                       # 空单元格整个不写，跟官方一致
            ref = f'{col}{ri}'
            if isinstance(v, (int, float)):
                n = int(v) if float(v) == int(v) else v   # 官方模板里 40/1000 都是整数
                cells.append(f'<c r="{ref}"><v>{n}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="s"><v>{sst.add(str(v))}</v></c>')
        out.append(f'<row r="{ri}" spans="1:10">' + ''.join(cells) + '</row>')
    out.append('</sheetData>')
    out.append('<pageMargins left="0.75" right="0.75" top="1" bottom="1" '
               'header="0.511805555555556" footer="0.511805555555556"/>')
    out.append('<pageSetup paperSize="9" orientation="portrait"/>')
    out.append('<headerFooter/></worksheet>')
    return ''.join(out)


def write_xlsx_like_official(rows, path):
    """以官方 icost_template.xlsx 为骨架，只替换 sheet1.xml 和 sharedStrings.xml。

    这样产出的文件与 App 自己给的模板在结构上完全一致（同样的 part 列表、
    同样的 Content_Types、同样的 styles/theme、plain <si><t> 字符串表）。
    """
    import zipfile
    data = [[r[c] for c in ICO_HDR] for r in rows]
    sst = _SST()
    sheet = _sheet_xml(data, sst)

    src = zipfile.ZipFile(OFFICIAL_TPL, 'r')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            nm = item.filename
            if nm == 'xl/worksheets/sheet1.xml':
                z.writestr(item, sheet.encode('utf-8'))
            elif nm == 'xl/sharedStrings.xml':
                z.writestr(item, sst.xml().encode('utf-8'))
            else:
                z.writestr(item, src.read(nm))
    src.close()


def write_manual_list(rows, path):
    """冲减类（退款入账/报销入账）iCost 导入识别不了，只能 App 内手工建。

    输出一份逐步操作清单：每笔给出 账本 / 类型 / 金额 / 分类 / 账户 / 备注，
    并按类型给出「怎么点」的最短路径。
    """
    rows = sorted(rows, key=lambda x: -x['金额'])
    total = sum(r['金额'] for r in rows)
    L = ["# 手工补录清单（退款 / 报销）\n",
         "iCost 的导入只认 `支出 / 收入 / 转账` 三种类型（官方模板里也只有这三种），",
         "**`退款入账` `报销入账` 导入不进去**。这几笔钱已经真的进了你的账户，",
         "如果不补录，账户余额会少记、储蓄率会算错。\n",
         f"共 **{len(rows)} 笔 / {total:,.2f} 元**，在 App 里手工建即可（约 2 分钟）。\n",
         "| # | 日期 | 类型 | 金额 | 账本 | 一级/二级 | 账户 | 备注 |",
         "|---:|---|---|---:|---|---|---|---|"]
    for i, r in enumerate(rows, 1):
        c2 = '/' + r['二级分类'] if r['二级分类'] else ''
        L.append(f"| {i} | {r['日期'][:11]} | **{r['类型']}** | {r['金额']:,.2f} | "
                 f"{r['账本']} | {r['一级分类']}{c2} | {r['账户1']} | {r['备注']} |")
    spec = [r for r in rows if r.get('操作')]
    if spec:
        L.append("\n### 每笔的具体点法\n")
        for i, r in enumerate(rows, 1):
            if r.get('操作'):
                L.append(f"{i}. **{r['金额']:,.2f} 元（{r['类型']}｜{r['账本']}）**\n   {r['操作']}")

    L.append("\n## 怎么点（按类型）\n")
    if any(r['类型'] == '退款入账' for r in rows):
        L += ["### 退款入账\n",
              "**最短路径（能自动关联原订单，推荐）**\n",
              "1. 切到备注里写的那个账本\n",
              "2. 找到原订单那笔支出\n",
              "3. **长按该笔 → 退款** → 填金额 → 保存\n",
              "   （分类自动沿用原订单，不用手填）\n",
              "找不到原订单时，才新建：+ → 类型选「退款入账」→ 按上表填分类和账户。\n"]
    if any(r['类型'] == '报销入账' for r in rows):
        L += ["### 报销入账\n",
              "**如果原订单已经标记成「待报销」**\n",
              "1. 切到该账本 → 打开那笔待报销记录\n",
              "2. 状态改成「已报销」→ 系统会自动生成配对的报销入账\n",
              "**没有对应待报销记录时**\n",
              "1. 切到上表的账本 → 新建 → 类型选「报销入账」\n",
              "2. 分类、账户、备注按上表填\n"]

    L += ["\n## 为什么要手工，不能直接导成「收入」\n",
          "这几笔本质是**冲减**（把之前多记的支出抵回来），不是新赚的钱。",
          f"如果导成收入，本月收入会虚增 {total:,.2f} 元，支出也不会被冲回，",
          "储蓄率和分类口径全都会失真 —— 比不记还糟。\n"]
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


# ---------------------------------------------------------------- 主流程
def build(src_xlsx, outdir):
    tree, known_accs, fix = load_cfg()
    cat_map = fix.get('cat_map', {})
    patch = {p['key']: p for p in fix.get('patch', [])}
    drop = {d['key']: d for d in fix.get('drop', [])}
    # 官方导入模板只演示了 支出/收入/转账。实测 iCost 导入**识别不了**
    # 退款入账 / 报销入账 这类冲减类型（2026-10-17 用户实测反馈）。
    # 因此默认把它们从导入文件里剔除，改走「App 内手工补录」，
    # 这样既不会虚增收入，也不会污染分类口径。
    off_importable = fix.get('_导入选项', {}).get('冲减类型可导入', False)

    wb = openpyxl.load_workbook(src_xlsx)
    ws = wb.active
    hdr = [c.value for c in ws[1]]
    raw = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)]

    issues = []      # (级别, 行, 说明)
    rows = []
    dropped = []

    for i, r in enumerate(raw, start=2):
        key = f"{r.get('日期')}|{r.get('金额')}"
        if key in drop:
            dropped.append((key, drop[key]['原因']))
            issues.append(('修复', i, f"已剔除：{key}（{drop[key]['原因']}）"))
            continue

        d = {
            '日期': fmt_date(r.get('日期')),
            '类型': cell(r.get('类型')),
            '金额': num(r.get('金额')),
            '一级分类': cell(r.get('一级分类')),
            '二级分类': cell(r.get('二级分类')),
            '账户1': cell(r.get('账户1')),
            '账户2': cell(r.get('账户2')),
            '备注': cell(r.get('备注')),
            '货币': cell(r.get('货币')) or 'CNY',
            '标签': cell(r.get('标签')),
            '账本': cell(r.get('账本')) or '日常/收入',
        }

        # 1) 人工修正表
        if key in patch:
            for k, v in patch[key]['set'].items():
                d[k] = v
            issues.append(('修复', i, f"{key} → {patch[key]['set']}（{patch[key]['原因']}）"))

        # 2) 类型合法性
        if d['类型'] not in VALID_TYPES:
            issues.append(('错误', i, f"类型「{d['类型']}」不在 支出/收入/转账/还款 中"))

        # 3) 分类树校验
        bk, sd = d['账本'], None
        if d['类型'] == '支出' or d['类型'] in OFFSET_TYPES:
            sd = '支出'
        elif d['类型'] == '收入':
            sd = '收入'
        if sd and bk in tree:
            c1 = d['一级分类']
            if c1 in cat_map:
                mp = cat_map[c1]
                issues.append(('修复', i, f"已归档分类「{c1}」→「{mp['to']}」（{mp['原因']}）"))
                c1 = d['一级分类'] = mp['to']
                d['二级分类'] = mp.get('c2', '')
            if c1 not in tree[bk][sd]:
                issues.append(('错误', i, f"账本「{bk}」{sd}侧没有一级分类「{c1}」"))
            else:
                subs = tree[bk][sd][c1]
                c2 = d['二级分类']
                if c2 and c2 not in subs:
                    issues.append(('修复', i, f"二级分类「{c2}」不属于「{c1}」，已清空"))
                    d['二级分类'] = ''
                if not subs and c2:
                    d['二级分类'] = ''

        # 4) 账户校验
        for k in ('账户1', '账户2'):
            v = d[k]
            if v and v not in known_accs:
                issues.append(('警告', i, f"{k}「{v}」不在账户列表中（可能是幽灵账户）"))
        if d['类型'] in NEED_ACC2 and not d['账户2']:
            issues.append(('错误', i, f"{d['类型']}缺少账户2（转出/转入必须成对）"))

        rows.append(d)

    # 5) 二次去重（日期+金额+类型+备注完全相同）
    seen = collections.Counter((r['日期'], r['金额'], r['类型'], r['备注']) for r in rows)
    rows2 = []
    for r in rows:
        k = (r['日期'], r['金额'], r['类型'], r['备注'])
        if seen[k] > 1 and any(x['日期'] == r['日期'] and x['金额'] == r['金额']
                               and x['类型'] == r['类型'] and x['备注'] == r['备注'] for x in rows2):
            dropped.append((f"{r['日期']}|{r['金额']}", '完全重复行'))
            continue
        rows2.append(r)
    rows = rows2

    # ---------------------------------------------------------- 冲减类分流
    off_rows = [r for r in rows if r['类型'] in OFFSET_TYPES]
    if off_importable:
        imp_rows = rows
        excluded = []
    else:
        imp_rows = [r for r in rows if r['类型'] not in OFFSET_TYPES]
        excluded = off_rows
        for r in excluded:
            issues.append(('剔除', '—', f"{r['日期'][:11]} {r['类型']} {r['金额']:,.2f}"
                                        f" → 转为 App 内手工补录（导入认不了这个类型）"))

    # ---------------------------------------------------------- 按账本输出
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):          # 清掉上一轮的产物，避免残留旧账本文件
        if re.match(r'^\d{2}_iCost导入_', f):
            os.remove(os.path.join(outdir, f))
    by_book = collections.defaultdict(list)
    for r in imp_rows:
        by_book[r['账本']].append(r)

    order = [b for b in BOOK_ORDER if b in by_book] + \
            [b for b in by_book if b not in BOOK_ORDER]

    files = []
    for idx, bk in enumerate(order, 1):
        rs = sorted(by_book[bk], key=lambda x: x['日期'])
        base = f'{idx:02d}_iCost导入_{BOOK_FNAME.get(bk, bk)}'
        xp = os.path.join(outdir, base + '.xlsx')
        cp = os.path.join(outdir, base + '.csv')

        write_xlsx_like_official(rs, xp)

        with open(cp, 'w', newline='', encoding='utf-8-sig') as f:
            wr = csv.writer(f)
            wr.writerow(ICO_HDR)
            for r in rs:
                wr.writerow([r['日期'], r['类型'], r['金额'], r['一级分类'], r['二级分类'],
                             r['账户1'], r['账户2'], r['备注'], r['货币'], r['标签']])
        files.append((bk, xp, cp, rs))

    # ---------------------------------------------------------- 诊断文件
    # 1) 原样复制官方模板（如果连它都导入失败，说明问题在导入方式，不在文件）
    if os.path.exists(OFFICIAL_TPL):
        import shutil
        shutil.copy(OFFICIAL_TPL, os.path.join(outdir, '00_诊断_官方原版模板.xlsx'))
    # 2) 只含 1 笔真实记录的探针文件，用来快速试错
    if rows:
        probe = [x for x in rows if x['类型'] == '支出'] or rows
        write_xlsx_like_official([probe[0]], os.path.join(outdir, '00_诊断_单笔探针.xlsx'))
    # 3) 冲减类手工补录清单（iCost 导入识别不了 退款入账/报销入账，只能手工建）
    if excluded:
        write_manual_list(excluded, os.path.join(outdir, '手工补录清单.md'))

    # ---------------------------------------------------------- 校验报告
    rep = [f"# iCost 导入文件校验报告\n",
           f"生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M}\n",
           f"源文件：`{os.path.basename(src_xlsx)}`\n",
           "\n## 一、格式对齐（对照 App 官方导入模板）\n",
           "| 项 | 官方模板 | 之前的文件 | 现在 |",
           "|---|---|---|---|",
           "| sheet 名 | icost_template | 收支账单 | ✅ icost_template |",
           "| 列数 | 10 | 21（导出格式） | ✅ 10 |",
           "| 日期 | 2011年01月11日 12:00:00 | 2026/09/22 01:39:50 | ✅ 中文格式 |",
           "| 金额 | 正数 | 支出为负数 | ✅ 全部正数 |",
           "| 空单元格 | 空 | None | ✅ 空字符串 |",
           "| 账本 | 不含（导入时选） | 混入 3 个账本 | ✅ 按账本拆成独立文件 |\n"]

    rep.append("\n## 二、拆分结果\n")
    rep.append("| 文件 | 账本 | 笔数 | 支出 | 收入 | 退款/报销冲减 | 转账·还款 |")
    rep.append("|---|---|---:|---:|---:|---:|---:|")
    for bk, xp, cp, rs in files:
        ex = sum(r['金额'] for r in rs if r['类型'] == '支出')
        inc = sum(r['金额'] for r in rs if r['类型'] == '收入')
        off = sum(r['金额'] for r in rs if r['类型'] in OFFSET_TYPES)
        tr = sum(1 for r in rs if r['类型'] in NEED_ACC2)
        rep.append(f"| `{os.path.basename(xp)}` | {bk} | {len(rs)} | "
                   f"{ex:,.2f} | {inc:,.2f} | {off:,.2f} | {tr} |")

    # ---- 退款 / 报销 专项 ----
    off_rows = [r for r in rows if r['类型'] in OFFSET_TYPES]
    if off_rows:
        rep.append("""
## 三、退款与报销的关联处理

### iCost 内部的机制

| 类型 | 作用 |
|---|---|
| 待报销 | 支出标记「这笔钱以后有人给我报」 |
| 已报销 | 报销款到账后，原支出记录**改状态**（不新建记录） |
| 报销入账 | 与「已报销」配套生成的入账，金额一一对应 |
| 退款入账 | 带**原订单分类**做冲减 |

iCost 的关联方式是「**类型 + 金额 + 分类**」三件套配对，**不是 ID**。

### 导入层面做不到硬链接

10 列导入模板里既没有「关联 ID」，也没有导出格式里的「退款」列。
所以**文件导入无法建立指向原订单的指针**，能做的只有三件事：

1. 类型用 `退款入账` / `报销入账` → 让 iCost 冲减，而不是计入收入
2. 一级/二级分类**沿用原订单的分类** → 冲减才能落在正确的分类上
   （不要塞进「其他」，否则冲减落不到原分类上）
3. 把原订单信息（日期 / 商户 / 金额）写进备注 → 供日后人工反查

### 本批次的冲减记录

| 日期 | 金额 | 修正后类型 | 账本 · 分类 | 备注里的关联信息 |
|---|---:|---|---|---|
""")
        for r in sorted(off_rows, key=lambda x: -x['金额']):
            c2 = '/' + r['二级分类'] if r['二级分类'] else ''
            rep.append(f"| {r['日期'][:10]} | {r['金额']:,.2f} | **{r['类型']}** | "
                       f"{r['账本']} · {r['一级分类']}{c2} | {r['备注']} |")

        real_inc = sum(r['金额'] for r in rows if r['类型'] == '收入')
        rep.append(f"""
**影响**：这 {len(off_rows)} 笔合计 **{sum(r['金额'] for r in off_rows):,.2f} 元**。
记成收入会让本月收入虚增这么多；改成冲减后，本月真实收入只剩 **{real_inc:,.2f} 元**，
支出也相应被冲回，储蓄率才准。

### ⚠️ 但这 {len(off_rows)} 笔已从导入文件里剔除

**实测结论：iCost 导入识别不了 `退款入账` / `报销入账`。**
官方模板的示例行只有 `支出 / 收入 / 转账` 三种，导入解析器只认这三种。

所以它们**不进任何一个按账本拆分的文件**，改走 `手工补录清单.md`
（清单里连「长按原订单 → 退款」这种点击路径都写好了）。

> 为什么不干脆导成「收入」？因为导成收入 = 虚增收入 {sum(r['金额'] for r in off_rows):,.2f} 元、
> 支出又不会被冲回，储蓄率和分类口径双重失真，**比不记还糟**。

### 账本对齐（最容易漏的一件事）

**冲减记录必须和原订单在同一个账本**，跨账本冲不掉，等于白记。
核对方法：找到原订单那笔支出，看它在哪个账本，冲减就填哪个账本。
如果原订单所在账本的分类树里没有对应分类（比如账本里根本没配「交通出行」），
就把冲减留在原订单账本，不要硬挪。
""")

    rep.append("""
## 四、xlsx 内部写法（第二轮排查才发现的真正原因）

光有 10 列表头还不够，openpyxl 生成的 xlsx 在 XML 层面和官方模板不一样：

| 内部细节 | 官方模板 | openpyxl 产出 | 现在 |
|---|---|---|---|
| 字符串表 | `<si><t>日期</t></si>` | `<si><r><t>日期</t></r></si>`（富文本 run） | ✅ 前者 |
| 单元格类型 | `t="s"` + 共享索引 | 有时退化成 `t="inlineStr"` | ✅ `t="s"` |
| 中文编码 | 真 UTF-8 汉字 | 有时写成 `&#26085;` 数字实体 | ✅ 真汉字 |
| dimension | `A1:J4` | `J4`（缺起点） | ✅ `A1:Jn` |
| 金额 | `40` | `40.0` | ✅ 整数去小数点 |

现在的实现是**直接拿官方模板当骨架**替换 sheet1.xml / sharedStrings.xml，
其余 part（Content_Types、styles、theme、rels）原样保留，结构上与 App 自带模板一致。

## 五、先跑这几个诊断文件

| 文件 | 用途 |
|---|---|
| `00_诊断_官方原版模板.xlsx` | **App 自带的原版模板，一字未改**。如果连它都解析失败 → 问题在导入方式（比如要通过「文件」App 分享进 iCost，或 iCost 版本不支持），不在我们的文件 |
| `00_诊断_单笔探针.xlsx` | 只有 1 笔真实记录。先导它：成功 → 说明格式对了，直接继续导 01/02/03；失败 → 说明数据里有某个字段不兼容 |

建议顺序：**官方原版 → 单笔探针 → 01/02/03**。

> `退款入账` / `报销入账` 已实测**导入不进去**（只认 支出/收入/转账），
> 这类冲减记录已剔除并转 `手工补录清单.md`，不用再试了。
""")

    rep.append("\n## 六、已自动修复\n")
    rep.append("| 级别 | 行 | 说明 |")
    rep.append("|---|---:|---|")
    for lv, i, msg in issues:
        rep.append(f"| {lv} | {i} | {msg} |")
    if not issues:
        rep.append("| — | — | 无 |")

    rep.append("\n## 七、剔除的记录\n")
    rep.append("| 标识 | 原因 |")
    rep.append("|---|---|")
    for k, why in dropped:
        rep.append(f"| {k} | {why} |")
    if not dropped:
        rep.append("| — | 无 |")

    rep.append("\n## 八、导入步骤\n")
    for bk, xp, cp, rs in files:
        rep.append(f"1. iCost → 切换到 **{bk}** 账本 → 导入 → 选择 "
                   f"`{os.path.basename(xp)}`（若 xlsx 仍失败，改试同名 csv）")
    rep.append("\n> 每个账本单独导一次，不要一次性选多个文件。\n")

    rp = os.path.join(outdir, '导入校验报告.md')
    with open(rp, 'w', encoding='utf-8') as f:
        f.write('\n'.join(rep))

    print(f'账本拆分：{[(b, len(rs)) for b, _, _, rs in files]}')
    print(f'修复/警告 {len(issues)} 条，剔除 {len(dropped)} 条')
    print(f'输出目录：{outdir}')
    print(f'报告：{rp}')
    return files, issues, dropped


if __name__ == '__main__':
    # 默认工作目录 = 系统目录的上一级（和 run.sh / split.sh 的约定一致）
    WORK = os.path.dirname(BASE)
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        WORK, '已确认明细.xlsx')
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(WORK, '导入文件')
    build(src, out)
