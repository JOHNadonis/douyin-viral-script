# -*- coding: utf-8 -*-
"""douyin-viral-script 工具链：句式库查询 / 脚本模板导出 / 校验 / 装配。

子命令：
  list      列出全部句式原型
  show      查看单个句式原型详情
  plan      导出脚本骨架模板（含随机组合的段落安排）
  check     校验脚本：字数区间、段落完整性、重复句
  assemble  把脚本装配成 Excel / Markdown
"""
import argparse
import json
import os
import random
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STRUCT_PATH = os.path.join(BASE, 'assets', 'structures.json')

SEC_ALIAS = {
    'o': 'openers', 'p': 'pains', 's': 'sells',
    't': 'trusts', 'r': 'prices', 'c': 'ctas', 'k': 'skeletons',
    'openers': 'openers', 'pains': 'pains', 'sells': 'sells',
    'trusts': 'trusts', 'prices': 'prices', 'ctas': 'ctas', 'skeletons': 'skeletons',
}


def load():
    with open(STRUCT_PATH, encoding='utf-8') as f:
        return json.load(f)


def cmd_list(args):
    d = load()
    for sec, items in d['sections'].items():
        print(f"\n===== {sec} ({len(items)}) =====")
        for it in items:
            extra = it.get('share') or it.get('freq') or it.get('words') or ''
            print(f"  [{it['id']}] {it['name']:<10} {extra:<8} {it.get('template', ' / '.join(it.get('flow', [])))}")


def cmd_show(args):
    d = load()
    for sec, items in d['sections'].items():
        for it in items:
            if it['id'].upper() == args.id.upper():
                print(json.dumps(it, ensure_ascii=False, indent=1))
                return
    print(f"未找到 id={args.id}", file=sys.stderr)
    sys.exit(1)


def cmd_plan(args):
    d = load()
    secs = d['sections']
    rng = random.Random(args.seed)

    if args.skeleton:
        sk = next((x for x in secs['skeletons'] if x['id'].upper() == args.skeleton.upper()), None)
        if not sk:
            print(f"未找到骨架 {args.skeleton}", file=sys.stderr)
            sys.exit(1)
    else:
        sk = secs['skeletons'][0]

    # 按骨架 flow 拆出需要的段落类型与数量
    need = []
    for step in sk['flow']:
        m = re.match(r'([OPSRTCK])(.+?)(?:×(\d+))?$', step)
        if not m:
            continue
        code, _, cnt = m.groups()
        need.append((code, int(cnt) if cnt else 1))

    plan = {'skeleton': sk, 'product': args.product or '', 'segments': []}
    used = {}  # 跨段落去重：同一 section 已被选用的 id
    for code, cnt in need:
        sec = SEC_ALIAS.get(code.lower())
        pool = secs.get(sec, [])
        if not pool:
            continue
        taken = used.setdefault(sec, set())
        avail = [x for x in pool if x['id'] not in taken]
        if not avail:  # 本段已全部用过，允许复用
            avail = pool[:]
        if args.prefer:
            prefs = [x.strip().upper() for x in args.prefer.split(',')]
            hit = [x for x in avail if x['id'].upper() in prefs]
            if hit:
                avail = hit + [x for x in avail if x not in hit]
        chosen = avail[:cnt] if cnt >= len(avail) else rng.sample(avail, cnt)
        for c in chosen:
            taken.add(c['id'])
            plan['segments'].append({
                'section': sec,
                'id': c['id'],
                'name': c['name'],
                'template': c.get('template', ''),
                'slots': c.get('slots', []),
                'examples': c.get('real', []),
            })

    out = json.dumps(plan, ensure_ascii=False, indent=1)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(out)
        print(f"已导出 {len(plan['segments'])} 个段落 -> {args.out}")
        print(f"骨架: {sk['name']} | 目标字数: {sk['words']}")
        for s in plan['segments']:
            print(f"  {s['id']:<3} {s['name']:<10} {s['template'][:50]}")
    else:
        print(out)


def cmd_check(args):
    """校验脚本。支持 JSON 数组（每条一个脚本）或 {"scripts":[...]}。"""
    scripts = load_scripts(args.scripts)
    lo, hi = (int(x) for x in args.range.split('-')) if args.range else (0, 10 ** 6)
    problems = []
    for i, s in enumerate(scripts):
        txt = s if isinstance(s, str) else s.get('text', '')
        clean = re.sub(r'[\s\n\r]+', '', txt)
        n = len(clean)
        if n < lo:
            problems.append((i, f'SHORT({n}<{lo})', txt[:30]))
        if n > hi:
            problems.append((i, f'LONG({n}>{hi})', txt[:30]))
        if args.require_cta and not re.search(r'小黄车|购物车|下方|左下角|链接', txt):
            problems.append((i, 'NO_CTA', txt[-30:]))
        if args.require_product:
            p = [x for x in args.require_product.split(',') if x]
            miss = [x for x in p if x not in txt]
            if len(miss) == len(p):
                problems.append((i, 'NO_PRODUCT', txt[:30]))

    print(f'共 {len(scripts)} 条脚本，问题 {len(problems)} 条')
    for p in problems:
        print('  ', p[1], '|', p[2])
    if problems and args.strict:
        sys.exit(2)
    return problems


def load_scripts(path):
    with open(path, encoding='utf-8') as f:
        raw = f.read().strip()
    if not raw:
        return []
    try:
        obj = json.loads(raw)
    except json.JSONDecodeError:
        obj = None
    if obj is None:
        return [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if isinstance(obj, dict):
        if 'scripts' in obj:
            return obj['scripts']
        if 'rows' in obj:
            return [r.get('text', '') for r in obj['rows']]
        return [obj.get('text', '')]
    if isinstance(obj, list):
        return obj
    return []


def cmd_assemble(args):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    with open(args.filled, encoding='utf-8') as f:
        data = json.load(f)
    rows = data.get('rows', data) if isinstance(data, dict) else data

    wb = Workbook()
    ws = wb.active
    ws.title = args.sheet or '带货脚本'

    head_fill = PatternFill('solid', fgColor='DCEBF7')
    thin = Side(style='thin', color='BFBFBF')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ['序号', '版本', '开口钩子', '痛点切入', '主打卖点', '信任背书', '价格处理', '结尾方式', '字数', '脚本正文']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, size=11)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bd

    seq = 0
    for r in rows:
        seq += 1
        txt = r.get('text', '')
        n = len(re.sub(r'[\s\n\r]+', '', txt))
        ws.append([seq, r.get('version', ''), r.get('opener', ''), r.get('pain', ''),
                   r.get('sell', ''), r.get('trust', ''), r.get('price', ''),
                   r.get('cta', ''), n, txt])

    widths = [6, 14, 14, 14, 16, 14, 14, 18, 8, 100]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = bd
            cell.alignment = Alignment(vertical='top', wrap_text=True,
                                       horizontal='center' if cell.column in (1, 9) or cell.column <= 8 else 'left')
        ws.row_dimensions[row[0].row].height = 90
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{ws.max_row}'

    if args.product_card and os.path.exists(args.product_card):
        with open(args.product_card, encoding='utf-8') as f:
            card = json.load(f)
        ws2 = wb.create_sheet('产品信息卡')
        for c in range(1, 3):
            cell = ws2.cell(row=1, column=c, value=['项目', '内容'][c - 1])
            cell.font = Font(bold=True, size=11)
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for k, v in card.items():
            ws2.append([k, v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)])
        ws2.column_dimensions['A'].width = 22
        ws2.column_dimensions['B'].width = 100
        for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=2):
            for cell in row:
                cell.border = bd
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            ws2.row_dimensions[row[0].row].height = 40

    wb.save(args.out)
    print(f'已生成 {seq} 条脚本 -> {args.out}')


def cmd_feynman(args):
    """导出费曼知识口播的大纲模板：钩子四选一 + 五段结构 + 类比/翻译/纠误要求。"""
    d = load()
    secs = d['sections']
    hooks = secs.get('feynman_hooks', [])
    flow = (secs.get('feynman_flow') or [{}])[0]
    core = secs.get('feynman_core', [])
    plats = secs.get('platforms', [])
    rng = random.Random(args.seed)

    plat = next((p for p in plats if p['name'] == args.platform), None)
    if args.platform and not plat:
        print(f"未知平台 {args.platform}，可选：{[p['name'] for p in plats]}", file=sys.stderr)
        sys.exit(1)

    if args.hook:
        hook = next((x for x in hooks if x['id'].upper() == args.hook.upper() or x['name'] == args.hook), None)
        if not hook:
            print(f"未知钩子 {args.hook}，可选：{[x['name'] for x in hooks]}", file=sys.stderr)
            sys.exit(1)
    else:
        hook = rng.choice(hooks)

    out = {
        'instruction': '按费曼学习法写口播稿：先讲通机制，再给结论。不要罗列知识点。',
        'topic': args.topic or '',
        'platform': plat or {},
        'hook': {
            'id': hook['id'], 'name': hook['name'], 'desc': hook['desc'],
            'template': hook['template'], 'slots': hook['slots'], 'examples': hook.get('real', []),
        },
        'flow': flow,
        'core_requirements': [
            {'id': c['id'], 'name': c['name'], 'desc': c['desc'], 'how': c.get('how', '')}
            for c in core
        ],
        'why': d.get('feynman', {}),
    }
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"已导出费曼大纲 -> {args.out}")
        print(f"钩子: {hook['id']} {hook['name']}  「{hook['template']}」")
        print(f"平台: {(plat or {}).get('name', '未指定')}  字数: {(plat or flow).get('words', '')}")
        for s in flow.get('steps', []):
            print(f"  {s['no']}. {s['name']:<10}{s['budget']:<10}{s['time']}")
    else:
        print(json.dumps(out, ensure_ascii=False, indent=1))


def cmd_check_feynman(args):
    """校验费曼口播稿：五段是否齐全、有无类比、有无术语翻译、有无纠误、字数。"""
    scripts = load_scripts(args.scripts)
    lo, hi = (int(x) for x in args.range.split('-')) if args.range else (240, 400)
    problems = []
    for i, s in enumerate(scripts):
        if isinstance(s, str):
            txt, ver = s, f'#{i}'
        else:
            txt, ver = s.get('text', ''), s.get('version', f'#{i}')
            if s.get('hook_type'):
                if s['hook_type'] not in ('反问', '反常识', '数字', '冲突'):
                    problems.append((ver, f"HOOK_INVALID({s['hook_type']})", ''))
        n = len(re.sub(r'\s', '', txt))
        if n < lo:
            problems.append((ver, f'SHORT({n}<{lo})', txt[:24]))
        if n > hi:
            problems.append((ver, f'LONG({n}>{hi})', txt[:24]))
        # 类比标记（放宽：兼容「就像X」「其实像X。」「跟X一样」等简式）
        if not re.search(r'就像|好比|打个比方|相当于|跟.{2,10}一样|像.{2,12}(似的|一样|那样|这种|这个)|[其实正]像.{2,12}[。，]', txt):
            problems.append((ver, 'NO_ANALOGY', '缺少生活类比'))
        # 术语翻译
        if not re.search(r'说白了|其实就是|换句话说|简单说|大概是|也就是', txt):
            problems.append((ver, 'NO_TRANSLATE', '缺少大白话翻译'))
        # 误解纠正
        if not re.search(r'很多人以为|不少人以为|你可能会以为|其实不是|并不是|恰恰相反', txt):
            problems.append((ver, 'NO_MISCONCEPTION', '缺少主动纠正误解'))
        # 元叙述：卖货/知识稿都不该出现「这期讲完」这类节目语言
        for w in ('这期讲完', '本期讲完', '今天讲', '这期我们', '来聊聊', '看完你就知道'):
            if w in txt:
                problems.append((ver, f'META_NARRATION:{w}', '出现元叙述，观众会出戏'))
                break
        # CTA（带货型的「小黄车/官网发货」也算，别只认知识型的关注收藏）
        if not re.search(r'关注|收藏|点赞|下期|接着讲|主页|评论区|小黄车|左下角|下方|链接', txt):
            problems.append((ver, 'NO_CTA', '缺少结尾引导'))

    print(f'共 {len(scripts)} 条费曼稿，问题 {len(problems)} 条')
    for p in problems:
        print(f"  [{p[0]}] {p[1]}  {p[2]}")
    if problems and args.strict:
        sys.exit(2)
    return problems



def cmd_short45(args):
    """45秒三段式：输出「文案 + 画面」逐行拍摄脚本模板。"""
    d = load()
    secs = d['sections']
    sk = (secs.get('short45') or [{}])[0]
    hooks = secs.get('hook45', [])
    reasons = secs.get('reasons', [])
    closing = secs.get('closing45', [])
    rng = random.Random(args.seed)

    if args.hook:
        h = next((x for x in hooks if x['id'].upper() == args.hook.upper() or x['name'] == args.hook), None)
        if not h:
            print(f"未知钩子 {args.hook}，可选：{[(x['id'], x['name']) for x in hooks]}", file=sys.stderr)
            sys.exit(1)
    else:
        h = hooks[0]

    n_reason = args.reasons or 3
    chosen = reasons if n_reason >= len(reasons) else rng.sample(reasons, n_reason)
    chosen = sorted(chosen, key=lambda x: x['id'])

    rows = []
    segs = {x['seg']: x for x in sk.get('flow', [])}
    # 开头：钩子 + 场景引入
    rows.append({'seg': '开头', 'time': '0-3秒', 'text': h['template'], 'visual': h['visual'],
                 'role': '痛点阐述', 'goal': '留人：2s出'})
    rows.append({'seg': '开头', 'time': '3-5秒', 'text': '[场景/人群补充一句]', 'visual': '[继续展示痛点场景]',
                 'role': '引入场景', 'goal': '留人：5s完成'})
    # 中段：介绍产品 + 每个购买理由
    rows.append({'seg': '中段', 'time': '5-8秒', 'text': '直到我用了这款[产品名]', 'visual': '产品镜头',
                 'role': '介绍产品', 'goal': '商品点击'})
    span = max(1, 22 // max(1, len(chosen)))
    for i, r in enumerate(chosen):
        st = 8 + i * span
        rows.append({'seg': '中段', 'time': f'{st}-{st+span}秒', 'text': r['template'],
                     'visual': r['visual'], 'role': f"购买理由（{r['name']}）", 'goal': '商品转化'})
    # 结尾
    for c in closing:
        rows.append({'seg': '结尾', 'time': '30-45秒', 'text': c['template'], 'visual': c['visual'],
                     'role': c['name'], 'goal': '商品转化' if c['id'] == 'E1' else '临门一脚'})

    out = {
        'instruction': '按 rows 顺序写：每行都要有文案和画面。文案是口播，画面是拍什么/贴什么图。不要出现「这期讲完」这类元叙述。',
        'skeleton': sk, 'product': args.product or '',
        'hook': {'id': h['id'], 'name': h['name'], 'fit': h['fit'], 'why': h['why']},
        'rows': rows,
    }
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"已导出 45秒脚本模板 -> {args.out}")
        print(f"钩子: {h['id']} {h['name']}（适合：{h['fit']}）")
        print(f"购买理由 {len(chosen)} 个: {[x['name'] for x in chosen]}")
        print(f"共 {len(rows)} 行（文案+画面配对）")
    else:
        print(json.dumps(out, ensure_ascii=False, indent=1))


def _check_one(rows, lo, hi, require_cta, label=''):
    """校验单条 45 秒脚本，返回问题列表。"""
    problems = []
    texts = [r.get('text', '') for r in rows]
    total = len(re.sub(r'\s', '', ''.join(texts)))
    pre = f'[{label}] ' if label else ''

    if total < lo:
        problems.append(f'{pre}TOTAL_SHORT({total}<{lo})')
    if total > hi:
        problems.append(f'{pre}TOTAL_LONG({total}>{hi})')

    segs = [r.get('seg') for r in rows]
    for need in ('开头', '中段', '结尾'):
        if need not in segs:
            problems.append(f'{pre}MISSING_SEG:{need}')
    if segs and segs[0] != '开头':
        problems.append(f'{pre}ORDER:开头不在最前')
    if segs and segs[-1] != '结尾':
        problems.append(f'{pre}ORDER:结尾不在最后')

    for i, r in enumerate(rows, 1):
        v = (r.get('visual') or '').strip()
        if v in ('', '—') and r.get('role') != '行为指引':
            problems.append(f"{pre}NO_VISUAL:第{i}行「{r.get('text','')[:12]}」")
    for r in rows:
        for w in ('这期讲完', '本期', '今天讲', '来聊聊'):
            if w in r.get('text', ''):
                problems.append(f'{pre}META_NARRATION:{w}')
                break
    if require_cta and not re.search(r'小黄车|下方|左下角|链接|同款', ''.join(texts)):
        problems.append(f'{pre}NO_CTA')

    return problems, total


def cmd_check45(args):
    """校验45秒脚本。rows 里带 product 字段时自动按商品分组，逐条独立校验。"""
    with open(args.script, encoding='utf-8') as f:
        d = json.load(f)

    lo, hi = (int(x) for x in (args.range or '180-260').split('-'))
    problems = []

    # 优先用 products 结构；否则看 rows 里有没有 product 字段
    groups = []
    if isinstance(d, dict) and d.get('products'):
        for p in d['products']:
            groups.append((p.get('product', ''), p.get('rows', [])))
    else:
        rows = d.get('rows', d if isinstance(d, list) else [])
        if rows and any(r.get('product') for r in rows):
            order, bucket = [], {}
            for r in rows:
                k = r.get('product') or '未命名'
                if k not in bucket:
                    bucket[k] = []
                    order.append(k)
                bucket[k].append(r)
            groups = [(k, bucket[k]) for k in order]
        else:
            groups = [('', rows)]

    # 合并单元格回填：seg/time 为空时沿用上一行的值（Excel 合并后只有首格有值）
    for _, rows in groups:
        last_seg = last_time = None
        for r in rows:
            if r.get('seg'):
                last_seg = r['seg']
            elif last_seg:
                r['seg'] = last_seg
            if r.get('time'):
                last_time = r['time']
            elif last_time:
                r['time'] = last_time

    multi = len(groups) > 1
    for name, rows in groups:
        probs, total = _check_one(rows, lo, hi, args.require_cta,
                                  label=name if multi else '')
        problems.extend(probs)
        mark = '✓' if not probs else f'✗ {len(probs)}'
        print(f"{name or '脚本':<12} {len(rows):>2} 行  {total:>4} 字  {mark}")

    print(f'\n共 {len(groups)} 条脚本，合计 {sum(len(r) for _, r in groups)} 行，问题 {len(problems)} 条')
    for p in problems:
        print('  ', p)
    if problems and args.strict:
        sys.exit(2)
    return problems


def cmd_assemble45(args):
    """把45秒脚本装配成 Excel：结构/时间/文案/画面/作用/目的 六列，跟用户框架图一致。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    with open(args.filled, encoding='utf-8') as f:
        d = json.load(f)
    rows = d.get('rows', [])
    wb = Workbook()
    ws = wb.active
    ws.title = args.sheet or '45秒脚本'

    head = PatternFill('solid', fgColor='DCEBF7')
    thin = Side(style='thin', color='BFBFBF')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = ['结构', '时间', '文案', '画面', '作用', '目的']
    ws.append(headers)
    for c in range(1, 7):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, size=11)
        cell.fill = head
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bd

    for r in rows:
        ws.append([r.get('seg'), r.get('time'), r.get('text'), r.get('visual'),
                   r.get('role'), r.get('goal')])
    for i, w in enumerate([10, 12, 62, 42, 20, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=6):
        for cell in row:
            cell.border = bd
            cell.alignment = Alignment(vertical='center', wrap_text=True,
                                       horizontal='center' if cell.column in (1, 2, 6) else 'left')
        ws.row_dimensions[row[0].row].height = 46
    # 结构列纵向合并
    start = 2
    for r in range(3, ws.max_row + 2):
        prev = ws.cell(row=r - 1, column=1).value if r - 1 >= 2 else None
        cur = ws.cell(row=r, column=1).value if r <= ws.max_row else None
        if cur != prev:
            if r - 1 > start:
                ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
                ws.merge_cells(start_row=start, start_column=2, end_row=r - 1, end_column=2)
            start = r
    ws.freeze_panes = 'A2'
    wb.save(args.out)
    print(f'已生成 {len(rows)} 行拍摄脚本 -> {args.out}')


def main():
    ap = argparse.ArgumentParser(prog='dvt', description='抖音带货脚本工具链')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('short45', help='导出45秒三段式拍摄脚本模板')
    p.add_argument('--product', help='产品名')
    p.add_argument('--hook', help='H1-H4 或 痛点型/爽点型/强冲突/痒点型')
    p.add_argument('--reasons', type=int, default=3, help='中段放几个购买理由，默认3')
    p.add_argument('--seed', type=int, default=0)
    p.add_argument('--out')
    p.set_defaults(func=cmd_short45)

    p = sub.add_parser('check45', help='校验45秒脚本（行数/字数/画面/元叙述）')
    p.add_argument('--script', required=True)
    p.add_argument('--range', help='总字数区间，默认 180-260')
    p.add_argument('--require-cta', action='store_true')
    p.add_argument('--strict', action='store_true')
    p.set_defaults(func=cmd_check45)

    p = sub.add_parser('assemble45', help='装配45秒脚本 Excel（结构/时间/文案/画面/作用/目的）')
    p.add_argument('--filled', required=True)
    p.add_argument('--out', required=True)
    p.add_argument('--sheet')
    p.set_defaults(func=cmd_assemble45)

    p = sub.add_parser('feynman', help='导出费曼知识口播大纲')
    p.add_argument('--topic', help='主题')
    p.add_argument('--platform', help='抖音 / 视频号 / 小红书')
    p.add_argument('--hook', help='FH1-FH4 或 反问/反常识/数字/冲突')
    p.add_argument('--seed', type=int, default=0)
    p.add_argument('--out')
    p.set_defaults(func=cmd_feynman)

    p = sub.add_parser('check-feynman', help='校验费曼稿五段要素')
    p.add_argument('--scripts', required=True)
    p.add_argument('--range', help='字数区间，默认 240-400')
    p.add_argument('--strict', action='store_true')
    p.set_defaults(func=cmd_check_feynman)

    p = sub.add_parser('list', help='列出全部句式原型')
    p.set_defaults(func=cmd_list)

    p = sub.add_parser('show', help='查看单个句式原型')
    p.add_argument('--id', required=True, help='如 O1 / S2 / C1')
    p.set_defaults(func=cmd_show)

    p = sub.add_parser('plan', help='导出脚本骨架模板')
    p.add_argument('--skeleton', help='K1 标准60秒 / K2 短版30秒 / K3 长版90秒')
    p.add_argument('--product', help='产品名')
    p.add_argument('--prefer', help='优先选用这些 id，逗号分隔')
    p.add_argument('--seed', type=int, default=0, help='随机种子，固定可复现')
    p.add_argument('--out')
    p.set_defaults(func=cmd_plan)

    p = sub.add_parser('check', help='校验脚本字数/结构')
    p.add_argument('--scripts', required=True)
    p.add_argument('--range', help='字数区间，如 280-330')
    p.add_argument('--require-cta', action='store_true', help='要求含促单句')
    p.add_argument('--require-product', help='要求含产品词，逗号分隔')
    p.add_argument('--strict', action='store_true')
    p.set_defaults(func=cmd_check)

    p = sub.add_parser('assemble', help='装配 Excel')
    p.add_argument('--filled', required=True)
    p.add_argument('--out', required=True)
    p.add_argument('--product-card')
    p.add_argument('--sheet')
    p.set_defaults(func=cmd_assemble)

    args = ap.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
