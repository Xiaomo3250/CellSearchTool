"""问题点跟踪 — 从docx报告提取问题点/RF清单/新建清单，写入跟踪表
依赖: 工参管理器 cache.json，用于查询运营商和设备商
"""
import os, re, json
from docx import Document
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(ROOT, '..', '0513-问题点跟踪表', '0512_江布拉克道路质量分析优化报告.docx')
TRACKER = os.path.join(ROOT, '..', '0513-问题点跟踪表', '0512_江布拉克道路质量分析优化报告-问题点跟踪表.xlsx')

# ====== 0. 从 cache.json 预加载工参数据 ======
cell_db = {}
try:
    with open(os.path.join(ROOT, 'cache.json'), 'r', encoding='utf-8') as f:
        cache = json.load(f)
    for rec in cache.get('records', []):
        cn = str(rec.get('小区名', '')).strip()
        if cn:
            cell_db[cn] = {
                '运营商': str(rec.get('运营商', '')).replace('中国', ''),
                '设备商': str(rec.get('设备商', '')),
            }
except Exception:
    pass

# ====== 1. 提取报告问题点 ======
doc = Document(REPORT)
paras = [(p.style.name, p.text.strip()) for p in doc.paragraphs]
ch3_start = None
for i, (style, text) in enumerate(paras):
    if style == '二级标题' and '问题原因分类汇总' in text:
        ch3_start = i; break

problems = []
current = None
for style, text in paras[ch3_start:]:
    if style == '一级标题' and current and '总结' in text: break
    if style == '三级标题':
        if current: problems.append(current)
        current = {'title': text.strip(), 'desc': '', 'analysis': '', 'solutions': []}
        current['_section'] = ''; continue
    if not current: continue
    if text == '问题描述:': current['_section'] = 'desc'; continue
    if text == '问题分析:': current['_section'] = 'analysis'; continue
    if text.startswith('优化方案'): current['_section'] = 'solutions'; continue
    if current['_section'] == 'desc': current['desc'] += text
    elif current['_section'] == 'analysis': current['analysis'] += text
    elif current['_section'] == 'solutions' and text: current['solutions'].append(text)
if current: problems.append(current)

# ====== 辅助函数 ======
def get_cell(text):
    m = re.search(r'小区\s*([^\s，。；、（）时下方]+)', text)
    return m.group(1) if m else ""

def extract_cells(text):
    """从单条方案文本中提取所有小区名"""
    cells = set()
    for m in re.finditer(r'小区\s*([^\s，。；、（）时下方共]+)', text):
        name = m.group(1).strip().rstrip('，。；、（）的下时中方共')
        if len(name) > 3:
            cells.add(name)
    return cells

def lookup(cell):
    if cell in cell_db:
        return cell_db[cell]['运营商'], cell_db[cell]['设备商']
    c = ''
    if cell.startswith('CJ_'): c = '电信'
    for p in ["CJCJS","CJFKS","CJQTX","CJMLX","CJHTB","CJMNS","CJJMS","CJWJQ","CJFCH","CJWCW","CJXHN"]:
        if cell.startswith(p): c = '联通'; break
    return c, ''

def get_issue(title):
    if "无覆盖" in title: return "弱覆盖/无覆盖"
    if "质差" in title: return "质差"
    return "弱覆盖"

def get_measure(solutions):
    text = " ".join(solutions)
    m = set()
    if any(k in text for k in ("新建","建设基站","建立基站")): m.add("新建")
    if "调整" in text and ("下倾角" in text or "方位角" in text): m.add("RF优化")
    if any(k in text for k in ("提申","共享需求")): m.add("共享需求")
    if any(k in text for k in ("核查","告警")): m.add("排查故障")
    if any(k in text for k in ("添加","邻区")): m.add("邻区添加")
    return "/".join(m) if m else "RF优化"

# ====== 2a. 问题点 Sheet — 每条方案匹配小区后汇总运营商/厂家 ======
wb = load_workbook(TRACKER)
ws = wb['问题点']
ws.delete_rows(2, ws.max_row)

for i, p in enumerate(problems):
    measure = get_measure(p['solutions'])
    carriers = set()
    vendors = set()

    for sol in p['solutions']:
        # 新建 → 联通+华为
        if any(k in sol for k in ('新建','建设基站','建立基站','拉远')):
            carriers.add('联通')
            vendors.add('华为')
        # 从方案文本提取小区，查运营商/厂家
        cells = extract_cells(sol)
        if not cells:
            cells.add(get_cell(p['desc'] + p['analysis']))
        for cell in cells:
            c, v = lookup(cell)
            if c: carriers.add(c)
            if v: vendors.add(v)

    if not carriers:
        c, _ = lookup(get_cell(p['desc'] + p['analysis']))
        if c: carriers.add(c)
    if not vendors:
        _, v = lookup(get_cell(p['desc'] + p['analysis']))
        if v: vendors.add(v)

    ws.append([
        i + 1, "江布拉克", p['title'],
        get_issue(p['title']), measure,
        '/'.join(sorted(carriers)), '/'.join(sorted(vendors)),
        "优化", "未闭环", "", "", ""
    ])
print(f"问题点: {len(problems)} 条已写入")

# ====== 2b. RF优化清单 ======
raw_items = []
for prob in problems:
    title = prob['title']
    is_5g = '5G' in title
    for sol in prob['solutions']:
        if '调整' not in sol: continue
        cell_match = re.search(r'调整小区\s*(.+?)(?:下倾角|方位角|下压)', sol)
        if not cell_match: continue
        cell_name = cell_match.group(1).strip()
        tech = '5G' if is_5g else '4G'
        azim_full = re.search(r'方位角\s*(\d+°?\s*->\s*[-]?\d+°?(?:\(\d+\))?)', sol)
        azim = azim_full.group(1) if azim_full else ''
        tilt_full = re.search(r'下倾角\s*(\d+°?\s*->\s*\d+°?)', sol)
        tilt_down = re.search(r'下倾角下压\s*(\d+)°?', sol)
        if tilt_full: tilt = tilt_full.group(1)
        elif tilt_down: tilt = f'下压{tilt_down.group(1)}°'
        else: tilt = ''
        raw_items.append({
            '问题点': title, '小区名': cell_name,
            '调整后方位角': azim, '调整下倾角': tilt, '4G/5G': tech,
        })

merged = {}
for item in raw_items:
    key = (item['问题点'], item['小区名'])
    if key not in merged:
        merged[key] = {**item, '调整后方位角': '', '调整下倾角': ''}
    if item['调整后方位角']: merged[key]['调整后方位角'] = item['调整后方位角']
    if item['调整下倾角']: merged[key]['调整下倾角'] = item['调整下倾角']
rf_items = list(merged.values())

ws_rf = wb['RF优化清单']
ws_rf.delete_rows(2, ws_rf.max_row)
for i, item in enumerate(rf_items):
    row = i + 2; s = "'工参'"
    ws_rf.append([
        item['问题点'], item['小区名'],
        f'=INDEX({s}!A:Z,MATCH(B{row},{s}!F:F,0),MATCH("PCI",{s}!$1:$1,0))',
        f'=INDEX({s}!A:Z,MATCH(B{row},{s}!F:F,0),MATCH("下行频点",{s}!$1:$1,0))',
        f'=INDEX({s}!A:Z,MATCH(B{row},{s}!F:F,0),MATCH("经度",{s}!$1:$1,0))',
        f'=INDEX({s}!A:Z,MATCH(B{row},{s}!F:F,0),MATCH("纬度",{s}!$1:$1,0))',
        item['调整后方位角'], item['调整下倾角'],
        '', item['4G/5G'], '',
    ])
    a = item['调整后方位角'] or '-'; t = item['调整下倾角'] or '-'
    print(f"  [{i+1:2d}] {item['小区名'][:30]} 方位:{a:10s} 倾角:{t:10s} | {item['4G/5G']}")
print(f"RF清单: {len(raw_items)}→{len(rf_items)} 条已写入")

# ====== 2c. 新建清单 ======
new_items = []
for prob in problems:
    title = prob['title']
    is_5g = '5G' in title
    for sol in prob['solutions']:
        if not any(k in sol for k in ('新建','建设基站','建立基站','拉远')): continue
        coord_match = re.search(r'经纬度[：:]\s*(\d+\.?\d*)°?\s*[,，]\s*(\d+\.?\d*)°?', sol)
        if not coord_match: continue
        lng = coord_match.group(1); lat = coord_match.group(2)
        tech = '5G' if is_5g else '4G'
        name_match = re.search(r'(?:建立基站|建设基站)\s*([^\s，。；]+)', sol)
        remark = name_match.group(1) if name_match else ''
        new_items.append({
            '问题点道路': title, '经度': f'{lng}°', '纬度': f'{lat}°',
            '4G/5G': tech, '备注': remark,
        })

seen_first = {}
for item in new_items:
    key = (item['经度'], item['纬度'], item['4G/5G'])
    orig = item.get('备注', '')
    if key not in seen_first:
        seen_first[key] = item['问题点道路']
    else:
        short = seen_first[key].split('：')[-1] if '：' in seen_first[key] else seen_first[key]
        ref = f'同"{short}"'
        item['备注'] = f'{ref}, {orig}' if orig else ref

ws_new = wb['新建']
ws_new.delete_rows(2, ws_new.max_row)
for item in new_items:
    ws_new.append([item['问题点道路'], item['经度'], item['纬度'], item['4G/5G'], item['备注']])

# ====== 2d. 需要共享清单 ======
share_items = []
for prob in problems:
    title = prob['title']
    for sol in prob['solutions']:
        if not any(k in sol for k in ('提申', '共享需求', '共享')): continue
        cells = extract_cells(sol)
        for cell in cells:
            c, v = lookup(cell)
            share_items.append({
                '问题点': title, '小区名称': cell,
                '厂家': v if v else lookup(cell)[1],
            })

ws_share = wb['需要共享清单']
ws_share.delete_rows(2, ws_share.max_row)
for item in share_items:
    ws_share.append([item['问题点'], item['小区名称'], item['厂家']])
    print(f"  共享: {item['小区名称'][:40]} | {item['厂家']}")

print(f"共享清单: {len(share_items)} 条已写入")
wb.save(TRACKER)
print(f"新建清单: {len(new_items)} 条已写入\nDone")
