"""问题点输出 + RF优化清单生成 — 从报告提取，写入跟踪表
用法: 放在报告同目录，修改 REPORT/TRACKER 文件名后运行
"""
import os, re
from docx import Document
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(ROOT, '..', '0513-问题点跟踪表', '0512_江布拉克道路质量分析优化报告.docx')
TRACKER = os.path.join(ROOT, '..', '0513-问题点跟踪表', '0512_江布拉克道路质量分析优化报告-问题点跟踪表.xlsx')

# ====== 1. 提取报告问题点及优化方案 ======
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
    m = re.search(r'小区\s*([^(\s\u3002\uff0c\uff1b\uff09\uff08时]+)', text)
    return m.group(1) if m else ""

# 从 cache.json 预加载工参数据（读取工参管理器的导入缓存）
cell_db = {}
try:
    import json
    cache_path = os.path.join(ROOT, 'cache.json')
    with open(cache_path, 'r', encoding='utf-8') as f:
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

def get_carrier(cell):
    if cell in cell_db:
        c = cell_db[cell]['运营商']
        if c: return c
    if cell.startswith("CJ_"): return "电信"
    for p in ["CJCJS","CJFKS","CJQTX","CJMLX","CJHTB","CJMNS","CJJMS","CJWJQ","CJFCH","CJWCW","CJXHN"]:
        if cell.startswith(p): return "联通"
    return ""

def get_vendor(cell):
    if cell in cell_db:
        v = cell_db[cell]['设备商']
        if v: return v
    # 兜底
    for part in cell.split("_"):
        if len(part) >= 4:
            p2 = part[:2].upper()
            if p2 in ("GH","HB"): return "华为"
            if p2 in ("TS","DS"): return "大唐"
            if p2 == "NR": return "诺基亚"
    return "华为"

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

# ====== 2a. 写入问题点 Sheet1（运营商/厂家按措施类型分）======
wb = load_workbook(TRACKER)
ws = wb['问题点']
ws.delete_rows(2, ws.max_row)

for i, p in enumerate(problems):
    cell = get_cell(p['desc'] + p['analysis'])
    measure = get_measure(p['solutions'])
    actual_carrier = get_carrier(cell)
    actual_vendor = get_vendor(cell)
    # 运营商: 新建→联通, 其他→实际小区运营商, 混合→联通/电信
    carriers = set()
    vendors = set()
    if '新建' in measure:
        carriers.add('联通')
        vendors.add('华为')
    if any(k in measure for k in ('RF优化', '共享需求', '排查故障', '邻区添加')):
        if actual_carrier: carriers.add(actual_carrier)
        if actual_vendor: vendors.add(actual_vendor)
    if not carriers: carriers.add(actual_carrier or '')
    if not vendors: vendors.add(actual_vendor or '')
    carrier_str = '/'.join(sorted(carriers))
    vendor_str = '/'.join(sorted(vendors))
    ws.append([
        i + 1, "江布拉克", p['title'],
        get_issue(p['title']), measure,
        carrier_str, vendor_str,
        "优化", "未闭环", "", "", ""
    ])
print(f"问题点: {len(problems)} 条已写入")

# ====== 2b. 写入RF优化清单（INDEX+MATCH公式）======

# 提取RF调整项
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

# 合并同问题+同小区
merged = {}
for item in raw_items:
    key = (item['问题点'], item['小区名'])
    if key not in merged:
        merged[key] = {**item, '调整后方位角': '', '调整下倾角': ''}
    if item['调整后方位角']: merged[key]['调整后方位角'] = item['调整后方位角']
    if item['调整下倾角']: merged[key]['调整下倾角'] = item['调整下倾角']
rf_items = list(merged.values())

# 写入RF优化清单
ws_rf = wb['RF优化清单']
ws_rf.delete_rows(2, ws_rf.max_row)

for i, item in enumerate(rf_items):
    row = i + 2
    s = "'工参'"
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
    print(f"  [{i+1:2d}] {item['小区名'][:30]:30s} 方位:{a:10s} 倾角:{t:10s} | {item['4G/5G']}")

print(f"RF清单: {len(raw_items)} 条原始 → {len(rf_items)} 条合并后 已写入")

# ====== 2c. 写入新建基站清单 ======
new_items = []
for prob in problems:
    title = prob['title']
    is_5g = '5G' in title
    for sol in prob['solutions']:
        if not any(k in sol for k in ('新建', '建设基站', '建立基站', '拉远')):
            continue
        # 提取经纬度: "于经纬度：89.7°,43.5°" 或 "（同X)于经纬度:89.7,43.5"
        coord_match = re.search(r'经纬度[：:]\s*(\d+\.?\d*)°?\s*[,，]\s*(\d+\.?\d*)°?', sol)
        if not coord_match: continue
        lng = coord_match.group(1)
        lat = coord_match.group(2)
        tech = '5G' if is_5g else '4G'
        # 提取基站名（如有）
        name_match = re.search(r'(?:建立基站|建设基站)\s*([^\s，。；]+)', sol)
        remark = name_match.group(1) if name_match else ''
        new_items.append({
            '问题点道路': title, '经度': f'{lng}°', '纬度': f'{lat}°',
            '4G/5G': tech, '备注': remark,
        })

# 全部保留，同坐标+同制式的后续行在备注标注第一个问题点
seen_first = {}
for item in new_items:
    key = (item['经度'], item['纬度'], item['4G/5G'])
    orig_remark = item.get('备注', '')
    if key not in seen_first:
        seen_first[key] = item['问题点道路']
    else:
        first_title = seen_first[key]
        short = first_title.split('：')[-1] if '：' in first_title else first_title
        ref = f'同"{short}"'
        item['备注'] = f'{ref}, {orig_remark}' if orig_remark else ref

ws_new = wb['新建']
ws_new.delete_rows(2, ws_new.max_row)

for i, item in enumerate(new_items):
    ws_new.append([
        item['问题点道路'], item['经度'], item['纬度'],
        item['4G/5G'], item['备注'],
    ])
    print(f"  [{i+1:2d}] ({item['经度']}, {item['纬度']}) {item['4G/5G']} | {item['问题点道路'][:30]}...")

wb.save(TRACKER)
print(f"新建清单: {len(new_items)} 条已写入")
print(f"\nDone — {TRACKER}")
