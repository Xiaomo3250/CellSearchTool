import os, sys
import re
import json
import threading
import socket
import mimetypes
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
import pandas as pd

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(WORKSPACE, "uploaded_files")
CACHE_FILE = os.path.join(WORKSPACE, "cache.json")
STATIC_DIR = os.path.join(WORKSPACE, "static")
os.makedirs(UPLOAD_DIR, exist_ok=True)

__version__ = "0.3.0-alpha"

# ============ 全局数据库 ============
db_lock = threading.Lock()
records = []        # 每条记录含 _文件名 / _工作表 字段
file_sources = {}   # key=filename, val={sheets:[...], count:int, carrier, tech}

# 导入进度状态
import_status = {
    "running": False, "total_files": 0, "current_file": "",
    "current_index": 0, "current_sheet": "", "imported_records": 0,
    "done": False, "error": "",
}


# ============ 字段映射规则 ============
MAPPING_RULES = [
    {
        "carrier": "电信", "tech": "5G",
        "sheet_keywords": ["5G汇总工参", "5G工参", "汇总工参"],
        "exclude_keywords": ["删除", "Sheet1"],
        "field_map": {
            "运营商": "中国电信", "技术制式": "5G NR",
            "设备商": ["厂家"],
            "基站名": ["基站名称", "基站/楼盘名称", "站址"],
            "基站ID": ["gNodeB标识"],
            "小区名": ["NR小区名称"],
            "小区ID": ["小区ID", "小区本地ID"],
            "PCI": ["物理小区标识"],
            "下行频点": ["SSB绝对信道号", "下行频点"],
            "下倾角": ["下倾角", "机械下倾角", "电子下倾角"],
            "挂高": ["挂高", "天线挂高", "站高"],
            "方位角": ["方位角"],
            "经度": ["经度"],
            "纬度": ["纬度", "维度"],
            "频段": ["频带", "频段"],
            "共享": ["是否共享", "共享方"],
        },
    },
    {
        "carrier": "电信", "tech": "4G",
        "sheet_keywords": ["电信工参", "LTE工参", "L1.1室外宏站", "L2.1室内", "室外宏站", "室内"],
        "exclude_keywords": ["删除", "Sheet1"],
        "field_map": {
            "运营商": "中国电信", "技术制式": "4G LTE",
            "设备商": ["设备厂家", "厂家"],
            "基站名": ["EnodebName", "基站名称", "站址"],
            "基站ID": ["EnodebID", "基站ID"],
            "小区名": ["CellName", "小区名称"],
            "小区ID": ["CELLID", "本地CellID", "小区ID"],
            "PCI": ["PCI", "物理小区标识"],
            "下行频点": ["下行中心频点号", "EARFCN", "下行频点", "频点"],
            "下倾角": ["总下倾角", "Downtilt", "下倾角", "机械下倾角"],
            "挂高": ["天线挂高", "GroudHeight", "站高", "挂高"],
            "方位角": ["方位角", "Azimuth"],
            "经度": ["经度", "Longitude"],
            "纬度": ["纬度", "维度", "Latitude"],
            "频段": ["网络类型", "频段", "频带"],
            "共享": ["是否共享", "共享方"],
        },
    },
    {
        "carrier": "联通", "tech": "4G",
        "sheet_keywords": None,
        "exclude_keywords": ["删除", "Sheet1"],
        "field_map": {
            "运营商": "中国联通", "技术制式": "4G LTE",
            "设备商": ["设备厂家", "厂家"],
            "基站名": ["基站名称", "eNodeBName", "站址"],
            "基站ID": ["基站ID", "eNodeBID"],
            "小区名": ["小区名称", "CellName"],
            "小区ID": ["小区ID", "CELLID", "CellID", "本地小区标识"],
            "PCI": ["物理小区标识", "PCI"],
            "下行频点": ["频点", "EARFCN", "下行频点"],
            "下倾角": ["下倾角", "Downtilt", "机械下倾角"],
            "挂高": ["站高", "天线挂高", "GroudHeight", "挂高"],
            "方位角": ["方位角", "Azimuth"],
            "经度": ["经度", "Longitude"],
            "纬度": ["维度", "纬度", "Latitude"],
            "频段": ["频段", "频带"],
        },
    },
    {
        "carrier": "联通", "tech": "5G",
        "sheet_keywords": None,
        "exclude_keywords": ["删除", "Sheet1", "行政区划"],
        "field_map": {
            "运营商": "中国联通", "技术制式": "5G NR",
            "设备商": ["厂家"],
            "基站名": ["基站名称"],
            "基站ID": ["基站标识"],
            "小区名": ["小区名称"],
            "小区ID": ["NR小区标识", "小区标识"],
            "PCI": ["PCI"],
            "下行频点": ["SSB频点", "下行频点"],
            "下倾角": ["电子下倾角", "机械下倾角"],
            "挂高": ["挂高"],
            "方位角": ["方位角"],
            "经度": ["经度"],
            "纬度": ["纬度"],
            "频段": ["频带"],
            "共享": ["是否共享"],
        },
    },
]


def safe_str(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s in ("nan", "NaT", "None", "nat") else s


def detect_carrier(filename):
    if "电信" in filename: return "电信"
    if "联通" in filename: return "联通"
    if "移动" in filename: return "移动"
    return None


def detect_tech(filename):
    if any(k in filename for k in ("5G", "NR", "5g")): return "5G"
    if any(k in filename for k in ("4G", "LTE", "lte")): return "4G"
    return None


def get_matching_rule(filename):
    carrier = detect_carrier(filename)
    tech = detect_tech(filename)
    for rule in MAPPING_RULES:
        if rule["carrier"] == carrier and rule["tech"] == tech:
            return rule
    return None


# ============ 基于命名的运营商检测 ============
UNICOM_PREFIXES = {"CJCJS","CJFKS","CJQTX","CJMLX","CJHTB","CJMNS","CJJMS","CJWJQ","CJFCH","CJWCW","CJXHN"}
COUNTY_CODE_RE = re.compile(r'^[A-Z]{2,4}\d?$')
END_SEGMENT_RE = re.compile(r'^\d{1,2}$|^.+[EC]$')

def detect_carrier_from_names(cell_name, station_name):
    """从小区名/基站名判定运营商。返回 (carrier, share_type) 或 (None, None)

    规则优先级：
    1. 含 (LTGX) → 联通共享站
    2. 2段 + 联通区县前缀 → 联通自建
    3. 5-6段 + CJ开头 + 区县码格式 + 末段合法 → 电信
    """
    names = [n for n in (cell_name, station_name) if n and n.strip()]
    for name in names:
        if "(LTGX)" in name:
            return ("中国联通", "共享站")

    for name in names:
        parts = name.split("_")
        n = len(parts)
        first = parts[0] if n > 0 else ""
        last = parts[-1] if n > 0 else ""

        if n == 2 and any(first.startswith(p) for p in UNICOM_PREFIXES):
            return ("中国联通", "自建")

        if n in (5, 6, 7) and first == "CJ" and COUNTY_CODE_RE.match(parts[1]) and END_SEGMENT_RE.match(last):
            return ("中国电信", "非共享")

    return (None, None)


def get_relevant_sheets(xls, rule):
    """根据规则筛选推荐的 sheet（接收 pd.ExcelFile 对象）"""
    return _filter_sheets(xls.sheet_names, rule)


def get_relevant_sheets_by_names(sheet_names, rule):
    """根据规则筛选推荐的 sheet（接收 sheet 名列表）"""
    return _filter_sheets(sheet_names, rule)


def _filter_sheets(all_names, rule):
    exclude = rule.get("exclude_keywords", [])
    keywords = rule.get("sheet_keywords")
    if keywords is None:
        return [s for s in all_names if not any(kw in s for kw in exclude)]
    matched = []
    for sn in all_names:
        if any(kw in sn for kw in exclude): continue
        if any(kw in sn for kw in keywords):
            matched.append(sn)
    if not matched:
        matched = [s for s in all_names if not any(kw in s for kw in exclude)]
    return matched


def resolve_field(row, field_candidates):
    if isinstance(field_candidates, str): return field_candidates
    for col_name in field_candidates:
        s = safe_str(row.get(col_name))
        if s: return s
    return ""


# ============ 缓存 ============
def save_cache():
    """将当前 records / file_sources 序列化到磁盘"""
    try:
        with db_lock:
            data = {"records": records, "file_sources": file_sources}
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, default=str)
        print(f"[缓存] 已保存 {len(records)} 条记录")
    except Exception as e:
        print(f"[缓存] 保存失败: {e}")


def load_cache():
    """启动时从磁盘加载缓存"""
    global records, file_sources
    if not os.path.exists(CACHE_FILE):
        return
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        with db_lock:
            records = data.get("records", [])
            file_sources = data.get("file_sources", {})
        print(f"[缓存] 已加载 {len(records)} 条记录，{len(file_sources)} 个文件")
    except Exception as e:
        print(f"[缓存] 加载失败: {e}")


# ============ 预览（只解析sheet列表，不导入）============
def preview_file(filepath):
    """返回文件的 sheet 预览信息，不导入数据"""
    filename = os.path.basename(filepath)
    rule = get_matching_rule(filename)
    if rule is None:
        return {"filename": filename, "error": "无法识别运营商/技术制式", "sheets": []}
    try:
        # 用 openpyxl read_only 模式：一次打开拿所有 sheet 名和 max_row，不解析内容（~1.5s vs 原来 30s+）
        if HAS_OPENPYXL and filepath.lower().endswith((".xlsx", ".xlsm")):
            wb = load_workbook(filepath, read_only=True)
            all_sheets = wb.sheetnames
            # 在同一个 wb 上取每个 sheet 的 max_row（read_only 模式下非常快）
            row_counts = {}
            for sn in all_sheets:
                try:
                    mr = wb[sn].max_row
                    # 过滤空 sheet 或格式化填充行（max_row=1048576 是 Excel 空行上限）
                    row_counts[sn] = mr if 1 <= mr < 500000 else 0
                except Exception:
                    row_counts[sn] = 0
            wb.close()
            sheets_info = []
            # 用 sheet 名列表判定推荐 sheet（不需要 pd.ExcelFile）
            recommended = get_relevant_sheets_by_names(all_sheets, rule)
            for sn in all_sheets:
                sheets_info.append({
                    "name": sn,
                    "rows": row_counts[sn],
                    "recommended": sn in recommended,
                })
        else:
            xls = pd.ExcelFile(filepath)
            all_sheets = xls.sheet_names
            recommended = get_relevant_sheets(xls, rule)
            sheets_info = []
            for sn in all_sheets:
                try:
                    row_count_est = len(pd.read_excel(filepath, sheet_name=sn, usecols=[0]))
                except Exception:
                    row_count_est = 0
                sheets_info.append({
                    "name": sn,
                    "rows": row_count_est,
                    "recommended": sn in recommended,
                })
            xls.close()
        return {
            "filename": filename,
            "carrier": rule["carrier"],
            "tech": rule["tech"],
            "sheets": sheets_info,
        }
    except Exception as e:
        return {"filename": filename, "error": str(e), "sheets": []}


# ============ 导入（指定sheet列表）============
def import_file_sheets(filepath, sheet_names, status_update=None):
    """只导入 sheet_names 中指定的 sheet"""
    global records, file_sources
    filename = os.path.basename(filepath)
    rule = get_matching_rule(filename)
    if rule is None:
        return 0

    imported_count = 0
    for sn in sheet_names:
        try:
            if status_update:
                status_update(sheet=sn)
            df = pd.read_excel(filepath, sheet_name=sn)
            if df.empty:
                continue
            df = df.fillna("")
            new_records = []
            for _, row in df.iterrows():
                row_dict = row.to_dict()
                rec = {"_文件名": filename, "_工作表": sn}
                for out_field, candidates in rule["field_map"].items():
                    rec[out_field] = resolve_field(row_dict, candidates)

                # 跳过重复表头行（某些Excel第0行有英文列名做副标题）
                # 判断：小区名或基站名与字段候选名重合，则视为伪表头
                cell_name_val = rec.get("小区名", "")
                lng_val = rec.get("经度", "")
                all_candidates = [c for cands in rule["field_map"].values()
                                  if isinstance(cands, list) for c in cands]
                if cell_name_val in all_candidates or lng_val in all_candidates:
                    continue
                # 跳过经纬度明显无效的行（经度不在合理范围）
                try:
                    lng_f = float(lng_val)
                    if not (60 <= lng_f <= 140):
                        continue
                except (ValueError, TypeError):
                    if lng_val:   # 经度有值但不是数字 → 伪表头
                        continue

                # 基于小区名/基站名判定运营商归属
                carrier_detected, share_detected = detect_carrier_from_names(
                    rec.get("小区名", ""), rec.get("基站名", "")
                )
                if carrier_detected:
                    rec["运营商"] = carrier_detected
                    # 电信工参：以Excel"是否共享"列为准（命名检测只能区分电信/联通，无法判共享）
                    if carrier_detected == "中国电信":
                        share_val = rec.get("共享", "")
                        rec["共享"] = "共享" if share_val in ("是", "共享", "Y", "Yes") else "非共享"
                    else:
                        rec["共享"] = share_detected  # 联通：命名检测结果更权威
                else:
                    # 回退到文件名判定
                    rec["运营商"] = rule.get("field_map", {}).get("运营商", "")
                    # 共享字段：仅电信工参做标准化处理
                    if rule.get("carrier") == "电信":
                        share_val = rec.get("共享", "")
                        rec["共享"] = "非共享" if (share_val == "" or share_val == "未共享") else "共享"
                    else:
                        rec["共享"] = ""

                new_records.append(rec)
                imported_count += 1

            with db_lock:
                records.extend(new_records)
                # 统计该文件多数运营商
                carrier_counts = {}
                for r in new_records:
                    c = r.get("运营商", "")
                    carrier_counts[c] = carrier_counts.get(c, 0) + 1
                majority_carrier = max(carrier_counts, key=carrier_counts.get) if carrier_counts else rule["carrier"]
                if filename not in file_sources:
                    file_sources[filename] = {
                        "carrier": majority_carrier, "tech": rule["tech"],
                        "count": 0, "sheets": []
                    }
                else:
                    file_sources[filename]["carrier"] = majority_carrier
                file_sources[filename]["count"] += len(new_records)
                if sn not in file_sources[filename]["sheets"]:
                    file_sources[filename]["sheets"].append(sn)
                if status_update:
                    import_status["imported_records"] = len(records)
            print(f"  [Sheet] {sn}: {len(new_records)} 条")
        except Exception as e:
            print(f"  [Error] sheet {sn}: {e}")

    return imported_count


# ============ 删除指定文件/sheet ============
def remove_sheets(filename, sheet_names=None):
    """
    sheet_names=None  → 删除该文件所有记录
    sheet_names=[...] → 只删除指定 sheet 的记录
    """
    global records, file_sources
    with db_lock:
        if sheet_names is None:
            records = [r for r in records if r.get("_文件名") != filename]
            file_sources.pop(filename, None)
        else:
            sheet_set = set(sheet_names)
            records = [r for r in records
                       if not (r.get("_文件名") == filename and r.get("_工作表") in sheet_set)]
            if filename in file_sources:
                for sn in sheet_names:
                    if sn in file_sources[filename]["sheets"]:
                        file_sources[filename]["sheets"].remove(sn)
                # 重新计算 count
                file_sources[filename]["count"] = sum(
                    1 for r in records if r.get("_文件名") == filename
                )
                if not file_sources[filename]["sheets"]:
                    file_sources.pop(filename, None)
    save_cache()


# ============ 搜索 ============
def search_records(query, page=1, per_page=50):
    q = str(query).strip()
    if not q:
        total = len(records)
        start = (page - 1) * per_page
        return records[start:start + per_page], total

    results = []
    q_lower = q.lower().replace("-", "")
    q_parts = [p for p in q_lower.split() if p]

    for rec in records:
        cell_name   = rec.get("小区名", "").lower().replace("-", "")
        pci         = rec.get("PCI", "").lower().replace("-", "")
        cell_id     = rec.get("小区ID", "").lower().replace("-", "")
        station_name= rec.get("基站名", "").lower().replace("-", "")

        if (q_lower in cell_name or q_lower in pci or
                q_lower in cell_id or q_lower in station_name):
            results.append(rec)
            continue
        for part in q_parts:
            if (part in cell_name or part in pci or
                    part in cell_id or part in station_name):
                results.append(rec)
                break

    total = len(results)
    start = (page - 1) * per_page
    return results[start:start + per_page], total


# ============ 报告生成辅助函数 ============
def _format_cell_record(rec):
    """将内部记录格式化为报告 API 所需的字段（只提取核心字段，去掉内部标记）"""
    return {
        "技术制式": rec.get("技术制式", ""),
        "运营商": rec.get("运营商", ""),
        "设备商": rec.get("设备商", ""),
        "基站名": rec.get("基站名", ""),
        "基站ID": rec.get("基站ID", ""),
        "小区名": rec.get("小区名", ""),
        "小区ID": rec.get("小区ID", ""),
        "PCI": rec.get("PCI", ""),
        "下行频点": rec.get("下行频点", ""),
        "频段": rec.get("频段", ""),
        "下倾角": rec.get("下倾角", ""),
        "挂高": rec.get("挂高", ""),
        "方位角": rec.get("方位角", ""),
        "经度": rec.get("经度", ""),
        "纬度": rec.get("纬度", ""),
        "共享": rec.get("共享", ""),
    }


# ============ 报告生成：直接数据访问客户端 ============
class DirectCellClient:
    """直接查询 records 列表（无需 HTTP），供 ReportGenerator 使用"""

    def find_cell_by_name(self, cell_name: str) -> dict | None:
        """根据小区名查找工参，支持部分匹配"""
        if not cell_name:
            return None
        target = cell_name.lower().replace("-", "")
        # 先精确匹配
        with db_lock:
            for rec in records:
                if rec.get("小区名", "").lower().replace("-", "") == target:
                    return _format_cell_record(rec)
            # 截取末两段模糊匹配
            parts = cell_name.split("_")
            if len(parts) >= 2:
                short = "_".join(parts[-2:]).lower()
                for rec in records:
                    if short in rec.get("小区名", "").lower():
                        return _format_cell_record(rec)
            # 最后兜底：任意包含
            for rec in records:
                if target in rec.get("小区名", "").lower().replace("-", ""):
                    return _format_cell_record(rec)
        return None

    def get_nearby_cells(self, lng: float, lat: float, radius: float = 5.0) -> list[dict]:
        """按经纬度查找附近基站"""
        results = []
        with db_lock:
            for rec in records:
                try:
                    r_lng = float(rec.get("经度", 0))
                    r_lat = float(rec.get("纬度", 0))
                except (ValueError, TypeError):
                    continue
                dist = ((lng - r_lng) * 111.32 * 0.85) ** 2 + ((lat - r_lat) * 111.32) ** 2
                dist = dist ** 0.5
                if dist <= radius:
                    results.append({
                        **_format_cell_record(rec),
                        "distance_km": round(dist, 2),
                    })
        results.sort(key=lambda x: x["distance_km"])
        return results[:50]


# ============ 静态文件服务 ============
def serve_static(path):
    """读取静态文件内容，返回 (bytes, content_type) 或 (None, None)"""
    safe = os.path.normpath(os.path.join(STATIC_DIR, path.lstrip("/")))
    real = os.path.normpath(os.path.realpath(safe))
    # 安全检查：确保路径在 STATIC_DIR 下
    if not real.startswith(os.path.normpath(os.path.realpath(STATIC_DIR))):
        return None, None
    if not os.path.isfile(real):
        return None, None
    content_type, _ = mimetypes.guess_type(real)
    if content_type is None:
        content_type = "application/octet-stream"
    # 常见文本类型统一用 utf-8
    if content_type.startswith("text/") or content_type in ("application/javascript", "application/json"):
        content_type += "; charset=utf-8"
    with open(real, "rb") as f:
        return f.read(), content_type


# ============ HTTP 处理器 ============
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args): pass

    def do_GET(self):
        url = urlparse(self.path)
        p = url.path

        # 首页：从文件读取 index.html
        if p in ("/", "/index.html"):
            content, ct = serve_static("index.html")
            if content:
                self.send_response(200)
                self.send_header("Content-Type", ct)
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(b"index.html not found")

        # 静态资源：CSS / JS 等
        elif p.startswith("/static/"):
            # 去掉 /static/ 前缀，只传相对路径给 serve_static
            relative_path = p[len("/static/"):]
            content, ct = serve_static(relative_path)
            if content:
                self.send_response(200)
                self.send_header("Content-Type", ct)
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_response(404)
                self.end_headers()

        elif p == "/api/stats":
            with db_lock:
                sources = []
                for fn, v in file_sources.items():
                    sources.append({
                        "filename": fn,
                        "carrier": v.get("carrier", ""),
                        "tech": v.get("tech", ""),
                        "count": v.get("count", 0),
                        "sheets": v.get("sheets", []),
                    })
                # 基站去重：按 (基站名, 运营商, 制式) 三元组合并，同名不同运营商算不同基站
                station_set = set()
                for r in records:
                    name = r.get("基站名", "").strip()
                    carrier = r.get("运营商", "")
                    tech = r.get("技术制式", "")
                    if name:
                        station_set.add((name, carrier, tech))
                self._send_json(200, {
                    "total": len(records),
                    "stations": len(station_set),
                    "files": len(file_sources),
                    "sources": sources
                })

        elif p == "/api/import-status":
            with db_lock:
                st = dict(import_status)
                st["imported_records"] = len(records)
            self._send_json(200, st)

        elif p == "/api/search":
            params = parse_qs(url.query)
            q = params.get("q", [""])[0]
            page = max(1, int(params.get("page", ["1"])[0]))
            per_page = max(1, int(params.get("per_page", ["50"])[0]))
            with db_lock:
                results, total = search_records(q, page, per_page)
            pages = max(1, (total + per_page - 1) // per_page)
            self._send_json(200, {"results": results, "total": total, "page": page, "pages": pages})

        elif p == "/api/export":
            params = parse_qs(url.query)
            q = params.get("q", [""])[0]
            with db_lock:
                all_data, _ = search_records(q, 1, 999999)
            cols = ["技术制式", "运营商", "设备商", "基站名", "基站ID", "小区名", "PCI", "小区ID",
                    "下行频点", "下倾角", "挂高", "方位角", "经度", "纬度", "共享", "频段"]
            lines = ["\uFEFF" + ",".join(cols)]
            for r in all_data:
                lines.append(",".join('"' + str(r.get(c, "")).replace('"', '""') + '"' for c in cols))
            csv_content = "\n".join(lines)
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition",
                             "attachment; filename=export_" + __import__("time").strftime("%Y%m%d_%H%M%S") + ".csv")
            self.end_headers()
            self.wfile.write(csv_content.encode("utf-8"))

        # ===== 报告生成模块 API =====
        elif p == "/api/query-cell":
            # 按小区名模糊查询，返回报告所需的完整工参信息
            params = parse_qs(url.query)
            q = params.get("q", [""])[0]
            exact = params.get("exact", ["0"])[0] == "1"
            max_results = min(int(params.get("limit", ["10"])[0]), 100)

            with db_lock:
                if not q.strip():
                    self._send_json(200, {"results": [], "total": 0})
                    return
                results = []
                q_lower = q.lower().replace("-", "")
                for rec in records:
                    cell_name = rec.get("小区名", "").lower().replace("-", "")
                    station_name = rec.get("基站名", "").lower().replace("-", "")
                    if exact:
                        if q_lower == cell_name:
                            results.append(_format_cell_record(rec))
                    else:
                        if q_lower in cell_name or q_lower in station_name:
                            results.append(_format_cell_record(rec))
                    if len(results) >= max_results:
                        break
            self._send_json(200, {"results": results, "total": len(results)})

        elif p == "/api/cell-by-location":
            # 按经纬度范围查找附近基站（用于新建站选址参考）
            params = parse_qs(url.query)
            try:
                lng = float(params.get("lng", ["0"])[0])
                lat = float(params.get("lat", ["0"])[0])
                radius = float(params.get("radius", ["5"])[0])  # 默认5km
            except (ValueError, TypeError):
                self._send_json(400, {"error": "经纬度参数格式错误"})
                return

            with db_lock:
                nearby = []
                for rec in records:
                    try:
                        r_lng = float(rec.get("经度", 0))
                        r_lat = float(rec.get("纬度", 0))
                    except (ValueError, TypeError):
                        continue
                    # 简化的距离计算（适用于小范围）
                    dist = ((lng - r_lng) * 111.32 * 0.85) ** 2 + ((lat - r_lat) * 111.32) ** 2
                    dist = dist ** 0.5
                    if dist <= radius:
                        nearby.append({**_format_cell_record(rec), "distance_km": round(dist, 2)})
                # 按距离排序
                nearby.sort(key=lambda x: x["distance_km"])
                self._send_json(200, {"results": nearby[:50], "total": len(nearby)})

        # ===== 报告模板 =====
        elif p == "/api/report-template":
            template = {
                "area": "阿乌高速",
                "date": "2026年5月",
                "lte_coverage": "74.04%",
                "lte_avg_rsrp": "-95.82",
                "lte_avg_sinr": "11.23",
                "nr_coverage": "",
                "nr_avg_ss_rsrp": "",
                "nr_avg_ss_sinr": "",
                "test_tools": [
                    {"name": "测试手机 + Assistant平台", "purpose": "路测数据采集与优化分析"},
                    {"name": "工参管理器", "purpose": "基站工参查询"},
                ],
                "problems": [
                    {
                        "type": "覆盖类",
                        "network": "4G",
                        "location": "五家渠入口38km处",
                        "cell_name": "CJ_WJQ0_102团8连D_GHCNN_PRT4E_0",
                        "rsrp": "-118.88",
                        "sinr": "",
                        "distance": "8.2km",
                        "root_cause": "服务小区弱覆盖",
                        "cause_detail": "",
                        "nearby_cell": "",
                        "nearby_cell_rsrp": "",
                        "solutions": [
                            "调整{cell}下倾角3°->0°。",
                            "增加{cell}功率。"
                        ]
                    }
                ]
            }
            self._send_json(200, template)

        else:
            self.send_response(404); self.end_headers()

    def do_POST(self):
        url = urlparse(self.path)
        p = url.path

        if p == "/api/preview":
            self._handle_preview()
        elif p == "/api/import-sheets-async":
            self._handle_import_sheets_async()
        elif p == "/api/remove-sheet":
            self._handle_remove_sheet()
        elif p == "/api/clear":
            with db_lock:
                global records, file_sources
                records = []; file_sources = {}
            # 删除缓存文件
            try:
                if os.path.exists(CACHE_FILE): os.remove(CACHE_FILE)
            except Exception: pass
            self._send_json(200, {"ok": True})
        elif p == "/api/generate-report":
            self._handle_generate_report()
        else:
            self.send_response(404); self.end_headers()

    # ---- 预览：解析 sheet 列表，不导入 ----
    def _handle_preview(self):
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        if not content_length:
            self._send_json(400, {"error": "上传文件为空"}); return

        body = self.rfile.read(content_length)
        boundary = None
        for part in content_type.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary = part.split("=", 1)[1].strip('"'); break

        if not boundary:
            self._send_json(400, {"error": "无法解析上传数据"}); return

        files_data = self._parse_multipart(body, boundary.encode())
        results = []
        for fi in files_data:
            if not fi["filename"]: continue
            save_path = os.path.join(UPLOAD_DIR, fi["filename"])
            try:
                with open(save_path, "wb") as f:
                    f.write(fi["data"])
                info = preview_file(save_path)
                results.append(info)
            except Exception as e:
                results.append({"filename": fi["filename"], "error": str(e), "sheets": []})

        self._send_json(200, {"files": results})

    # ---- 异步导入指定 sheet ----
    def _handle_import_sheets_async(self):
        body = self._read_json_body()
        if body is None: return
        files = body.get("files", [])  # [{filename, sheets:[...]}, ...]
        if not files:
            self._send_json(400, {"error": "无文件信息"}); return

        import_status.update({
            "running": True, "total_files": len(files),
            "current_index": 0, "current_file": "",
            "current_sheet": "", "done": False, "error": "",
            "imported_records": len(records),
        })

        def do_import():
            for i, sel in enumerate(files):
                filename = sel.get("filename", "")
                sheets = sel.get("sheets", [])
                import_status["current_file"] = filename
                import_status["current_index"] = i + 1
                import_status["current_sheet"] = ""
                fp = os.path.join(UPLOAD_DIR, filename)
                if not os.path.exists(fp):
                    import_status["error"] += f"{filename}: 文件不存在; "
                    continue
                try:
                    def _su():
                        def status_update(sheet=None):
                            if sheet: import_status["current_sheet"] = sheet
                        return status_update
                    cnt = import_file_sheets(fp, sheets, status_update=_su())
                    print(f"[导入] {filename}: {cnt} 条")
                except Exception as e:
                    import_status["error"] += f"{filename}: {e}; "

            import_status.update({"running": False, "done": True, "current_sheet": ""})
            save_cache()

        threading.Thread(target=do_import, daemon=True).start()
        self._send_json(200, {"pending": True})

    # ---- 移除文件 / sheet ----
    def _handle_remove_sheet(self):
        body = self._read_json_body()
        if body is None: return
        filename = body.get("filename", "")
        sheets = body.get("sheets", None)  # None=删除整个文件
        if not filename:
            self._send_json(400, {"error": "缺少 filename"}); return
        remove_sheets(filename, sheets)
        self._send_json(200, {"ok": True})

    # ---- 报告生成 ----
    def _handle_generate_report(self):
        """接收问题清单 JSON，生成 docx 报告并返回下载"""
        body = self._read_json_body()
        if body is None:
            return
        problem_data = body.get("problem_data")
        if not problem_data:
            self._send_json(400, {"error": "缺少 problem_data"})
            return

        try:
            # 导入报告生成器
            from report_generator import ReportGenerator

            client = DirectCellClient()
            gen = ReportGenerator(client)
            # 生成到临时文件
            area = problem_data.get("area", "report")
            tmp_path = os.path.join(
                tempfile.gettempdir(),
                f"工参报告_{area}_{int(__import__('time').time())}.docx"
            )
            gen.generate(problem_data, tmp_path)

            # 读取文件并返回下载
            with open(tmp_path, "rb") as f:
                docx_data = f.read()
            # 清理临时文件
            try:
                os.remove(tmp_path)
            except Exception:
                pass

            self.send_response(200)
            safe_name = area.replace("/", "_").replace("\\", "_")
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition",
                             f"attachment; filename={safe_name}_优化报告.docx")
            self.send_header("Content-Length", str(len(docx_data)))
            self.end_headers()
            self.wfile.write(docx_data)
        except ImportError as e:
            self._send_json(500, {"error": f"报告生成模块缺失: {e}。请安装: pip install python-docx"})
        except Exception as e:
            import traceback
            traceback.print_exc()
            self._send_json(500, {"error": f"报告生成失败: {e}"})

    # ---- 工具方法 ----
    def _read_json_body(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length)
            return json.loads(raw.decode("utf-8"))
        except Exception as e:
            self._send_json(400, {"error": f"请求体解析失败: {e}"}); return None

    def _parse_multipart(self, body, boundary):
        files = []
        for part in body.split(boundary):
            if len(part) < 10 or b'filename="' not in part: continue
            try:
                header_end = part.find(b"\r\n\r\n")
                if header_end < 0: continue
                header = part[:header_end].decode("utf-8", errors="replace")
                data = part[header_end + 4:]
                if data.endswith(b"\r\n"): data = data[:-2]
                m = re.search(r'filename="([^"]+)"', header)
                if not m or not m.group(1): continue
                files.append({"filename": m.group(1), "data": data})
            except Exception: continue
        return files

    def _send_json(self, code, data):
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False, default=str).encode("utf-8"))


def open_browser(port):
    import subprocess
    subprocess.Popen(['cmd', '/c', 'start', '', f'http://127.0.0.1:{port}'],
                     creationflags=subprocess.CREATE_NO_WINDOW)


if __name__ == "__main__":
    PORT = 18888

    # 预先绑定端口
    test_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    test_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    test_sock.bind(("127.0.0.1", PORT))
    test_sock.close()

    # 启动时加载缓存
    load_cache()

    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()

    print(f"\n{'='*50}")
    print(f"  基站工参管理器已启动")
    print(f"  访问地址: http://127.0.0.1:{PORT}")
    if records:
        print(f"  已从缓存加载 {len(records)} 条记录")
    print(f"{'='*50}\n")

    threading.Timer(0.3, open_browser, args=(PORT,)).start()

    try:
        while True:
            import time; time.sleep(3600)
    except KeyboardInterrupt:
        server.shutdown()
