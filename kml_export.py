"""
KML 导出模块 — 生成标准基站扇区图层（兼容奥维地图/Google Earth）
按 5G / 4G 标准模板格式输出
"""
import json
import math

# ============ 扇形参数 ============
BEAM_WIDTH = 30       # 波瓣宽度（度）
RADIUS_M = 80         # 覆盖半径（米）
ARC_POINTS = 32       # 弧顶点数（总顶点=ARC_POINTS+1）

# ============ MOD3 颜色 ============
# KML 色值格式: AABBGGRR（AA=透明度, BB=蓝, GG=绿, RR=红）
MOD_COLORS = {
    0: {"poly": "E5ff0000", "line": "FFffff00"},  # 红
    1: {"poly": "E500ff00", "line": "FFffff00"},  # 绿
    2: {"poly": "E50000ff", "line": "FFffff00"},  # 蓝
}


def _calc_sector_polygon(lng: float, lat: float, azimuth: float) -> str:
    """
    计算扇形多边形坐标字符串
    参数：
        lng, lat: 中心点坐标（WGS84）
        azimuth: 方位角（0~360）
    返回：KML coordinates 字符串（空格分隔的 lng,lat,5）
    """
    az_rad = math.radians(azimuth)
    half_bw = math.radians(BEAM_WIDTH / 2.0)
    lat_rad = math.radians(lat)

    points = []
    # 起始点：中心点
    points.append(f"{lng:.6f},{lat:.6f},5")

    # 弧顶点：从 azimuth - half_bw 到 azimuth + half_bw
    for i in range(ARC_POINTS):
        angle = az_rad - half_bw + (2.0 * half_bw * i / (ARC_POINTS - 1))
        # 计算偏移（近似球面）
        d_lat = (RADIUS_M * math.cos(angle)) / 111320.0
        d_lng = (RADIUS_M * math.sin(angle)) / (111320.0 * math.cos(lat_rad))
        pt_lng = lng + d_lng
        pt_lat = lat + d_lat
        points.append(f"{pt_lng:.6f},{pt_lat:.6f},5")

    # 闭合点：回到中心点
    points.append(f"{lng:.6f},{lat:.6f},5")

    return " ".join(points)


def _get_mod3(pci_val) -> int:
    """从 PCI 计算 MOD3"""
    try:
        return int(float(pci_val)) % 3
    except (ValueError, TypeError):
        return 0


# ============ 4G LTE description 字段模板 ============
LTE_DESC_FIELDS = [
    ("ECI", ""),
    ("地州", "地州"),
    ("市/县", "市/县"),
    ("维护网格", "维护网格"),
    ("覆盖场景", "覆盖场景"),
    ("基站类型", "站型"),
    ("场景分类", "场景分类"),
    ("场景细化", "场景细化"),
    ("基站名称", "基站名"),
    ("基站ID", "基站ID"),
    ("小区名称", "小区名"),
    ("小区ID", "小区ID"),
    ("本地小区标识", "本地小区标识"),
    ("小区标识", "小区标识"),
    ("物理小区标识", "PCI"),
    ("根序列索引", "根序列索引"),
    ("跟踪区码", "TAC"),
    ("经度", "经度"),
    ("纬度", "纬度"),
    ("方位角", "方位角"),
    ("站高", "挂高"),
    ("频点", "下行频点"),
    ("频段", "频段"),
    ("带宽", "带宽"),
    ("机械下倾角", "机械下倾角"),
    ("电子下倾角", "电子下倾角"),
    ("下倾角", "下倾角"),
    ("站点配置", "站点配置"),
    ("基站入网时间", "入网时间"),
    ("工程期数", "工程期数"),
    ("站址", "站址"),
    ("塔架类型", "塔桅类型"),
    ("天线类型", "天线类型"),
    ("RRU型号", "RRU型号"),
    ("RRU数量", "RRU数量"),
    ("天线厂家", "天线厂家"),
    ("5G小区名称", "5G小区名称"),
    ("业务IP地址", "业务IP地址"),
    ("信令面IP地址", "信令面IP地址"),
    ("是否共享", "共享"),
    ("乡、镇（团场）、街道", "乡镇街道"),
    ("行政村（连队）、居委会（社区）", "行政村"),
    ("厂家", "设备商"),
]


# ============ 5G NR description 字段模板 ============
NR_DESC_FIELDS = [
    ("厂家", "设备商"),
    ("地州", "地州"),
    ("区县", "区县"),
    ("gNodeB标识_小区ID", ""),
    ("gNodeB标识", "基站ID"),
    ("基站名称", "基站名"),
    ("NR小区名称", "小区名"),
    ("经度", "经度"),
    ("纬度", "纬度"),
    ("小区本地ID", "小区本地ID"),
    ("小区ID", "小区ID"),
    ("物理小区标识", "PCI"),
    ("跟踪区域码", "TAC"),
    ("频带", "频段"),
    ("方位角", "方位角"),
    ("挂高", "挂高"),
    ("上行频点", "上行频点"),
    ("下行频点", "下行频点"),
    ("SSB绝对信道号", "SSB绝对信道号"),
    ("SSB频域位置", "SSB频域位置"),
    ("上行带宽", "上行带宽"),
    ("下行带宽", "下行带宽"),
    ("子载波间隔(KHz)", "子载波间隔"),
    ("根序列", "根序列"),
    ("发送和接收模式", "收发模式"),
    ("入网日期", "入网日期"),
    ("是否共享", "共享"),
    ("站型", "站型"),
    ("站址编码", "站址编码"),
    ("站址", "站址"),
    ("塔桅类型", "塔桅类型"),
    ("覆盖场景", "覆盖场景"),
    ("五高一地一景", "五高一地一景"),
    ("详细覆盖场景", "详细覆盖场景"),
    ("BBU SN码", "BBU SN码"),
    ("AAU SN码", "AAU SN码"),
    ("AAU型号", "AAU型号"),
    ("共站4GCI", "共站4GCI"),
    ("共5G站的4G小区名称\n（优先级1.8G>2.1G>800）", "共站5G小区名"),
    ("基站/楼盘名称", "基站/楼盘名称"),
]


def _build_description(rec: dict, fields_template: list) -> str:
    """
    根据字段模板构建 description HTML 表格
    fields_template: [(显示名, 标准化列名或"")]
    空字符串表示仅用于占位（显示名存在但值留空）
    优先从标准化列取值，否则从 _raw JSON 中匹配显示名
    """
    # 尝试解析 _raw
    raw_dict = {}
    raw_str = rec.get("_raw", "{}")
    if isinstance(raw_str, str):
        try:
            raw_dict = json.loads(raw_str)
        except (json.JSONDecodeError, TypeError):
            raw_dict = {}
    elif isinstance(raw_str, dict):
        raw_dict = raw_str

    rows = []
    for display_name, std_col in fields_template:
        # 取值优先级：标准化列 > _raw 中的显示名 > _raw 中模糊匹配 > "-"
        val = ""
        if std_col:
            val = str(rec.get(std_col, "") or "")
        if not val:
            val = str(raw_dict.get(display_name, "") or "")
        if not val:
            # 模糊匹配：去掉换行和括号后匹配
            clean_display = display_name.split("\n")[0].strip()
            for rk, rv in raw_dict.items():
                if clean_display in str(rk):
                    val = str(rv)
                    break
        if not val:
            val = "-"

        rows.append(
            f'<tr><th scope="col">{display_name}</th>'
            f'<th scope="col">{_xml_escape(val)}</th></tr>'
        )

    table = f"<table width='300' border='1'>{''.join(rows)}</table>"
    return table


def _xml_escape(s: str) -> str:
    """XML/HTML 转义"""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def generate_kml(records: list, tech_type: str = "4G") -> str:
    """
    生成标准 KML 基站扇区图层
    tech_type: "4G" 或 "5G"
    """
    is_5g = (tech_type == "5G")
    fields_template = NR_DESC_FIELDS if is_5g else LTE_DESC_FIELDS
    doc_name = f"{'5G' if is_5g else '4G'}工参-基站扇区"

    # KML 头部（含 Style 定义）
    header = f"""<?xml version="1.0" encoding="utf-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>{doc_name}</name>
<Style id="mod0">
  <LineStyle><width>1</width><color>{MOD_COLORS[0]['line']}</color></LineStyle>
  <PolyStyle><color>{MOD_COLORS[0]['poly']}</color></PolyStyle>
</Style>
<Style id="mod1">
  <LineStyle><width>1</width><color>{MOD_COLORS[1]['line']}</color></LineStyle>
  <PolyStyle><color>{MOD_COLORS[1]['poly']}</color></PolyStyle>
</Style>
<Style id="mod2">
  <LineStyle><width>1</width><color>{MOD_COLORS[2]['line']}</color></LineStyle>
  <PolyStyle><color>{MOD_COLORS[2]['poly']}</color></PolyStyle>
</Style>
<Folder>
<name>{doc_name}</name>
"""

    # 生成每个小区的 Placemark
    placemarks = []
    skipped = 0
    for rec in records:
        # 获取坐标
        try:
            lng = float(rec.get("经度", 0) or 0)
            lat = float(rec.get("纬度", 0) or 0)
        except (ValueError, TypeError):
            skipped += 1
            continue
        if not (60 <= lng <= 140 and 15 <= lat <= 55):
            skipped += 1
            continue

        # 获取方位角
        try:
            azimuth = float(rec.get("方位角", 0) or 0)
        except (ValueError, TypeError):
            azimuth = 0

        # 小区名
        cell_name = str(rec.get("小区名", "") or "")

        # PCI MOD3 颜色
        pci_val = rec.get("PCI", "0")
        mod3 = _get_mod3(pci_val)

        # 构建 description
        description = _build_description(rec, fields_template)

        # 扇形多边形
        coords = _calc_sector_polygon(lng, lat, azimuth)

        placemark = f"""<Placemark>
  <name>{_xml_escape(cell_name)}</name>
  <description><![CDATA[{description}]]></description>
  <styleUrl>#mod{mod3}</styleUrl>
  <Polygon>
    <extrude>1</extrude>
    <altitudeMode>relativeToGround</altitudeMode>
    <outerBoundaryIs>
      <LinearRing>
        <coordinates>{coords}</coordinates>
      </LinearRing>
    </outerBoundaryIs>
  </Polygon>
</Placemark>
"""
        placemarks.append(placemark)

    # KML 尾部
    footer = """</Folder>
</Document>
</kml>"""

    kml_content = header + "".join(placemarks) + footer
    print(f"[KML] 生成 {len(placemarks)} 个扇区，跳过 {skipped} 个（无效坐标）")
    return kml_content
