"""
纯 Python 生成 MapInfo 原生 TAB 格式 (.tab/.dat/.map/.id)
参考: MITAB 开源库格式规范
"""
import struct
import io
import zipfile
import os

def generate_mapinfo_tab(records, output_prefix):
    """
    生成 MapInfo 原生 4 文件
    
    records: list of dict, 每条记录需含以下字段:
        eNBname, eNBid, Longitude, Latitude, Azimuth, 基站类型, PCI, EARFCN, CellName, CellID
    """
    
    # 字段定义 (顺序必须与 .tab 一致)
    fields = [
        ("eNBname", "C", 100),   # Char(100)
        ("eNBid", "C", 100),
        ("Longitude", "F", 0),    # Float
        ("Latitude", "F", 0),
        ("Azimuth", "N", 11),     # Integer (N=numeric)
        ("基站类型", "C", 100),
        ("PCI", "C", 100),
        ("EARFCN", "C", 100),
        ("CellName", "C", 100),
        ("CellID", "C", 100),
    ]
    
    # ====== 1. 生成 .tab ======
    tab_lines = [
        "!table",
        "!version 300",
        "!charset Neutral",
        "",
        "Definition Table",
        "  Type NATIVE Charset \"Neutral\"",
        f"  Fields {len(fields)}",
    ]
    for name, ftype, width in fields:
        if ftype == "F":
            tab_lines.append(f"    {name} Float ;")
        elif ftype == "N":
            tab_lines.append(f"    {name} Integer ;")
        else:
            tab_lines.append(f"    {name} Char ({width}) ;")
    tab_lines.append("")
    tab_content = "\r\n".join(tab_lines)
    
    # ====== 2. 生成 .dat (原生格式) ======
    # 计算每个字段在记录中的宽度
    col_widths = []
    for name, ftype, width in fields:
        if ftype == "F":
            col_widths.append(8)   # double
        elif ftype == "N":
            col_widths.append(4)   # int32
        else:
            col_widths.append(width)  # char width
    
    record_size = sum(col_widths)
    
    # .dat header: type(1) + header_size(2) + field_defs...
    # field_def: name\0 + type(1) + width(1) + decimals(1)
    # type codes: 0x43='C'(char), 0x46='F'(float), 0x4E='N'(integer)
    type_codes = {'C': 0x43, 'F': 0x46, 'N': 0x4E}
    
    # Build header
    header_parts = []
    for name, ftype, width in fields:
        name_bytes = name.encode('gbk') + b'\x00'
        header_parts.append(name_bytes)
        if ftype == "F":
            header_parts.append(bytes([0x46, 8, 15]))  # Float: width=8, decimals=15
        elif ftype == "N":
            header_parts.append(bytes([0x4E, 4, 0]))   # Integer: width=4, decimals=0
        else:
            header_parts.append(bytes([0x43, width, 0]))  # Char
    
    header_body = b''.join(header_parts)
    hdr_size = 3 + len(header_body)  # 1(type) + 2(hdr_size) + field_defs
    
    # 构建记录数据
    record_data = bytearray()
    for rec in records:
        buf = bytearray(record_size)
        pos = 0
        for fi, (name, ftype, width) in enumerate(fields):
            val = rec.get(name, "")
            if ftype == "F":
                try:
                    struct.pack_into('<d', buf, pos, float(val or 0))
                except (ValueError, TypeError):
                    struct.pack_into('<d', buf, pos, 0.0)
                pos += 8
            elif ftype == "N":
                try:
                    struct.pack_into('<i', buf, pos, int(float(val or 0)))
                except (ValueError, TypeError):
                    struct.pack_into('<i', buf, pos, 0)
                pos += 4
            else:
                # Char field: encode to GBK, pad with spaces
                val_str = str(val or "")[:width]
                val_bytes = val_str.encode('gbk', errors='replace')[:width]
                buf[pos:pos+len(val_bytes)] = val_bytes
                pos += width
        record_data.extend(buf)
    
    dat_content = bytes([0x03]) + struct.pack('<H', hdr_size) + header_body + record_data
    
    # ====== 3. 生成 .map ======
    map_parts = []
    for rec in records:
        try:
            lng = float(rec.get("Longitude", 0))
            lat = float(rec.get("Latitude", 0))
        except (ValueError, TypeError):
            lng, lat = 0.0, 0.0
        
        # MapInfo point region: type(1) + coords
        # For a simple point with no extra data: type=0x01
        point_data = (
            struct.pack('<B', 0x01) +           # point type
            struct.pack('<i', 0) +               # num_points (0 = uses next fields)
            struct.pack('<i', 1) +               # num_segments
            struct.pack('<i', 1) +               # num_points_in_segment
            struct.pack('<d', lng) +             # x
            struct.pack('<d', lat)               # y
        )
        # Pad to 8-byte boundary (MITAB requirement)
        pad = (8 - len(point_data) % 8) % 8
        map_parts.append(point_data + b'\x00' * pad)
    
    map_content = b''.join(map_parts)
    
    # ====== 4. 生成 .id ======
    # .id: rec_count(4) + for each rec: map_offset(4) + dat_offset(4)
    rec_count = len(records)
    id_parts = [struct.pack('<I', rec_count)]
    
    # Calculate per-record sizes in .map and .dat
    map_record_size = 4 + 4 + 4 + 4 + 8 + 8  # 32 bytes per point
    map_record_size = ((map_record_size + 7) // 8) * 8  # align to 8
    
    for i in range(rec_count):
        map_off = i * map_record_size
        dat_off = i * record_size
        id_parts.append(struct.pack('<II', map_off, dat_off))
    
    id_content = b''.join(id_parts)
    
    # ====== 5. 写入文件 ======
    prefix = output_prefix
    with open(prefix + '.tab', 'w', encoding='gbk') as f:
        f.write(tab_content)
    with open(prefix + '.dat', 'wb') as f:
        f.write(dat_content)
    with open(prefix + '.map', 'wb') as f:
        f.write(map_content)
    with open(prefix + '.id', 'wb') as f:
        f.write(id_content)
    
    return [prefix + ext for ext in ['.tab', '.dat', '.map', '.id']]


# 测试生成
if __name__ == '__main__':
    test_recs = [
        {"eNBname":"测试站A","eNBid":"12345","Longitude":"88.068375","Latitude":"44.090481",
         "Azimuth":"60","基站类型":"宏站","PCI":"100","EARFCN":"1850","CellName":"测试小区A","CellID":"1"},
        {"eNBname":"测试站B","eNBid":"12346","Longitude":"87.500000","Latitude":"44.200000",
         "Azimuth":"120","基站类型":"宏站","PCI":"101","EARFCN":"1850","CellName":"测试小区B","CellID":"2"},
    ]
    files = generate_mapinfo_tab(test_recs, r'C:\Users\12931\Desktop\test_tab')
    for f in files:
        print(f"生成: {f} ({os.path.getsize(f)} bytes)")
