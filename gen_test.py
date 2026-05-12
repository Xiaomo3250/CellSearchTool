"""生成测试 MapInfo 文件到 test_output 目录"""
import os, sys, io, zipfile
sys.path.insert(0, os.path.dirname(__file__))
from mapinfo_writer import generate_mapinfo_tab

out = os.path.join(os.path.dirname(__file__), 'test_output')
os.makedirs(out, exist_ok=True)

# ===== MIF/MID =====
mif = '''Version 300
Charset "WindowsSimpChinese"
Delimiter ","
Columns 10
  eNBname Char(100)
  eNBid Char(100)
  Longitude Float
  Latitude Float
  Azimuth Integer
  基站类型 Char(100)
  PCI Char(100)
  EARFCN Char(100)
  CellName Char(100)
  CellID Char(100)
Data

Point 88.068375 44.090481
  Symbol (34,0,12)
Point 87.500000 44.200000
  Symbol (34,0,12)
'''
mid = '"测试站A","12345",88.068375,44.090481,60,"宏站","100","1850","测试小区A","1"\n"测试站B","12346",87.500000,44.200000,120,"宏站","101","1850","测试小区B","2"\n'

with open(os.path.join(out, 'test.mif'), 'w', encoding='gbk') as f: f.write(mif)
with open(os.path.join(out, 'test.mid'), 'w', encoding='gbk') as f: f.write(mid)

# ===== TAB 原生格式 =====
recs = [
    {'eNBname':'测试站A','eNBid':'12345','Longitude':'88.068375','Latitude':'44.090481','Azimuth':'60','基站类型':'宏站','PCI':'100','EARFCN':'1850','CellName':'测试小区A','CellID':'1'},
    {'eNBname':'测试站B','eNBid':'12346','Longitude':'87.500000','Latitude':'44.200000','Azimuth':'120','基站类型':'宏站','PCI':'101','EARFCN':'1850','CellName':'测试小区B','CellID':'2'},
]
generate_mapinfo_tab(recs, os.path.join(out, 'test_tab'))

# ===== ZIP 打包 =====
for prefix, label in [('test', 'MIF'), ('test_tab', 'TAB')]:
    exts = ['.mif','.mid'] if label == 'MIF' else ['.tab','.dat','.map','.id']
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for ext in exts:
            fp = os.path.join(out, prefix + ext)
            if os.path.exists(fp):
                zf.write(fp, prefix + ext)
    with open(os.path.join(out, f'{prefix}_{label}.zip'), 'wb') as f:
        f.write(buf.getvalue())

print('生成完成:')
for f in sorted(os.listdir(out)):
    sz = os.path.getsize(os.path.join(out, f))
    print(f'  {f} ({sz} bytes)')
