import sys
sys.path.insert(0, r'D:\XiaoMo\XiaoMo-WorkBuddy\0423-工参管理器')
from app import import_file, records, file_sources, search_records

records.clear()
file_sources.clear()

# Test 1: China Telecom 4G LTE
print("=" * 50)
print("Test 1: Import Telecom LTE")
cnt = import_file(r'D:\XiaoMo\XiaoMo-WorkBuddy\0422\工参\电信_昌吉LTE汇总工参20260421.xlsx')
print(f"LTE import: {cnt}")

dx4g = [r for r in records if r.get('\u6280\u672f\u5236\u5f0f') == '4G LTE']
print(f"4G LTE count: {len(dx4g)}")
if dx4g:
    r = dx4g[0]
    print(f"Sample record:")
    print(f"  cell_name: {r.get('\u5c0f\u533a\u540d')}")
    print(f"  ECI: {r.get('ECI')}")
    print(f"  station: {r.get('\u57fa\u7ad9\u540d')}")
    print(f"  station_id: {r.get('\u57fa\u7ad9ID')}")
    print(f"  freq: {r.get('\u4e0b\u884c\u9891\u70b9')}")
    print(f"  vendor: {r.get('\u8bbe\u5907\u5546')}")

# Test 2: Search
print("\n" + "=" * 50)
print("Test 2: Search")
results, total = search_records('CJ_QT0', 1, 10)
print(f"Search 'CJ_QT0': {total} results")
if results:
    print(f"  First: {results[0].get('\u5c0f\u533a\u540d')}")

results2, total2 = search_records('829312', 1, 10)
print(f"Search '829312': {total2} results")
if results2:
    print(f"  First: {results2[0].get('\u5c0f\u533a\u540d')}")

# Test 3: 5G import
print("\n" + "=" * 50)
print("Test 3: Import Telecom 5G")
cnt5g = import_file(r'D:\XiaoMo\XiaoMo-WorkBuddy\0422\工参\电信_昌吉5G汇总工参20260421.xlsx')
print(f"5G import: {cnt5g}")

dx5g = [r for r in records if r.get('\u6280\u672f\u5236\u5f0f') == '5G NR']
print(f"5G NR count: {len(dx5g)}")

# Test 4: Unicom
print("\n" + "=" * 50)
print("Test 4: Import Unicom LTE")
cnt4 = import_file(r'D:\XiaoMo\XiaoMo-WorkBuddy\0422\工参\联通_新疆联通网优集采LTE工参-昌吉20260413.xlsx')
print(f"Unicom import: {cnt4}")

lt4g = [r for r in records if r.get('\u8fd0\u8425\u5546') == '\u4e2d\u56fd\u8054\u901a']
print(f"Unicom count: {len(lt4g)}")

print("\n" + "=" * 50)
print(f"TOTAL records: {len(records)}")
print(f"Sources: {file_sources}")
