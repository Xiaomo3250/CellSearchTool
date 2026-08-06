import 'dart:convert';
import 'dart:isolate';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
import '../models/station.dart';
import 'database.dart';

/// 工参导入器 — 解析 Excel 并标准化字段后写入 SQLite
class Importer {
  final AppDatabase _db = AppDatabase();

  // 字段映射（中文值都是字符串 key，Dart 安全）
  static const _fieldMap = <String, List<String>>{
    '设备商':   ['厂家', '设备厂家'],
    '基站名':   ['基站名称', '网元名称', 'EnodebName', 'eNodeBName', '基站/楼盘名称', '站址'],
    '基站ID':   ['gNodeB标识', 'gNodeB ID', 'EnodebID', 'eNodeBID', '基站ID', '基站标识'],
    '小区名':   ['NR小区名称', '小区名称', 'CellName'],
    '小区ID':   ['小区ID', '小区本地ID', 'CELLID', '本地CellID', 'NR小区标识', '小区标识', 'CellID', '本地小区标识'],
    'PCI':      ['物理小区标识', 'PCI'],
    '下行频点': ['SSB频点', 'SSB绝对信道号', '下行中心频点号', 'EARFCN', '频点', '下行频点'],
    '下倾角':   ['总下倾角', 'Downtilt', '下倾角', '机械下倾角', '电子下倾角'],
    '挂高':     ['挂高', '天线挂高', '站高', 'GroudHeight'],
    '方位角':   ['方位角', 'Azimuth'],
    '经度':     ['经度', 'Longitude'],
    '纬度':     ['纬度', '维度', 'Latitude'],
    '频段':     ['频带', '频段', '网络类型'],
    '共享':     ['是否共享', '共享方'],
    'TAC':      ['TAC', '跟踪区码', '跟踪区域码', 'TAL'],
  };

  /// 快速获取工作表名列表（只读 workbook.xml，不解析数据）
  static List<String> listSheetNames(List<int> bytes) {
    try {
      final archive = ZipDecoder().decodeBytes(bytes);
      final wb = archive.findFile('xl/workbook.xml');
      if (wb == null) return [];
      wb.decompress();
      final doc = XmlDocument.parse(utf8.decode(wb.content));
      return doc.findAllElements('sheet')
          .map((e) => e.getAttribute('name') ?? '')
          .where((n) => n.isNotEmpty)
          .toList();
    } catch (_) {
      return [];
    }
  }

  /// 返回 {ok, total, skipEmpty, skipLng, sheets, cols, error}
  /// [sheetIndices] 指定要导入的工作表索引（null=全部）
  Future<Map<String, dynamic>> importFromBytes(List<int> bytes, String filename, {Set<int>? sheetIndices}) async {
    final result = <String, dynamic>{
      'ok': 0, 'total': 0, 'skipEmpty': 0, 'skipLng': 0,
      'sheets': <String>[], 'cols': <String>[], 'error': '',
    };

    try {
      return await _doImport(bytes, filename, result, sheetIndices: sheetIndices);
    } catch (e, st) {
      // FATAL 前先做裸 ZIP 诊断，确保不丢失
      String preDiag;
      try { preDiag = _diagnoseRaw(bytes); } catch (_) { preDiag = 'ZIPerr'; }
      result['error'] = '$preDiag || FATAL: $e\n$st';
      return result;
    }
  }

  Future<Map<String, dynamic>> _doImport(List<int> bytes, String filename, Map<String, dynamic> result, {Set<int>? sheetIndices}) async {
    // 后台 isolate 解析 XML（大数据不卡 UI）
    final filter = sheetIndices;
    final parsed = await Isolate.run(() => _parseXlsxRaw(bytes, sheetIndices: filter));
    if (parsed == null) {
      result['error'] = '无法解析 xlsx 文件';
      return result;
    }
    final sheetNames = parsed.sheetNames;
    final allSheets = parsed.sheets; // List<List<Map<String, String>>>

    if (allSheets.isEmpty) {
      result['error'] = '未找到工作表';
      return result;
    }

    final tech = _detectTech(filename);
    final stations = <Station>[];
    int totalRows = 0, skipEmpty = 0, skipLng = 0;
    final sheetList = <String>[];
    List<String> firstCols = [];

    for (var si = 0; si < allSheets.length; si++) {
      final allRows = allSheets[si];
      final sName = si < sheetNames.length ? sheetNames[si] : 'Sheet${si + 1}';

      if (allRows.length < 2) continue;

      // 表头行：列字母 → 列名
      final headerRow = allRows.first;
      final colLetterToName = <String, String>{};
      for (final e in headerRow.entries) {
        final name = e.value.trim();
        if (name.isNotEmpty) colLetterToName[e.key] = name;
      }
      if (colLetterToName.isEmpty) continue;
      if (firstCols.isEmpty) firstCols = colLetterToName.values.take(12).toList();

      sheetList.add('$sName(${allRows.length - 1}row)');

      // 逐行解析
      for (var i = 1; i < allRows.length; i++) {
        totalRows++;
        final rowData = allRows[i];

        final dict = <String, String>{};
        var hasVal = false;
        for (final e in rowData.entries) {
          final colName = colLetterToName[e.key];
          if (colName != null) {
            final v = e.value.trim();
            if (v.isNotEmpty) hasVal = true;
            dict[colName] = v;
          }
        }
        if (!hasVal) { skipEmpty++; continue; }

        // 运营商检测（多级回退，与桌面端逻辑一致）
        final ctor = dict['承建方'] ?? '';
        var carrier = _detectCarrier(ctor, filename, dict);

        // 字段映射
        final vals = <String, String>{};
        for (final e in _fieldMap.entries) {
          vals[e.key] = _pick(dict, e.value);
        }
        // 设备商单独映射（不在 17 个标准字段中，但卡片需要显示）
        vals['设备商'] = _pick(dict, ['厂家', '设备厂家', '设备商']);

        // 小区ID 规范化：联通4G 长格式(基站ID+小区标识) → 截取短标识
        final rawSiteId = vals['基站ID'] ?? '';
        final rawCellId = vals['小区ID'] ?? '';
        final shortCellId = (rawSiteId.isNotEmpty && rawCellId.startsWith(rawSiteId))
            ? rawCellId.substring(rawSiteId.length)
            : rawCellId;
        vals['小区ID'] = shortCellId;

        // 经纬度校验
        final lngStr = vals['经度'] ?? '';
        double? lng;
        if (lngStr.isNotEmpty) lng = double.tryParse(lngStr);
        if (lng != null && (lng < 50 || lng > 150)) { skipLng++; continue; }

        stations.add(Station(
          tech: tech, carrier: carrier,
          vendor: vals['设备商'] ?? '',
          siteName: vals['基站名'] ?? '',
          siteId: vals['基站ID'] ?? '',
          cellName: vals['小区名'] ?? '',
          cellId: vals['小区ID'] ?? '',
          pci: vals['PCI'] ?? '',
          dlFreq: vals['下行频点'] ?? '',
          tilt: vals['下倾角'] ?? '',
          height: vals['挂高'] ?? '',
          azimuth: vals['方位角'] ?? '',
          lng: vals['经度'] ?? '',
          lat: vals['纬度'] ?? '',
          share: vals['共享'] ?? '',
          band: vals['频段'] ?? '',
          tac: vals['TAC'] ?? '',
          filename: filename, sheet: sName,
          raw: jsonEncode(dict),
        ));
      }
    }

    if (stations.isNotEmpty) {
      await _db.deleteByFilename(filename);
      await _db.insertRecords(stations);
    }

    result['ok'] = stations.length;
    result['total'] = totalRows;
    result['skipEmpty'] = skipEmpty;
    result['skipLng'] = skipLng;
    result['sheets'] = sheetList;
    result['cols'] = firstCols;
    return result;
  }

  /// 纯 XML 解析 xlsx → 返回工作表名和行数据（静态方法，可在 isolate 中运行）
  static ({List<String> sheetNames, List<List<Map<String, String>>> sheets})? _parseXlsxRaw(List<int> bytes, {Set<int>? sheetIndices}) {
    final Archive archive;
    try {
      archive = ZipDecoder().decodeBytes(bytes);
    } catch (_) {
      return null;
    }

    // 1. 读 shared strings
    final sharedStrings = <String>[];
    final ss = archive.findFile('xl/sharedStrings.xml');
    if (ss != null) {
      ss.decompress();
      final doc = XmlDocument.parse(utf8.decode(ss.content));
      for (final si in doc.findAllElements('si')) {
        sharedStrings.add(si.findAllElements('t').map((t) => t.innerText).join());
      }
    }

    // 2. 读 rels → rId → target 映射
    final targetMap = <String, String>{};
    final rels = archive.findFile('xl/_rels/workbook.xml.rels');
    if (rels != null) {
      rels.decompress();
      final doc = XmlDocument.parse(utf8.decode(rels.content));
      for (final rel in doc.findAllElements('Relationship')) {
        final id = rel.getAttribute('Id');
        final target = rel.getAttribute('Target');
        final type = rel.getAttribute('Type') ?? '';
        if (id != null && target != null && type.contains('worksheet')) {
          String t = target;
          if (t.startsWith('/')) t = t.substring(1);
          if (!t.startsWith('xl/')) t = 'xl/$t';
          targetMap[id] = t;
        }
      }
    }

    // 3. 读 workbook.xml → 获取工作表名和 rId
    final sheetNames = <String>[];
    final sheetRIds = <String>[];
    final wb = archive.findFile('xl/workbook.xml');
    if (wb != null) {
      wb.decompress();
      final doc = XmlDocument.parse(utf8.decode(wb.content));
      for (final sheet in doc.findAllElements('sheet')) {
        final name = sheet.getAttribute('name') ?? '';
        final rId = sheet.getAttribute('r:id');
        if (rId != null) {
          sheetNames.add(name);
          sheetRIds.add(rId);
        }
      }
    }

    // 4. 逐个工作表解析（支持 sheetIndices 过滤）
    final allSheets = <List<Map<String, String>>>[];
    for (var si = 0; si < sheetRIds.length; si++) {
      // 如果指定了索引列表，跳过未选中的 sheet
      if (sheetIndices != null && !sheetIndices.contains(si)) continue;

      final rId = sheetRIds[si];
      final target = targetMap[rId];
      if (target == null) continue;
      final ws = archive.findFile(target);
      if (ws == null) continue;
      ws.decompress();
      final doc = XmlDocument.parse(utf8.decode(ws.content));

      final rows = <Map<String, String>>[];
      for (final row in doc.findAllElements('row')) {
        final rowData = <String, String>{};
        for (final cell in row.findAllElements('c')) {
          final ref = cell.getAttribute('r');
          if (ref == null) continue;
          final colLetter = ref.replaceAll(RegExp(r'\d'), '');
          final t = cell.getAttribute('t');

          String value = '';
          if (t == 's') {
            final idx = int.tryParse(cell.findAllElements('v').firstOrNull?.innerText ?? '');
            if (idx != null && idx < sharedStrings.length) {
              value = sharedStrings[idx];
            }
          } else if (t == 'inlineStr') {
            value = cell.findAllElements('t').map((e) => e.innerText).join();
          } else {
            value = cell.findAllElements('v').firstOrNull?.innerText ?? '';
            // 规范化数值：去 ".0" 后缀（PCI=14.0→14），展开科学计数法（1.3529691E7→13529691）
            final numVal = double.tryParse(value);
            if (numVal != null) {
              if (numVal == numVal.truncateToDouble()) {
                value = numVal.truncate().toString();
              } else {
                value = numVal.toString();
              }
            }
          }
          rowData[colLetter] = value;
        }
        if (rowData.isNotEmpty) rows.add(rowData);
      }
      if (rows.isNotEmpty) allSheets.add(rows);
    }

    return (sheetNames: sheetNames, sheets: allSheets);
  }

  String _pick(Map<String, String> row, List<String> keys) {
    for (final k in keys) {
      final v = row[k];
      if (v != null && v.isNotEmpty) return v;
    }
    return '';
  }

  /// 裸 ZIP 诊断 — 直接解压 xlsx 并报告内部结构
  String _diagnoseRaw(List<int> bytes) {
    final sb = StringBuffer();
    try {
      final archive = ZipDecoder().decodeBytes(bytes);
      sb.write('ZIP:${archive.files.length}files ');

      // 列出所有文件名
      final names = archive.files.take(30).map((f) => f.name).join(',');
      sb.write('files:[$names] ');

      // 读 workbook.xml 获取工作表名
      final wb = archive.findFile('xl/workbook.xml');
      if (wb != null) {
        wb.decompress();
        final doc = XmlDocument.parse(utf8.decode(wb.content));
        final sheets = doc.findAllElements('sheet')
            .map((e) => e.getAttribute('name') ?? '?')
            .join(',');
        sb.write('WBsheets:[$sheets] ');
      } else {
        sb.write('NO_workbook.xml ');
      }

      // 读 workbook.xml.rels — excel 包炸掉的关键位置
      final rels = archive.findFile('xl/_rels/workbook.xml.rels');
      if (rels != null) {
        rels.decompress();
        final doc = XmlDocument.parse(utf8.decode(rels.content));
        final targets = doc.findAllElements('Relationship')
            .map((e) => '${e.getAttribute('Id')}=>${e.getAttribute('Target')}')
            .join(',');
        sb.write('REL:$targets ');
      } else {
        sb.write('NO_rels ');
      }

      // 读 sharedStrings.xml 统计
      final ss = archive.findFile('xl/sharedStrings.xml');
      if (ss != null) {
        ss.decompress();
        final doc = XmlDocument.parse(utf8.decode(ss.content));
        final count = doc.findAllElements('si').length;
        sb.write('SScount:$count ');
      } else {
        sb.write('NO_sharedStrings ');
      }

      // 读第一个工作表的前几行
      final sheetFile = archive.findFile('xl/worksheets/sheet1.xml');
      if (sheetFile != null) {
        sheetFile.decompress();
        final doc = XmlDocument.parse(utf8.decode(sheetFile.content));
        final rows = doc.findAllElements('row').toList();
        sb.write('Sheet1rows:${rows.length} ');
        // 取第一行前5个 cell
        if (rows.isNotEmpty) {
          final cells = rows.first.findAllElements('c').toList();
          final cellInfo = cells.take(5).map((c) {
            final ref = c.getAttribute('r') ?? '?';
            final t = c.getAttribute('t') ?? '';
            final v = c.findAllElements('v').firstOrNull?.innerText ?? '';
            final inline = c.findAllElements('t').firstOrNull?.innerText ?? '';
            return '$ref${t.isNotEmpty?'t=$t':''}:${inline.isNotEmpty?inline:v}';
          }).join('|');
          sb.write('row1:[$cellInfo] ');
        }
      } else {
        sb.write('NO_sheet1.xml ');
      }
    } catch (e) {
      sb.write('ZIPerr:${e.toString().replaceAll('\n',' ')} ');
    }
    sb.write('|| ');
    return sb.toString();
  }

  /// 运营商智能检测（与桌面端逻辑一致）
  static String _detectCarrier(String contractor, String filename, Map<String, String> dict) {
    // 1. 承建方列（最可靠）
    if (contractor.contains('联通')) return '中国联通';
    if (contractor.contains('电信')) return '中国电信';
    if (contractor.contains('移动')) return '中国移动';

    // 2. 文件名
    if (filename.contains('联通')) return '中国联通';
    if (filename.contains('电信')) return '中国电信';
    if (filename.contains('移动')) return '中国移动';

    // 3. 小区名/基站名关键字
    final cn = dict['小区名'] ?? dict['小区名称'] ?? '';
    final sn = dict['基站名'] ?? dict['基站名称'] ?? '';
    final combined = '$cn$sn';
    if (combined.contains('联通')) return '中国联通';
    if (combined.contains('电信')) return '中国电信';
    if (combined.contains('移动')) return '中国移动';

    // 4. 小区名前缀判定（CJ_ → 电信，桌面端相同逻辑）
    if (cn.startsWith('CJ_') || sn.startsWith('CJ_')) return '中国电信';
    // 联通区县前缀
    const unicomPrefixes = ['CJCJS','CJFKS','CJQTX','CJMLX','CJHTB','CJMNS',
      'CJJMS','CJWJQ','CJFCH','CJWCW','CJXHN','WLMDQ','WLXSQ','WL_CJ'];
    for (final p in unicomPrefixes) {
      if (cn.startsWith(p) || sn.startsWith(p)) return '中国联通';
    }

    // 5. (LTGX) 标记 → 电信
    if (cn.contains('(LTGX)') || sn.contains('(LTGX)')) return '中国电信';

    // 6. 最终兜底：按制式推测（5G 更可能是电信/联通，4G 更可能是联通）
    final tech = _detectTech(filename);
    if (tech.contains('5G')) return '中国电信'; // 5G 默认电信
    return '中国联通'; // 4G 默认联通
  }

  static String _detectTech(String f) => f.toUpperCase().contains('5G') ? '5G NR' : '4G LTE';
}
