import 'package:sqflite/sqflite.dart';
import 'package:path/path.dart';
import '../models/station.dart';

/// SQLite 数据库层 — 建表/索引/搜索/统计
/// 与桌面版 db.py 逻辑一致
class AppDatabase {
  static final AppDatabase _instance = AppDatabase._internal();
  factory AppDatabase() => _instance;
  AppDatabase._internal();

  Database? _db;

  Future<Database> get database async {
    if (_db != null) return _db!;
    _db = await _initDb();
    return _db!;
  }

  /// 初始化数据库：建表 + 索引
  Future<Database> _initDb() async {
    final dbPath = await getDatabasesPath();
    final path = join(dbPath, 'field_cache.db');
    return openDatabase(path, version: 2,
      onCreate: (db, version) async {
      // 17 个标准化字段 + 3 个元数据字段
      await db.execute('''
        CREATE TABLE records (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          技术制式 TEXT, 运营商 TEXT, 设备商 TEXT,
          基站名 TEXT, 基站ID TEXT, 小区名 TEXT, 小区ID TEXT,
          PCI TEXT, 下行频点 TEXT, 下倾角 TEXT, 挂高 TEXT, 方位角 TEXT,
          经度 TEXT, 纬度 TEXT, 共享 TEXT, 频段 TEXT, TAC TEXT,
          _文件名 TEXT, _工作表 TEXT, _raw TEXT, _search TEXT
        )
      ''');
      // 为常用搜索字段建索引（加速精确匹配）
      for (final col in [
        '小区名', '基站名', 'PCI', '小区ID', '基站ID',
        'TAC', '下行频点', '技术制式', '_search'
      ]) {
        await db.execute('CREATE INDEX IF NOT EXISTS idx_$col ON records("$col")');
      }
    },
    onUpgrade: (db, oldVersion, newVersion) async {
      if (oldVersion < 2) {
        // v1→v2: 规范化小区ID，长格式(基站ID+小区标识)→截短标识
        await db.rawUpdate('''
          UPDATE records SET "小区ID" = substr("小区ID", length("基站ID")+1)
          WHERE length("基站ID") > 0 AND "小区ID" LIKE "基站ID" || '%'
            AND length("小区ID") > length("基站ID")
        ''');
      }
    });
  }

  /// 批量插入记录
  Future<int> insertRecords(List<Station> stations) async {
    final db = await database;
    final batch = db.batch();
    for (final s in stations) {
      batch.insert('records', s.toMap());
    }
    final results = await batch.commit(noResult: false);
    return results.length;
  }

  /// 获取统计信息（总记录数、文件数）
  Future<Map<String, int>> getStats() async {
    final db = await database;
    final total = Sqflite.firstIntValue(
      await db.rawQuery('SELECT COUNT(*) FROM records'),
    ) ?? 0;
    final files = Sqflite.firstIntValue(
      await db.rawQuery('SELECT COUNT(DISTINCT _文件名) FROM records'),
    ) ?? 0;
    return {'total': total, 'files': files};
  }

  /// 多字段搜索 — 数字字段精确匹配(=)走索引，文本字段模糊 LIKE
  Future<List<Station>> search({
    String? tech,     // 制式: 4G/5G
    String? tac,      // 跟踪区码
    String? siteId,   // 基站ID
    String? cellId,   // 小区ID
    String? pci,      // 物理小区标识
    String? freq,     // 下行频点
    String? siteName, // 基站名（模糊）
    String? cellName, // 小区名（模糊）
    int limit = 50,
    int offset = 0,
  }) async {
    final db = await database;
    final conditions = <String>[];
    final params = <dynamic>[];

    // 数字字段：纯数字→精确匹配=；含字符→LIKE
    void addEq(String col, String? val) {
      if (val == null || val.isEmpty) return;
      if (int.tryParse(val) != null) {
        conditions.add('"$col" = ?');
        params.add(val);
      } else {
        conditions.add('"$col" LIKE ?');
        params.add('%$val%');
      }
    }

    // 文本字段：始终模糊 LIKE
    void addLike(String col, String? val) {
      if (val == null || val.isEmpty) return;
      conditions.add('"$col" LIKE ?');
      params.add('%$val%');
    }

    // 制式特殊处理：简写映射
    if (tech != null && tech.isNotEmpty) {
      final t = tech.toUpperCase();
      if (t == '4G' || t == 'LTE') {
        conditions.add('"技术制式" LIKE ?');
        params.add('%4G%');
      } else if (t == '5G' || t == 'NR') {
        conditions.add('"技术制式" LIKE ?');
        params.add('%5G%');
      }
    }
    addEq('TAC', tac);
    addEq('基站ID', siteId);
    addEq('小区ID', cellId);
    addEq('PCI', pci);
    addEq('下行频点', freq);
    addLike('基站名', siteName);
    addLike('小区名', cellName);

    final where = conditions.isEmpty ? '1=1' : conditions.join(' AND ');
    final rows = await db.rawQuery(
      'SELECT * FROM records WHERE $where ORDER BY id LIMIT ? OFFSET ?',
      [...params, limit, offset],
    );
    return rows.map((r) => Station.fromMap(r)).toList();
  }

  /// 清空全部数据
  Future<void> clearAll() async {
    final db = await database;
    await db.delete('records');
  }

  /// 按文件名删除（重新导入前清理旧数据）
  Future<void> deleteByFilename(String filename) async {
    final db = await database;
    await db.delete('records', where: '_文件名 = ?', whereArgs: [filename]);
  }

  /// 获取已导入文件列表（含记录数）
  Future<List<Map<String, dynamic>>> getFileStats() async {
    final db = await database;
    return db.rawQuery('''
      SELECT _文件名 as name, COUNT(*) as count,
             MAX(技术制式) as tech
      FROM records GROUP BY _文件名 ORDER BY _文件名
    ''');
  }

  Future<void> close() async {
    await _db?.close();
    _db = null;
  }
}
