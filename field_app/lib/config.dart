import 'dart:convert';
import 'dart:io';
import 'package:path_provider/path_provider.dart';

/// 简单的 JSON 配置文件读写
class AppConfig {
  static AppConfig? _instance;
  late String _path;
  Map<String, dynamic> _data = {};

  AppConfig._();

  static Future<AppConfig> load() async {
    if (_instance != null) return _instance!;
    _instance = AppConfig._();
    final dir = await getApplicationDocumentsDirectory();
    _instance!._path = '${dir.path}/config.json';
    try {
      final file = File(_instance!._path);
      if (await file.exists()) {
        _instance!._data = jsonDecode(await file.readAsString()) as Map<String, dynamic>;
      }
    } catch (_) {
      _instance!._data = {};
    }
    return _instance!;
  }

  bool get isDark => _data['isDark'] == true;
  set isDark(bool v) {
    _data['isDark'] = v;
    _save();
  }

  Future<void> _save() async {
    try {
      await File(_path).writeAsString(jsonEncode(_data));
    } catch (_) {}
  }
}
