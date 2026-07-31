import 'package:flutter/material.dart';
import '../db/database.dart';
import '../theme_notifier.dart';

/// 数据管理页 — 查看/删除已导入的工参文件
class DataManagerPage extends StatefulWidget {
  const DataManagerPage({super.key});

  @override
  State<DataManagerPage> createState() => _DataManagerPageState();
}

class _DataManagerPageState extends State<DataManagerPage> {
  final _db = AppDatabase();
  List<Map<String, dynamic>> _files = [];
  bool _loading = true;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    setState(() => _loading = true);
    final files = await _db.getFileStats();
    final total = files.fold<int>(0, (sum, f) => sum + ((f['count'] as int?) ?? 0));
    if (mounted) {
      setState(() {
        _files = files;
        _loading = false;
      });
    }
  }

  Future<void> _deleteFile(String filename) async {
    final ok = await showDialog<bool>(
      context: context,
      builder: (_) => AlertDialog(
        title: const Text('删除确认'),
        content: Text('将删除「$filename」的全部数据，不可恢复。'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context, false), child: const Text('取消')),
          TextButton(
            onPressed: () => Navigator.pop(context, true),
            child: const Text('删除', style: TextStyle(color: Colors.red)),
          ),
        ],
      ),
    );
    if (ok == true) {
      await _db.deleteByFilename(filename);
      await _load();
    }
  }

  Future<void> _clearAll() async {
    final ok = await showDialog<bool>(
      context: context,
      builder: (_) => AlertDialog(
        title: const Text('清空全部'),
        content: Text('将删除全部 ${_files.fold<int>(0, (s, f) => s + ((f['count'] as int?) ?? 0))} 条记录，不可恢复。'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context, false), child: const Text('取消')),
          TextButton(
            onPressed: () => Navigator.pop(context, true),
            child: const Text('清空', style: TextStyle(color: Colors.red)),
          ),
        ],
      ),
    );
    if (ok == true) {
      await _db.clearAll();
      await _load();
    }
  }

  @override
  Widget build(BuildContext context) {
    final total = _files.fold<int>(0, (s, f) => s + ((f['count'] as int?) ?? 0));
    return Scaffold(
      appBar: AppBar(
        title: const Text('数据管理'),
        actions: [
          // 黑夜模式切换
          ValueListenableBuilder<ThemeMode>(
            valueListenable: themeNotifier,
            builder: (_, mode, __) => IconButton(
              icon: Icon(mode == ThemeMode.dark ? Icons.light_mode : Icons.dark_mode),
              tooltip: mode == ThemeMode.dark ? '切换日间模式' : '切换黑夜模式',
              onPressed: toggleTheme,
            ),
          ),
          if (_files.isNotEmpty)
            IconButton(
              icon: const Icon(Icons.delete_sweep),
              tooltip: '清空全部',
              onPressed: _clearAll,
            ),
        ],
      ),
      body: _loading
          ? const Center(child: CircularProgressIndicator())
          : _files.isEmpty
              ? const Center(child: Text('暂无导入数据'))
              : Column(
                  children: [
                    // 汇总栏
                    Container(
                      width: double.infinity,
                      padding: const EdgeInsets.all(12),
                      color: Theme.of(context).colorScheme.surfaceContainerHighest,
                      child: Text('共 ${_files.length} 个文件 · $total 条记录',
                          style: const TextStyle(fontWeight: FontWeight.bold)),
                    ),
                    Expanded(
                      child: ListView.builder(
                        itemCount: _files.length,
                        itemBuilder: (_, i) {
                          final f = _files[i];
                          final name = f['name'] as String? ?? '?';
                          final count = f['count'] as int? ?? 0;
                          final tech = f['tech'] as String? ?? '';
                          return Card(
                            margin: const EdgeInsets.symmetric(horizontal: 12, vertical: 4),
                            child: ListTile(
                              leading: const Icon(Icons.table_chart),
                              title: Text(name, overflow: TextOverflow.ellipsis),
                              subtitle: Text('$count 条 · $tech'),
                              trailing: IconButton(
                                icon: const Icon(Icons.delete_outline, color: Colors.red),
                                onPressed: () => _deleteFile(name),
                              ),
                            ),
                          );
                        },
                      ),
                    ),
                  ],
                ),
    );
  }
}
