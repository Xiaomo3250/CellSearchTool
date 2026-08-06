import 'dart:io';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:file_picker/file_picker.dart';
import 'db/database.dart';
import 'db/importer.dart';
import 'models/station.dart';
import 'pages/data_manager.dart';
import 'theme_notifier.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await loadTheme(); // 启动前恢复主题，避免闪白
  runApp(const FieldApp());
}

/// 应用根组件
class FieldApp extends StatelessWidget {
  const FieldApp({super.key});
  @override
  Widget build(BuildContext context) {
    return ValueListenableBuilder<ThemeMode>(
      valueListenable: themeNotifier,
      builder: (_, mode, __) => MaterialApp(
        title: '基站工参管理器',
        debugShowCheckedModeBanner: false,
        theme: ThemeData(colorSchemeSeed: Colors.blue, useMaterial3: true, brightness: Brightness.light),
        darkTheme: ThemeData(colorSchemeSeed: Colors.blue, useMaterial3: true, brightness: Brightness.dark),
        themeMode: mode,
        home: const SearchScreen(),
      ),
    );
  }
}

/// 主界面 — 搜索面板 + 卡片列表
class SearchScreen extends StatefulWidget {
  const SearchScreen({super.key});
  @override
  State<SearchScreen> createState() => _SearchScreenState();
}

class _SearchScreenState extends State<SearchScreen> with WidgetsBindingObserver {
  final _db = AppDatabase();
  final _importer = Importer();

  // 搜索结果状态
  List<Station> _results = [];
  int _totalRecords = 0;
  int _fileCount = 0;
  bool _loading = false;
  int _page = 0;
  static const _perPage = 50;

  // 搜索字段控制器
  final _techCtrl = TextEditingController();
  final _tacCtrl = TextEditingController();
  final _bidCtrl = TextEditingController();
  final _cidCtrl = TextEditingController();
  final _pciCtrl = TextEditingController();
  final _freqCtrl = TextEditingController();
  final _bsNameCtrl = TextEditingController();
  final _cellNameCtrl = TextEditingController();
  String _techValue = ''; // 制式下拉框选中值

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _loadStats();
    _setupSharedFileListener();
  }

  /// 监听 Android 端主动推送的分享文件通知
  void _setupSharedFileListener() {
    const channel = MethodChannel('com.example.field_app/shared');
    channel.setMethodCallHandler((call) async {
      if (call.method == 'onFileReady') {
        final data = call.arguments as String?;
        if (data == null || data.isEmpty) return;
        final parts = data.split('|');
        if (parts.length != 2) return;
        final filePath = parts[0];
        final filename = parts[1];
        await _importSharedFile(filePath, filename);
      }
    });
    // 冷启动兜底：也轮询一次 getSharedFile
    _pollSharedFileOnce();
  }

  Future<void> _pollSharedFileOnce() async {
    const channel = MethodChannel('com.example.field_app/shared');
    for (var i = 0; i < 3; i++) {
      try {
        final result = await channel.invokeMethod<String?>('getSharedFile');
        if (result != null && result.isNotEmpty) {
          final parts = result.split('|');
          if (parts.length == 2) {
            await _importSharedFile(parts[0], parts[1]);
            return;
          }
        }
      } catch (_) {}
      await Future.delayed(const Duration(milliseconds: 500));
    }
  }

  /// 分享文件入口 → 走统一导入流程
  Future<void> _importSharedFile(String filePath, String filename) async {
    final file = File(filePath);
    if (!await file.exists()) return;
    final bytes = await file.readAsBytes();
    await file.delete(); // 清理缓存
    if (mounted) await _importWithUI(filename, bytes);
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    _techCtrl.dispose(); _tacCtrl.dispose(); _bidCtrl.dispose();
    _cidCtrl.dispose(); _pciCtrl.dispose(); _freqCtrl.dispose();
    _bsNameCtrl.dispose(); _cellNameCtrl.dispose();
    super.dispose();
  }

  /// APP 从后台恢复时重新检查分享文件
  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    if (state == AppLifecycleState.resumed) {
      _pollSharedFileOnce();
    }
  }

  /// 加载数据库统计信息
  Future<void> _loadStats() async {
    final stats = await _db.getStats();
    setState(() {
      _totalRecords = stats['total'] ?? 0;
      _fileCount = stats['files'] ?? 0;
    });
    if (_totalRecords > 0) _doSearch(reset: true);
  }

  /// 执行多字段搜索
  /// [reset]=true 时重置分页
  Future<void> _doSearch({bool reset = false}) async {
    if (reset) {
      _page = 0;
      _results.clear();
    }
    setState(() => _loading = true);

    final list = await _db.search(
      tech: _techValue,
      tac: _tacCtrl.text.trim(),
      siteId: _bidCtrl.text.trim(),
      cellId: _cidCtrl.text.trim(),
      pci: _pciCtrl.text.trim(),
      freq: _freqCtrl.text.trim(),
      siteName: _bsNameCtrl.text.trim(),
      cellName: _cellNameCtrl.text.trim(),
      limit: _perPage,
      offset: _page * _perPage,
    );

    setState(() {
      if (reset) {
        _results = list;
      } else {
        _results.addAll(list); // 瀑布流追加
      }
      _page++;
      _loading = false;
    });
  }

  /// Sheet 多选对话框（默认全不选，手动勾选）
  Future<Set<int>?> _showSheetPicker(List<String> names) async {
    final sel = <int>{};
    return showDialog<Set<int>>(
      context: context,
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setD) => AlertDialog(
          title: const Text('选择工作表'),
          content: SizedBox(
            width: double.maxFinite,
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                // 全选/全不选 快捷按钮
                Row(
                  children: [
                    TextButton(
                      onPressed: () => setD(() => sel.addAll(Iterable.generate(names.length))),
                      child: const Text('全选'),
                    ),
                    TextButton(
                      onPressed: () => setD(() => sel.clear()),
                      child: const Text('全不选'),
                    ),
                  ],
                ),
                Flexible(
                  child: ListView(
                    shrinkWrap: true,
                    children: [
                      for (var i = 0; i < names.length; i++)
                        CheckboxListTile(
                          title: Text(names[i]),
                          value: sel.contains(i),
                          onChanged: (v) => setD(() => v == true ? sel.add(i) : sel.remove(i)),
                        ),
                    ],
                  ),
                ),
              ],
            ),
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(ctx, null), child: const Text('取消')),
            TextButton(
              onPressed: () => Navigator.pop(ctx, sel),
              child: Text('导入选中${sel.isNotEmpty ? '(${sel.length})' : ''}'),
            ),
          ],
        ),
      ),
    );
  }

  /// 文件选择器 — 选文件后走统一导入流程
  Future<void> _pickAndImport() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
      allowMultiple: true,
      withData: true,
    );
    if (result == null || result.files.isEmpty) return;

    for (final file in result.files) {
      List<int>? bytes = file.bytes;
      if (bytes == null && file.path != null) {
        try { bytes = await File(file.path!).readAsBytes(); } catch (_) {}
      }
      if (bytes != null && mounted) {
        await _importWithUI(file.name, bytes);
      }
    }
  }

  /// 统一导入流程：选 sheet → 加载弹窗 → 导入 → 结果弹窗
  Future<void> _importWithUI(String filename, List<int> bytes) async {
    // 1. 读 sheet 名 & 多选
    final sheetNames = Importer.listSheetNames(bytes);
    Set<int>? indices;
    if (sheetNames.length > 1 && mounted) {
      indices = await _showSheetPicker(sheetNames);
      if (indices == null) return;
    }

    // 2. 加载弹窗
    if (mounted) {
      showDialog(
        context: context,
        barrierDismissible: false,
        builder: (_) => const AlertDialog(
          title: Text('导入中...'),
          content: SizedBox(height: 60, child: Center(child: CircularProgressIndicator())),
        ),
      );
    }

    // 3. 执行导入
    int totalOk = 0, totalRows = 0, skipEmpty = 0, skipLng = 0;
    String sheets = '', diag = '';
    try {
      final info = await _importer.importFromBytes(bytes, filename, sheetIndices: indices);
      totalOk = info['ok'] as int;
      totalRows = info['total'] as int;
      skipEmpty = info['skipEmpty'] as int;
      skipLng = info['skipLng'] as int;
      sheets = (info['sheets'] as List).join(', ');
      diag = info['error'] as String? ?? '';
    } catch (e) {
      diag = '导入异常: $e';
    }
    await _loadStats();

    // 4. 关闭加载弹窗 & 显示结果
    if (mounted) {
      Navigator.pop(context);
      final body = StringBuffer();
      body.writeln('$totalOk 条 \u2714');
      body.writeln('扫描 $totalRows 行 | 空行 $skipEmpty | 经纬异常 $skipLng');
      body.writeln('工作表: ${sheets.isNotEmpty ? sheets : "(无)"}');
      if (diag.isNotEmpty) {
        body.writeln('---诊断---');
        body.write(diag);
      }
      showDialog(
        context: context,
        builder: (_) => AlertDialog(
          title: const Text('导入结果'),
          content: SizedBox(
            width: double.maxFinite,
            child: SingleChildScrollView(
              child: Text(body.toString(), style: const TextStyle(fontSize: 13)),
            ),
          ),
          actions: [
            TextButton(
              onPressed: () => Clipboard.setData(ClipboardData(text: body.toString())),
              child: const Text('复制信息'),
            ),
            TextButton(onPressed: () => Navigator.pop(context), child: const Text('确定')),
          ],
        ),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('基站工参管理器'),
        actions: [
          // 顶栏统计
          Center(child: Text('$_totalRecords条 | $_fileCount个文件',
              style: const TextStyle(fontSize: 12))),
          // 数据管理
          IconButton(
            icon: const Icon(Icons.folder_open),
            tooltip: '数据管理',
            onPressed: () => Navigator.push(context,
                  MaterialPageRoute(builder: (_) => const DataManagerPage()))
                  .then((_) => _loadStats()),
          ),
          // 导入按钮
          IconButton(
            icon: const Icon(Icons.file_upload),
            tooltip: '导入工参',
            onPressed: _pickAndImport,
          ),
        ],
      ),
      body: Column(children: [
        _buildSearchPanel(), // 搜索面板
        Expanded(
          child: _loading && _results.isEmpty
              ? const Center(child: CircularProgressIndicator())
              : _results.isEmpty
                  ? const Center(child: Text('暂无数据，请导入工参文件'))
                  : _buildCardList(), // 卡片瀑布流
        ),
      ]),
    );
  }

  /// 构建搜索面板 — 2x4 网格布局
  Widget _buildSearchPanel() {
    return Container(
      padding: const EdgeInsets.all(12),
      color: Theme.of(context).colorScheme.surfaceContainerHighest,
      child: Column(children: [
        // 第1行：制式 + TAC
        Row(children: [
          Expanded(
            child: DropdownButtonFormField<String>(
              value: _techValue.isEmpty ? null : _techValue,
              decoration: const InputDecoration(
                  labelText: '制式', isDense: true,
                  contentPadding: EdgeInsets.symmetric(horizontal: 8, vertical: 8)),
              items: const [
                DropdownMenuItem(value: '', child: Text('全部')),
                DropdownMenuItem(value: '4G', child: Text('4G LTE')),
                DropdownMenuItem(value: '5G', child: Text('5G NR')),
              ],
              onChanged: (v) => setState(() => _techValue = v ?? ''),
            ),
          ),
          const SizedBox(width: 8),
          Expanded(child: TextField(controller: _tacCtrl,
              decoration: const InputDecoration(labelText: 'TAC', isDense: true))),
        ]),
        const SizedBox(height: 8),
        // 第2行：基站ID + CellID（主要搜索维度）
        Row(children: [
          Expanded(child: TextField(controller: _bidCtrl,
              decoration: const InputDecoration(labelText: '基站ID', isDense: true))),
          const SizedBox(width: 8),
          Expanded(child: TextField(controller: _cidCtrl,
              decoration: const InputDecoration(labelText: 'CellID', isDense: true))),
        ]),
        const SizedBox(height: 8),
        // 第3行：PCI + 频点
        Row(children: [
          Expanded(child: TextField(controller: _pciCtrl,
              decoration: const InputDecoration(labelText: 'PCI', isDense: true))),
          const SizedBox(width: 8),
          Expanded(child: TextField(controller: _freqCtrl,
              decoration: const InputDecoration(labelText: '频点', isDense: true))),
        ]),
        const SizedBox(height: 8),
        // 第4行：基站名 + 小区名（文本模糊搜索）
        Row(children: [
          Expanded(child: TextField(controller: _bsNameCtrl,
              decoration: const InputDecoration(labelText: '基站名', isDense: true))),
          const SizedBox(width: 8),
          Expanded(child: TextField(controller: _cellNameCtrl,
              decoration: const InputDecoration(labelText: '小区名', isDense: true))),
        ]),
        const SizedBox(height: 8),
        // 操作按钮
        Row(mainAxisAlignment: MainAxisAlignment.end, children: [
          TextButton(onPressed: () {
            // 清空所有搜索条件
            _techCtrl.clear(); _tacCtrl.clear(); _bidCtrl.clear();
            _cidCtrl.clear(); _pciCtrl.clear(); _freqCtrl.clear();
            _bsNameCtrl.clear(); _cellNameCtrl.clear();
            setState(() => _techValue = '');
            _doSearch(reset: true);
          }, child: const Text('清空')),
          const SizedBox(width: 8),
          FilledButton(
              onPressed: () => _doSearch(reset: true),
              child: const Text('搜索')),
        ]),
      ]),
    );
  }

  /// 构建卡片瀑布流列表（滚动到底自动加载更多）
  Widget _buildCardList() {
    return NotificationListener<ScrollNotification>(
      onNotification: (n) {
        if (n is ScrollEndNotification &&
            n.metrics.extentAfter < 100 && !_loading) {
          _doSearch(); // 接近底部时加载下一页
        }
        return false;
      },
      child: ListView.builder(
        padding: const EdgeInsets.all(8),
        itemCount: _results.length,
        itemBuilder: (_, i) => _buildCard(_results[i]),
      ),
    );
  }

  /// 构建单张查询卡片
  Widget _buildCard(Station s) {
    final is4G = s.tech.contains('4G');
    final isDX = s.carrier.contains('电信');

    return Card(
      margin: const EdgeInsets.only(bottom: 8),
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // ---- 左侧：工参信息 ----
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  // 制式 + 运营商 + 设备商
                  Row(children: [
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                      decoration: BoxDecoration(
                        color: is4G ? Colors.orange.shade100 : Colors.blue.shade100,
                        borderRadius: BorderRadius.circular(4)),
                      child: Text(s.techLabel,
                        style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold,
                          color: is4G ? Colors.orange.shade800 : Colors.blue.shade800)),
                    ),
                    const SizedBox(width: 4),
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                      decoration: BoxDecoration(
                        color: isDX ? Colors.teal.shade50 : Colors.purple.shade50,
                        borderRadius: BorderRadius.circular(4)),
                      child: Text(s.shortCarrier,
                        style: TextStyle(fontSize: 11,
                          color: isDX ? Colors.teal.shade700 : Colors.purple.shade700)),
                    ),
                    if (s.vendor.isNotEmpty) ...[
                      const SizedBox(width: 4),
                      Container(
                        padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                        decoration: BoxDecoration(
                          color: Colors.grey.shade100,
                          borderRadius: BorderRadius.circular(4)),
                        child: Text(s.vendor,
                          style: TextStyle(fontSize: 11, color: Colors.grey.shade700)),
                      ),
                    ],
                  ]),
                  const SizedBox(height: 6),
                  // 小区名（加粗，完整显示）
                  Text(s.cellName,
                    style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                  const SizedBox(height: 4),
                  // PCI + 频点
                  Text('PCI: ${s.pci}  频点: ${s.dlFreq}',
                    style: TextStyle(fontSize: 13, color: Colors.grey.shade700)),
                  const SizedBox(height: 2),
                  // 基站ID + CellID
                  Text('基站ID: ${s.siteId}  CellID: ${s.cellId}',
                    style: TextStyle(fontSize: 13, color: Colors.grey.shade700)),
                ],
              ),
            ),
            // ---- 右侧：复制按钮 ----
            Column(children: [
              OutlinedButton.icon(
                onPressed: () => _copy(s.cellName, '小区名'),
                icon: const Icon(Icons.copy, size: 14),
                label: const Text('小区名', style: TextStyle(fontSize: 11)),
                style: OutlinedButton.styleFrom(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4)),
              ),
              const SizedBox(height: 4),
              FilledButton.tonalIcon(
                onPressed: () => _copy(s.quickRef, 'ID'),
                icon: const Icon(Icons.bolt, size: 14),
                label: const Text('ID', style: TextStyle(fontSize: 11)),
                style: FilledButton.styleFrom(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4)),
              ),
              const SizedBox(height: 4),
              OutlinedButton.icon(
                onPressed: () => _copy('${s.lng},${s.lat}', '经纬度'),
                icon: const Icon(Icons.location_on, size: 14),
                label: const Text('经纬度', style: TextStyle(fontSize: 11)),
                style: OutlinedButton.styleFrom(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4)),
              ),
            ]),
          ],
        ),
      ),
    );
  }

  /// 复制文本到剪贴板
  void _copy(String text, String label) {
    Clipboard.setData(ClipboardData(text: text));
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('已复制$label'), duration: const Duration(seconds: 1)));
  }
}
