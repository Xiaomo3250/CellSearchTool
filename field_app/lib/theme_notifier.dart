import 'package:flutter/material.dart';
import 'config.dart';

/// 全局主题模式 —— 自动持久化到 config.json
final themeNotifier = ValueNotifier<ThemeMode>(ThemeMode.light);

/// 从配置文件恢复主题（runApp 前调用，避免闪白）
Future<void> loadTheme() async {
  final cfg = await AppConfig.load();
  themeNotifier.value = cfg.isDark ? ThemeMode.dark : ThemeMode.light;
}

/// 切换主题并保存
void toggleTheme() {
  final isDark = themeNotifier.value == ThemeMode.dark;
  themeNotifier.value = isDark ? ThemeMode.light : ThemeMode.dark;
  AppConfig.load().then((cfg) => cfg.isDark = !isDark);
}
