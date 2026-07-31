package com.example.field_app

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import io.flutter.embedding.android.FlutterActivity
import io.flutter.embedding.engine.FlutterEngine
import io.flutter.plugin.common.MethodChannel

/// 接收外部 APP（微信/QQ/钉钉）分享的 Excel 文件，传给 Flutter 端导入
class MainActivity : FlutterActivity() {
    private val CHANNEL = "com.example.field_app/shared"
    private var sharedFilePath: String? = null
    private var channel: MethodChannel? = null

    override fun configureFlutterEngine(flutterEngine: FlutterEngine) {
        super.configureFlutterEngine(flutterEngine)
        channel = MethodChannel(flutterEngine.dartExecutor.binaryMessenger, CHANNEL)
        channel?.setMethodCallHandler { call, result ->
            if (call.method == "getSharedFile") {
                val path = sharedFilePath
                sharedFilePath = null
                result.success(path)
            } else {
                result.notImplemented()
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        handleIntent(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleIntent(intent)
    }

    /// 解析分享 Intent，支持 ACTION_VIEW（文件管理器打开）和 ACTION_SEND（微信/QQ/钉钉分享）
    private fun handleIntent(intent: Intent?) {
        if (intent == null) return
        var uri: Uri? = null
        when (intent.action) {
            Intent.ACTION_VIEW -> uri = intent.data
            Intent.ACTION_SEND -> uri = intent.getParcelableExtra(Intent.EXTRA_STREAM)
            else -> return
        }
        if (uri == null) return

        try {
            val filename = getFileName(uri) ?: "imported.xlsx"
            val cacheFile = java.io.File(cacheDir, filename)
            contentResolver.openInputStream(uri)?.use { input ->
                cacheFile.outputStream().use { output ->
                    input.copyTo(output)
                }
            }
            sharedFilePath = "${cacheFile.absolutePath}|${filename}"
            // 主动通知 Flutter 端文件已就绪
            channel?.invokeMethod("onFileReady", sharedFilePath)
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    /// 尝试从 URI 中提取文件名
    private fun getFileName(uri: Uri): String? {
        var name: String? = null
        try {
            contentResolver.query(uri, null, null, null, null)?.use { cursor ->
                if (cursor.moveToFirst()) {
                    val idx = cursor.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
                    if (idx >= 0) name = cursor.getString(idx)
                }
            }
        } catch (_: Exception) {}
        if (name != null) return name
        return uri.lastPathSegment
    }
}

