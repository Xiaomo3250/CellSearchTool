# 工参管理器 项目交接文档

> 当前版本：**v0.3.0-alpha** | 交接日期：2026-05-05  
> 开发环境：WorkBuddy → 迁移至 VSCode + GitHub Copilot  
> 仓库：https://github.com/Xiaomo3250/CellSearchTool

---

## 一、项目概述

基站工参管理器是一个 **Python Web 应用**（前后端分离），用于电信/联通基站工参数据的批量导入、模糊搜索和导出。

- **端口**：`18888`
- **技术栈**：Python 3.8+（内置 http.server）+ 独立 HTML/CSS/JS 前端
- **依赖**：pandas + openpyxl
- **数据存储**：内存（records 列表）+ 磁盘缓存（cache.json）

---

## 二、文件清单（核心 7 文件）

| 文件 | 说明 |
|------|------|
| `app.py` | 后端 API 服务器（~800 行），包含数据模型 + 全部业务逻辑 |
| `static/index.html` | 前端 HTML 页面结构 |
| `static/css/style.css` | 样式表 |
| `static/js/app.js` | 前端交互逻辑（搜索/分页/导入/导出） |
| `README.md` | 功能说明和快速开始指南 |
| `.gitignore` | Git 排除规则 |
| `启动.bat` | Windows 一键启动脚本（含 Python 环境和依赖检测） |

---

## 三、架构概览

```
┌──────────────────────────────────────────────┐
│  浏览器 (http://127.0.0.1:18888)              │
│  ┌──────────────────────────────────────┐    │
│  │  前端 (static/ 独立文件)               │    │
│  │  - static/index.html  (页面结构)       │    │
│  │  - static/css/style.css (样式)         │    │
│  │  - static/js/app.js    (交互逻辑)      │    │
│  │  功能: 文件管理/搜索框/统计/分页/导出    │    │
│  └──────────────┬───────────────────────┘    │
└─────────────────│────────────────────────────┘
                  │ REST API + 静态文件请求
┌─────────────────│────────────────────────────┐
│  Python HTTP Server (ThreadingHTTPServer)     │
│  ┌──────────────────────────────────────┐    │
│  │  /, /index.html    → 静态 HTML 页面    │    │
│  │  /static/*         → CSS/JS 静态资源   │    │
│  │  /api/stats        → 统计数据           │    │
│  │  /api/search       → 模糊搜索+分页      │    │
│  │  /api/export       → 导出 CSV           │    │
│  │  /api/preview      → 文件预览(sheet列表) │    │
│  │  /api/import-sheets-async → 异步导入     │    │
│  │  /api/remove-sheet → 移除文件/工作表     │    │
│  │  /api/import-status → 导入进度          │    │
│  │  /api/clear        → 清空全部数据        │    │
│  │  /api/query-cell   → 小区查询（报告用）  │    │
│  │  /api/cell-by-location → 经纬度附近查询  │    │
│  └──────────────┬───────────────────────┘    │
│                  │                            │
│  ┌───────────────▼───────────────────────┐    │
│  │  核心逻辑                               │    │
│  │  - MAPPING_RULES: 字段映射规则引擎      │    │
│  │  - detect_carrier_from_names(): 命名检测│    │
│  │  - import_file_sheets(): 逐条导入       │    │
│  │  - search_records(): 搜索+排序         │    │
│  │  - serve_static(): 静态文件服务+安全检查 │    │
│  └──────────────┬───────────────────────┘    │
│                  │                            │
│  ┌───────────────▼───────────────────────┐    │
│  │  数据层                                 │    │
│  │  - records[]  : 内存记录列表 (thread-safe)│   │
│  │  - file_sources: 文件来源元数据          │    │
│  │  - cache.json : 磁盘缓存               │    │
│  └──────────────────────────────────────┘    │
└──────────────────────────────────────────────┘
```

**v0.2.0 关键变更**：
- **前后端分离**：前端 HTML/CSS/JS 拆分为 `static/` 下独立文件，后端从文件读取并服务
- **静态文件安全**：`serve_static()` 使用 `os.path.normpath` 防路径穿越
- **`启动.bat` 升级**：三阶段诊断（Python 存根检测 + 实际可执行验证 + 依赖检查与自动安装）
- **电信工参共享字段修复**：`field_map` 列名改为 `["是否共享", "共享方"]`，导入逻辑改用 Excel 原值
- **UI 优化**：文件卡片单行横向滚动、分页支持页码跳转输入框

---

## 四、字段映射规则（MAPPING_RULES）

4 套规则按**文件名关键词**自动匹配：

| # | carrier | tech | sheet 匹配关键词 | 特殊说明 |
|---|---------|------|-----------------|---------|
| 1 | 电信 | 5G | `5G汇总工参` / `5G工参` / `汇总工参` | |
| 2 | 电信 | 4G | `电信工参` / `LTE工参` / `室外宏站` / `室内` | |
| 3 | 联通 | 4G | 无（所有 sheet 都匹配） | sheet_keywords=None |
| 4 | 联通 | 5G | 无（排除 `行政区划`） | v0.1.2 新增 |

**字段映射示例**（以电信5G为例）：
```python
"field_map": {
    "运营商": "中国电信",           # 硬编码值
    "技术制式": "5G NR",            # 硬编码值
    "设备商": ["厂家"],             # 多候选列名优先级匹配
    "基站名": ["基站名称", "基站/楼盘名称", "站址"],
    "小区名": ["NR小区名称"],
    "下行频点": ["SSB绝对信道号", "下行频点"],  # 电信5G 实际列名是 "SSB绝对信道号"
    "共享": ["是否共享", "共享方"],  # v0.2.0: 电信工参实际列名是"是否共享"
    # ...
}
```

**v0.2.0 共享字段处理逻辑**：
- 电信工参：`运营商` 由命名检测判定，`共享` 以 Excel 的"是否共享"列为准（值="是"→"共享"，否则→"非共享"）
- 联通工参：`运营商` 和 `共享` 均由命名检测判定（`(LTGX)` 标识共享站）
- 电信 5G "下行频点"候选列名：`["SSB绝对信道号", "下行频点"]`（实际列名在前）

---

## 五、运营商归属智能检测（核心特性 v0.1.2）

### 检测函数

```python
def detect_carrier_from_names(cell_name, station_name) -> (carrier, share_type)
```

### 三级优先级

```
优先级 1: 联通共享站
  ├─ 触发条件: 小区名或基站名含 "(LTGX)"
  └─ 结果: carrier="中国联通", share="共享站"

优先级 2: 联通自建
  ├─ 触发条件: "_"分割后段数=2 且 首段以联通前缀开头
  └─ 结果: carrier="中国联通", share="自建"

优先级 3: 电信
  ├─ 触发条件: "_"分割后段数∈{5,6,7} 且 首段="CJ" 且
  │   第二段匹配 [A-Z]{2,4}\d?$ 且 末段为数字或以E/C结尾
  └─ 结果: carrier="中国电信", share="非共享"

回退: 以上均不匹配 → 文件名关键词判定
```

### 联通区县前缀（11个）

```
CJCJS(昌吉) CJFKS(阜康) CJQTX(奇台) CJMLX(木垒) CJHTB(呼图壁)
CJMNS(玛纳斯) CJJMS(吉木萨尔) CJWJQ(五家渠) CJFCH(芳草湖)
CJWCW(五彩湾) CJXHN(新湖农场)
```

### 覆盖率

| 场景 | 覆盖率 | 备注 |
|------|--------|------|
| 联通5G工参 | 100%（6377条） | 2627自建 + 3750共享 |
| 全量缓存 | 99.0%（36305/36680） | 375条电信室分站走fallback |

### 调用位置

`import_file_sheets()` 中逐条记录调用，覆盖 `运营商` 字段。

**v0.2.0 逻辑变更**：
- 命中电信时：`运营商` 设为"中国电信"，`共享` 以 Excel "是否共享"列为准，不再硬写"非共享"
- 命中联通时：`运营商` 和 `共享` 均由命名检测结果覆盖
- 未命中时：回退到文件名判定 + 列值标准化

---

## 六、基站/小区命名规范速查

### 小区名（CellName）三段式

| 类型 | 格式 | 示例 |
|------|------|------|
| 电信5G/4G（6段） | `CJ_QT0_描述N_编码1_编码2_序号` | `CJ_QT0_北塔山牧业一队N_GHCNN_PRL5E_0` |
| 联通自建（2段） | `前缀描述_编码` | `CJCJS昌吉天佑房产L(R)_GBF1A` |
| 联通共享站（5段） | `CJ_QT0_编号描述_编码1_编码2E-序号` | `CJ_QT0_829045奇台二畦基站F_TSCNN_SRT4E-1` |

### 基站名（5段 vs 2段）

- **基站名和小区名是两套独立体系，不能互相推导**
- 电信：`CJ_{区县}_{描述}_{编码1}_{编码2}`（编码2无E=5G，E结尾=4G，C结尾=CDMA）
- 联通自建：`{前缀}{描述}_{编码}`（1基站对多小区）

### 区县码对照

| 电信码 | 联通前缀 | 区县 |
|--------|----------|------|
| CJ0~CJ8 | CJCJS | 昌吉市 |
| FK0 | CJFKS | 阜康 |
| QT0 | CJQTX | 奇台 |
| ML0 | CJMLX | 木垒 |
| HTB0 | CJHTB | 呼图壁 |
| MNS0 | CJMNS | 玛纳斯 |
| JMSE0 | CJJMS | 吉木萨尔/准东 |
| WJQ0 | CJWJQ | 五家渠 |
| ZD0 | CJWCW | 准东/五彩湾 |
| - | CJFCH | 芳草湖 |
| - | CJXHN | 新湖农场 |

---

## 七、API 端点速查

| 端点 | 方法 | 参数 | 说明 |
|------|------|------|------|
| `/` | GET | - | 返回 `static/index.html` 前端页面 |
| `/static/*` | GET | 路径 | 静态资源（CSS/JS），含路径穿越防护 |
| `/api/stats` | GET | - | 统计数据（总数/基站数/文件数） |
| `/api/search` | GET | `q`, `page`, `per_page` | 模糊搜索+分页 |
| `/api/export` | GET | `q` | 导出 CSV |
| `/api/preview` | POST | multipart file | 预览文件 sheet 列表 |
| `/api/import-sheets-async` | POST | `{files: [{filename, sheets}]}` (JSON) | 异步批量导入 |
| `/api/remove-sheet` | POST | `{filename, sheets?}` (JSON) | 移除文件/工作表 |
| `/api/import-status` | GET | - | 当前导入进度 |
| `/api/clear` | POST | - | 清空全部数据 |
| `/api/query-cell` | GET | `q`, `exact?`, `limit?` | 小区查询（报告用） |
| `/api/cell-by-location` | GET | `lng`, `lat`, `radius?` | 经纬度附近基站查询 |

---

## 八、如何运行

### 环境要求

- Python 3.8+
- pip install pandas openpyxl

### 方式一：命令行

```bash
cd 0503-工参管理器
python app.py
```
浏览器自动打开 `http://127.0.0.1:18888`

### 方式二：启动脚本（Windows）

双击 `启动.bat`。脚本内置三阶段诊断：
1. 检测 Python 文件是否存在
2. 验证 Python 能实际执行（排除 Windows Store 存根）
3. 检查 pandas/openpyxl 依赖，缺失则自动安装

同时自动释放占用端口并打开浏览器。Python 路径已配置为：
```batch
set "PYTHON=C:\Users\12931\AppData\Local\Python\pythoncore-3.14-64\python.exe"
```

### 方式三：VSCode 中调试

在 `.vscode/launch.json` 中配置：
```json
{
    "name": "工参管理器",
    "type": "debugpy",
    "request": "launch",
    "program": "${workspaceFolder}/app.py",
    "cwd": "${workspaceFolder}",
    "console": "integratedTerminal"
}
```

---

## 九、运行流程

1. 启动时检查 `cache.json`，有则加载历史数据
2. 打开浏览器访问 Web 界面
3. **导入数据**：
   - 方式A：上传 xlsx 文件 → 预览 sheet 列表 → 选择导入
   - 方式B：调用 `/api/import` 批量导入
4. **搜索**：输入小区名/PCI/基站ID → 实时返回结果
5. **导出**：搜索结果可一键导出 CSV
6. 关闭时自动保存 `cache.json`

---

## 十、后续开发建议

### 已知问题

1. **375 条电信室分站**无法通过命名规则识别，完全依赖文件名 fallback
2. **联通4G MAPPING_RULE** 无 sheet_keywords（sheet_keywords=None），可能误匹配无关 sheet
3. **大文件导入时页面阻塞**——导入在请求线程中同步执行，前端通过轮询 `/api/import-status` 获取进度

### 已完成的优化（v0.2.0）

| 优化 | 说明 |
|------|------|
| ✅ 前端拆分 | HTML/CSS/JS 从 Python 字符串抽为 `static/` 独立文件 |
| ✅ 启动脚本升级 | 三阶段诊断：文件存在 → 可执行 → 依赖检查+自动安装 |
| ✅ 电信共享字段修复 | "是否共享"列名匹配 + Excel 原值为准的导入逻辑 |
| ✅ UI 优化 | 文件卡片单行横向滚动、分页页码跳转输入框 |

### 可优化方向

| 优化 | 说明 |
|------|------|
| 异步导入 | 用后台线程+WebSocket 推送进度，避免轮询 |
| 数据库 | 当前 records[] 全在内存，大数据量时可换 sqlite |
| 搜索索引 | 当前是 O(n) 全表扫描，可加倒排索引 |
| 室分站支持 | 扩展命名规则或新增规则覆盖电信室分站 |
| 移动支持 | 当前仅支持电信/联通，可扩展移动工参 |

### 修改注意事项

- **全局 `records` 和 `file_sources` 操作必须持有 `db_lock`**
- **字段映射修改**：改 `MAPPING_RULES` 列表，注意多候选列名按优先级排列
- **运营商检测修改**：改 `detect_carrier_from_names()`
- **前端修改**：编辑 `static/` 下的独立文件，后端无需重启即可刷新看效果（静态文件每次请求从磁盘读取）
- **静态文件安全**：`serve_static()` 的路径穿越检测（`os.path.normpath`）是关键安全防线，不可移除

---

## 十一、Git 信息

```
仓库: https://github.com/Xiaomo3250/CellSearchTool.git
当前版本: v0.2.1-alpha  (app.py 中 __version__ = "0.2.1-alpha")

版本历史:
  v0.3.0-alpha  → 报告生成器 WebUI 集成 — 页面内直接生成下载 docx (2026-05-05)
  v0.2.1-alpha  → 补全 _format_cell_record、报告工具前端入口
  v0.2.0-alpha  → 前后端分离、启动脚本升级、共享字段修复、UI 优化
  v0.1.2-alpha  → 运营商归属智能检测重构
  v0.1.1        → 设备商标签、列宽拖拽、点击复制、共享字段
  v0.1          → 首个完整版本
```

推送命令：
```bash
git push origin master --tags
```
