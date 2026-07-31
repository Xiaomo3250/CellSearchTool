/// 基站工参数据模型
/// Dart 不支持中文标识符，属性用英文命名，数据库映射用中文列名
class Station {
  final int? id;

  // -- 核心工参字段 --
  final String tech;    // 技术制式 (4G LTE / 5G NR)
  final String carrier; // 运营商 (中国电信/中国联通)
  final String vendor;  // 设备商 (华为/大唐/中兴)
  final String siteName; // 基站名
  final String siteId;   // 基站ID (eNodeB/gNodeB)
  final String cellName; // 小区名
  final String cellId;   // 小区ID
  final String pci;      // 物理小区标识
  final String dlFreq;   // 下行频点 (EARFCN/SSB)
  final String tilt;     // 下倾角
  final String height;   // 挂高/站高
  final String azimuth;  // 方位角
  final String lng;      // 经度
  final String lat;      // 纬度
  final String share;    // 共享状态
  final String band;     // 频段
  final String tac;      // 跟踪区码

  // -- 元数据 --
  final String filename; // 来源文件名
  final String sheet;    // 来源工作表名
  final String raw;      // 原始行数据 JSON

  Station({
    this.id,
    this.tech = '',
    this.carrier = '',
    this.vendor = '',
    this.siteName = '',
    this.siteId = '',
    this.cellName = '',
    this.cellId = '',
    this.pci = '',
    this.dlFreq = '',
    this.tilt = '',
    this.height = '',
    this.azimuth = '',
    this.lng = '',
    this.lat = '',
    this.share = '',
    this.band = '',
    this.tac = '',
    this.filename = '',
    this.sheet = '',
    this.raw = '{}',
  });

  /// 从数据库 Map 构建（SQLite 列名为中文，取值时用字符串 key）
  factory Station.fromMap(Map<String, dynamic> map) => Station(
        id:       map['id'] as int?,
        tech:     map['技术制式'] as String? ?? '',
        carrier:  map['运营商']   as String? ?? '',
        vendor:   map['设备商']   as String? ?? '',
        siteName: map['基站名']   as String? ?? '',
        siteId:   map['基站ID']   as String? ?? '',
        cellName: map['小区名']   as String? ?? '',
        cellId:   map['小区ID']   as String? ?? '',
        pci:      map['PCI']      as String? ?? '',
        dlFreq:   map['下行频点'] as String? ?? '',
        tilt:     map['下倾角']   as String? ?? '',
        height:   map['挂高']     as String? ?? '',
        azimuth:  map['方位角']   as String? ?? '',
        lng:      map['经度']     as String? ?? '',
        lat:      map['纬度']     as String? ?? '',
        share:    map['共享']     as String? ?? '',
        band:     map['频段']     as String? ?? '',
        tac:      map['TAC']      as String? ?? '',
        filename: map['_文件名']  as String? ?? '',
        sheet:    map['_工作表']  as String? ?? '',
        raw:      map['_raw']     as String? ?? '{}',
      );

  /// 转为数据库 Map（中文列名，兼容 SQlite）
  Map<String, dynamic> toMap() => {
        '技术制式': tech,
        '运营商':   carrier,
        '设备商':   vendor,
        '基站名':   siteName,
        '基站ID':   siteId,
        '小区名':   cellName,
        '小区ID':   cellId,
        'PCI':      pci,
        '下行频点': dlFreq,
        '下倾角':   tilt,
        '挂高':     height,
        '方位角':   azimuth,
        '经度':     lng,
        '纬度':     lat,
        '共享':     share,
        '频段':     band,
        'TAC':      tac,
        '_文件名':  filename,
        '_工作表':  sheet,
        '_raw':     raw,
      };

  /// 制式简写 (4G / 5G)
  String get techLabel => tech.contains('5G') ? '5G' : '4G';
  /// 运营商短名（去"中国"前缀）
  String get shortCarrier => carrier.replaceAll('中国', '');
  /// 速查格式：制式_频点_基站ID_CellID (例: 4G_1600_620055_2)
  String get quickRef => '$techLabel\_$dlFreq\_$siteId\_$cellId';
}
