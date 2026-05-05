"""
生成道路质量分析优化报告模板（docx）。
用法: python generate_report_template.py <输出路径> --area "<区域名>"
"""

import sys
import argparse
from pathlib import Path

# 将 scripts 目录加入路径
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

try:
    from create_docx import create_document
except ImportError:
    print("错误：无法导入 create_docx 模块", file=sys.stderr)
    sys.exit(1)


def generate_report(output_path: str, area: str = "XXX区域"):
    """生成完整的报告模板"""

    content = [
        # ===== 封面 =====
        {"type": "heading", "level": 0, "text": f"{area}道路质量分析优化报告"},
        {"type": "paragraph", "text": "2026年X月"},
        {"type": "paragraph", "text": "华为技术服务有限公司"},
        {"type": "page_break"},

        # ===== 一、道路测试优化概述 =====
        {"type": "heading", "level": 1, "text": "一、道路测试优化概述"},

        {"type": "heading", "level": 2, "text": "测试目的"},
        {"type": "paragraph", "text": "通过本轮遍历性测试，利用Assistant平台分析优化，远近协同，旨在快速提升目标区域路测指标，减少网内干扰，夯实网络质量，提升用户体验。"},

        {"type": "heading", "level": 2, "text": "测试工具"},
        {"type": "paragraph", "text": "测试工具如下表："},
        {"type": "paragraph", "text": "（此处插入测试工具表）"},

        {"type": "heading", "level": 2, "text": "测试路线"},
        {"type": "paragraph", "text": "（此处描述测试路线或插入路线图）"},

        # ===== 二、道路优化指标 =====
        {"type": "heading", "level": 1, "text": "二、道路优化指标"},

        {"type": "heading", "level": 2, "text": "指标情况"},
        {"type": "paragraph", "text": "本次测试的主要指标如下："},
        {"type": "paragraph", "text": "（此处插入指标汇总表）"},

        {"type": "heading", "level": 2, "text": "测试覆盖图"},
        {"type": "paragraph", "text": f"{area}测试分布图："},
        {"type": "paragraph", "text": "（此处插入测试轨迹分布图）"},
        {"type": "paragraph", "text": "RSRP分布图："},
        {"type": "paragraph", "text": "（此处插入RSRP分布图）"},
        {"type": "paragraph", "text": "SINR分布图："},
        {"type": "paragraph", "text": "（此处插入SINR分布图）"},

        # ===== 三、问题原因分类汇总 =====
        {"type": "heading", "level": 1, "text": "三、问题原因分类汇总"},
        {"type": "paragraph", "text": "（此处插入问题分类饼图/柱状图，按覆盖类、切换类、质差类进行统计）"},

        # ===== 四、问题点分析 =====
        {"type": "heading", "level": 1, "text": "四、问题点分析"},
        {"type": "paragraph", "text": "（★ 核心区域 — 使用以下话术模板逐一填写各问题点）"},

        # --- 覆盖类 ---
        {"type": "heading", "level": 2, "text": "4.1 问题点分析（覆盖类）"},

        {"type": "heading", "level": 3, "text": f"{area}道路4G弱覆盖：【位置】道路4G弱覆盖1"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"测试车辆行驶至【具体位置】，UE占用小区【小区名称】时，存在4G弱覆盖问题，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【小区名称】，RSRP【值】，【根因分析】，导致此路段4G弱覆盖。"},
        {"type": "paragraph", "text": "【优化方案】"},
        {"type": "number", "text": "调整【小区名称】下倾角【原值】°->【目标值】°。"},
        {"type": "number", "text": "调整【小区名称】方位角【原值】°->【目标值】°。"},
        {"type": "number", "text": "增加【小区名称】功率。"},

        {"type": "heading", "level": 3, "text": f"{area}道路4G弱覆盖：【位置】道路4G弱覆盖2"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"测试车辆行驶至【具体位置】，UE占用小区【小区名称】时，存在4G弱覆盖问题，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【服务小区】，RSRP【值】，较近小区【邻近小区】对当前道路覆盖较弱，RSRP【值】，导致此路段4G弱覆盖。"},
        {"type": "paragraph", "text": "【优化方案】"},
        {"type": "number", "text": "调整【邻近小区】下倾角【原值】°->【目标值】°。"},

        {"type": "heading", "level": 3, "text": f"{area}道路4G弱覆盖/无覆盖：【位置】道路4G弱覆盖/无覆盖1"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"测试车辆行驶至【具体位置】，UE占用小区【小区名称】时，存在4G弱覆盖/无覆盖问题，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【小区名称】(RSRP:【值】)时，【该路段/服务小区与此路段】存在【山体/建筑物】阻挡，导致此路段4G弱覆盖/无覆盖。"},
        {"type": "paragraph", "text": "【优化方案】"},
        {"type": "paragraph", "text": "方案一：于经纬度: 【lng】, 【lat】 建设基站【基站名】拉远【小区名】(EARFCN:【值】)，小区方位角【值】°，下倾角【值】°。"},
        {"type": "paragraph", "text": "方案二：若上述方案较难推动，可尝试于经纬度: 【lng】, 【lat】新建基站……但由于【原因】，建议谨慎评估此方案。"},

        # --- 切换类 ---
        {"type": "heading", "level": 2, "text": "4.2 问题点分析（切换类）"},

        {"type": "heading", "level": 3, "text": f"{area}道路4G弱覆盖：【位置】道路4G弱覆盖1"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"测试车辆行驶至【具体位置】，UE占用小区【服务小区】时，存在4G弱覆盖问题，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【服务小区】，未切换至较近且覆盖更好小区【目标小区】，进而导致弱覆盖。"},
        {"type": "paragraph", "text": "【优化方案】"},
        {"type": "number", "text": "添加小区【服务小区】到小区【目标小区】双向邻区关系。"},
        {"type": "number", "text": "降低小区【服务小区】的A2门限。"},

        # --- 质差类 ---
        {"type": "heading", "level": 2, "text": "4.3 问题点分析（质差类）"},

        {"type": "heading", "level": 3, "text": f"{area}道路4G质差：【位置】道路4G质差1"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"行驶至【具体位置】时，UE占用【运营商】小区【小区名称】，存在4G质差，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【小区名称】(距离:【值】,RSRP:【值】,SINR:【值】)时，近点小区【邻近小区】弱覆盖，此路段无主覆盖小区，导致该路段4G质差。"},
        {"type": "paragraph", "text": "【解决方案】"},
        {"type": "number", "text": "调整小区【邻近小区】方位角【原值】°->【目标值】°。"},
        {"type": "number", "text": "调整小区【邻近小区】下倾角【原值】°->【目标值】°。"},

        # --- 同位置质差变体 ---
        {"type": "heading", "level": 3, "text": f"{area}道路4G质差：【位置】道路4G质差2"},
        {"type": "paragraph", "text": "【问题描述】"},
        {"type": "paragraph", "text": f"行驶至【具体位置】时，UE占用小区【小区名称】，存在4G质差，问题路段持续约【距离】。"},
        {"type": "paragraph", "text": "【问题分析】"},
        {"type": "paragraph", "text": "UE占用小区【服务小区】(RSRP:【值】)时，RSRP与近点小区【邻近小区】(RSRP:【值】)相近，重叠覆盖导致质差。"},
        {"type": "paragraph", "text": "【解决方案】"},
        {"type": "number", "text": "调整小区【服务小区】下倾角【原值】°->【目标值】°。"},

        # ===== 五、测试总结 =====
        {"type": "heading", "level": 1, "text": "五、测试总结"},
        {"type": "paragraph", "text": "本轮测试LTE覆盖率（（-105&-3）*占用时长占比）【值】%，平均RSRP为【值】dBm，平均SINR为【值】dB。"},
        {"type": "paragraph", "text": "分析问题点【N】个，其中弱覆盖问题【n1】个，切换问题【n2】个，质差问题【n3】个，需推动处理后进行下一轮测试。"},
    ]

    tables = [
        {
            "caption": "表1：测试工具清单",
            "headers": ["工具名称", "用途"],
            "rows": [
                ["（工具1）", "（用途说明）"],
                ["（工具2）", "（用途说明）"],
            ],
        },
        {
            "caption": "表2：关键指标汇总",
            "headers": ["指标项", "LTE", "NR", "备注"],
            "rows": [
                ["覆盖率(-105&-3)", "（值）%", "（值）%", ""],
                ["平均RSRP/SS-RSRP", "（值）dBm", "（值）dBm", ""],
                ["平均SINR/SS-SINR", "（值）dB", "（值）dB", ""],
                ["问题点总数", "—", "—", "（值）个"],
            ],
        },
    ]

    create_document(output_path, title=None, content=content, tables=tables)
    print(f"✅ 报告模板已生成: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="生成道路质量分析优化报告模板")
    parser.add_argument("output", help="输出 docx 文件路径")
    parser.add_argument("--area", default="XXX区域", help="区域名称（如：阿乌高速、天山天池）")
    args = parser.parse_args()

    generate_report(args.output, args.area)


if __name__ == "__main__":
    main()
