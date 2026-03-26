"""
负责组装结构化上下文数据。

当前模板按两类来源预留数据入口：
1. 前端用户输入：平台基础信息、改造范围、文字结论、施工方案等；
2. SACS 输入文件（如 .inp / .factor）：环境参数、分析工况、校核结果、附录原始输出等。
"""

import os
import xml.etree.ElementTree as ET
from typing import Optional

from src.inp_reader import SACSInpReader


def paragraph(text: str) -> dict:
    return {"type": "paragraph", "text": text}


def table_block(table_key: str, template_index: int = 0) -> dict:
    return {"type": "table", "table_key": table_key, "template_index": template_index}


def template_element_block(element_index: int) -> dict:
    return {"type": "template_element", "element_index": element_index}


FRONTEND_SECTION_IDS = {
    "overview",
    "assessment_overview",
    "platform_overview_history",
    "assessment_basis",
    "assessment_process",
    "basic_data",
    "deck_loads",
    "construction_plan",
    "work_scope",
    "underwater_installation",
    "above_water_installation",
    "conclusion_recommendation",
    "conclusions",
    "recommendations",
    "appendix_b",
}

SACS_SECTION_IDS = {
    "environment_conditions",
    "water_depth",
    "wave",
    "current",
    "wind",
    "marine_growth",
    "splash_zone",
    "scour",
    "analysis_model",
    "program_coordinate",
    "structure_modeling",
    "pile_modeling",
    "design_level_analysis",
    "environment_load_parameters",
    "dynamic_response",
    "basic_load_cases",
    "load_combinations",
    "analysis_results",
    "nominal_stress_check",
    "effective_length_factor",
    "joint_shear_check",
    "pile_check_results",
    "ultimate_strength_analysis",
    "ultimate_strength_overview",
    "ultimate_strength_theory",
    "deck_wave_analysis",
    "ultimate_environment_loads",
    "ultimate_results",
    "fatigue_analysis",
    "fatigue_overview",
    "fatigue_program",
    "fatigue_model",
    "natural_period_modes",
    "transfer_function",
    "fatigue_results",
    "appendix_a",
    "appendix_c",
}

TEMPLATE_FIXED_SECTION_IDS = {
    "code_basis",
    "assessment_process",
}


class ReportDataLoader:
    def __init__(
        self,
        sacs_source_path: Optional[str] = None,
        structure_path: Optional[str] = None,
        frontend_payload: Optional[dict] = None,
    ):
        self.sacs_source_path = sacs_source_path
        self.structure_path = structure_path
        self.frontend_payload = frontend_payload or {}

    def get_context_data(self) -> dict:
        frontend_data = self._build_frontend_defaults()
        sacs_data = self._build_sacs_context()

        context = {
            "platform_name": frontend_data["platform_name"],
            "oil_name": frontend_data["oil_name"],
            "report_date": frontend_data["report_date"],
            "report_code": frontend_data["report_code"],
            "footer_title": frontend_data["footer_title"],
            "frontend": frontend_data,
            "sacs": sacs_data,
            "data_sources": {
                "frontend": "前端录入",
                "sacs": sacs_data["source_status"],
            },
            "sections": self._build_sections(frontend_data, sacs_data),
            "tables": self._build_tables(frontend_data, sacs_data),
        }
        return context

    def _build_frontend_defaults(self) -> dict:
        platform_name = "文昌8-3A"
        oil_name = "文昌19-1油田"
        project_name = "结构强度改造"
        default_data = {
            "platform_name": platform_name,
            "field_name": "文昌油田群",
            "oil_name":oil_name,
            "project_name": project_name,
            "report_date": "2026年03月10日",
            "report_code": "FS-WEN8-3A-RPT-ST-001",
            "footer_title": f"{platform_name}平台改建可行性评估",
            "location": "南海东部海域",
            "water_depth": "112.2m",
            "retrofit_scope": "新增井槽、局部甲板加固、立管及附属构件复核",
            "assessment_scope": "结构强度、桩基承载、极限强度、疲劳寿命及施工适应性",
            "design_basis": "API RP 2A / API RP 2SIM / 项目改造任务书",
            "construction_window": "计划结合年度检修窗口实施",
            "history_rows": [
                {"index": "1", "project_name": f"{platform_name}增加调整井", "year": "2009年"},
                {"index": "2", "project_name": f"{oil_name}产能释放", "year": "2016年"},
                {"index": "3", "project_name": f"{oil_name}油田开发工程详细设计项目旧平台改造结构", "year": "2011年"},
            ],
            "basic_data_files": [
                {
                    "category": "平台原设计文件",
                    "file_name": "WC8-3A-STR-001 平台原设计总图.pdf",
                    "version": "V1.0",
                    "source": "历史设计档案",
                    "upload_date": "2026-03-10",
                    "remark": "用于确认平台原始结构布置和杆件尺寸",
                },
                {
                    "category": "建造、安装完工文件",
                    "file_name": "WC8-3A-AsBuilt 完工报告.pdf",
                    "version": "V2.1",
                    "source": "竣工资料库",
                    "upload_date": "2026-03-10",
                    "remark": "用于核对建造完成后的实际安装状态",
                },
                {
                    "category": "基线检测报告",
                    "file_name": "WC8-3A-Inspection 基线检测报告.pdf",
                    "version": "V1.3",
                    "source": "检测资料库",
                    "upload_date": "2026-03-11",
                    "remark": "用于识别结构现状、腐蚀和缺陷信息",
                },
                {
                    "category": "历次改造设计文件及安装报告",
                    "file_name": "WC19-1 改造设计与安装报告.zip",
                    "version": "V3.0",
                    "source": "改造项目归档",
                    "upload_date": "2026-03-12",
                    "remark": "用于追踪历史改造边界和新增构件信息",
                },
                {
                    "category": "平台海生物清除及五年特检报告",
                    "file_name": "WC8-3A 2024五年特检报告.pdf",
                    "version": "V1.0",
                    "source": "特检资料库",
                    "upload_date": "2026-03-12",
                    "remark": "用于支撑飞溅区、水下结构和腐蚀状态评估",
                },
            ],
        }

        frontend_data = {**default_data, **self.frontend_payload}
        history_rows = self.frontend_payload.get("history_rows")
        basic_data_files = self.frontend_payload.get("basic_data_files")
        frontend_data["history_rows"] = history_rows if history_rows is not None else default_data["history_rows"]
        frontend_data["basic_data_files"] = self._normalize_basic_data_files(
            basic_data_files if basic_data_files is not None else default_data["basic_data_files"]
        )
        frontend_data["footer_title"] = self.frontend_payload.get(
            "footer_title",
            f"{frontend_data['platform_name']}平台改建可行性评估",
        )
        frontend_data["basic_data_interface"] = {
            "field": "basic_data_files",
            "source": "PyQt5 文件管理系统",
            "status": "已接入前端数据" if "basic_data_files" in self.frontend_payload else "已预留接口，当前使用写死展示数据",
        }
        return frontend_data

    def _build_sacs_context(self) -> dict:
        source_file = os.path.basename(self.sacs_source_path) if self.sacs_source_path else "待接入"
        source_status = (
            f"已指定待接入文件：{source_file}"
            if self.sacs_source_path
            else "尚未接入 SACS 输入文件，当前使用模板占位值"
        )

        data = {
            "source_file": source_file,
            "source_status": source_status,
            "model_name": "WEN8-3A_INPLACE",
            "analysis_standard": "SACS inplace / fatigue / collapse 模块",
            "storm_condition": "100年一遇风浪流组合",
            "operating_condition": "操作工况 + 改建附加载荷",
            "max_member_uc": "待导入",
            "max_joint_uc": "待导入",
            "min_pile_safety_factor": "待导入",
            "min_rsr": "待导入",
            "min_fatigue_life": "待导入",
            "notes": "后续可由 SACS 文件解析程序直接覆盖本占位数据。",
            "encoding": "未读取",
            "total_lines": "0",
            "non_empty_lines": "0",
            "basic_load_case_count": "0",
            "basic_load_case_lines": [],
            "preview_lines": [],
            "top_cards": [],
        }

        if self.sacs_source_path:
            data.update(self._parse_sacs_file())

        return data

    def _build_sections(self, frontend_data: dict, sacs_data: dict) -> dict:
        # 主要章节生成
        sections = {
            "assessment_overview": {
                "blocks": [
                    paragraph(
                        f"{frontend_data['platform_name']}平台平台于2006年安装并投产，平台原设计寿命15年，在投产后平台进行了一系列的改造，其中较大的改造包括xxx油田开发工程项目依托电缆护管和立管、增加结构房间、2016年增加救生筏和逃生软梯、"
                        f"{frontend_data['oil_name']}产能释放项目等。"
                        f"评估范围覆盖{frontend_data['assessment_scope']}。"
                    ),
                    paragraph(
                        "2018年五年特检报告显示，水下导管架杆件结构完整、未发现凹陷变形与机械损伤，牺牲阳极均在位，连接牢固；焊缝检测未发现缺陷，飞溅区构件测厚结果最小腐蚀厚度0.1mm，最大腐蚀厚度0.5mm，平台水下结构状况良好。"
                    ),
                    paragraph(
                        "根据平台的空间布置、改造工作量和结构强度等方面的综合评估，在原井口区南侧新增两根914mm的隔水套管，利用平台钻机可打6口井。"
                    ),
                    paragraph(
                        "对平台增加隔水套管后的整体结构进行设计水平强度分析，结果显示所有杆件	UC值小于1.0，桩基承载力安全系数>1.5，满足规范要求；极限强度分析结果显示最小RSR为2.1，满足规范要求；疲劳分析结果显示节点寿命均满足要求。综合以上结果，认为增加隔水套管可行。"
                    ),
                ]
            },
            "platform_overview_history": {
                "blocks": [
                    paragraph(
                        f"{frontend_data['platform_name']}平台位于{frontend_data['location']}，平台坐标：东经	624 392 .900mE；北纬2 185 633.600mN；平台方位：平台北为真北偏东45。"
                    ),
                    paragraph(
                        f"{frontend_data['platform_name']}平台原有4个腿、8根裙桩、1根立管和2根电缆护管及2根泵护管、6个井口，后增加2根立管和1根电缆护管。"
                        f"平台桩径2134mm直径，入泥87米。导管架工作点标高为EL(+)8.5m ，工作点尺寸18X12米，设置有7个水平层，分别为：EL(+)7m, EL(-)10m，EL(-)30m，EL(-)52m，EL(-)77m，EL(-)102.2m，EL(-)112.2m。"
                        f"导管架水平层自上而下从第一至六层设有隔水套管导向和传递隔水套管水平力的构件，泥面处设置防沉板。平台飞溅区范围为EL.(-)3.6米至EL.(+)8.6米，飞溅区内的竖向杆件（包括斜向杆件）考虑了6毫米腐蚀余量，水平杆件考虑了4.5毫米腐蚀余量。"
                    ),
                    paragraph(
                        "平台上部组块主要有三层甲板，上甲板标高EL.(+) 31 000，中层甲板标高EL.(+) 23 000，下甲板标高EL.(+)18 000。上甲板上布置有修井机，设管子堆场，EL.(+) 36 000处设置有直升机甲板。"
                    ),
                    paragraph(
                        "平台历次改造清单如下："
                    ),
                    table_block("history"),
                ]
            },
            "risk_level": {
                "blocks": [
                    paragraph(
                        f"根据API RP 2SIM，{frontend_data['platform_name']}平台为有人可撤离，生命安全分级为S-2；井口平台，水深接近120米，没有储油设施，设置有井下安全阀，失效后果定义为C-1；因此暴露等级为L-1。"
                    ),
                    template_element_block(1),
                    template_element_block(2),
                    template_element_block(3),
                    template_element_block(4),
                    template_element_block(5),
                ]
            },
            "basic_data": {
                "blocks": self._build_basic_data_blocks(frontend_data)
            },
            "deck_loads": {
                "blocks": [
                    paragraph(
                        "甲板荷载章节用于统一维护现状荷载、改建新增荷载和施工阶段临时荷载，是前端录入与分析输入之间的重要衔接层。"
                    ),
                    paragraph(
                        "后续可根据前端表单直接生成该表，并同步写入设计水平分析与极限强度分析的荷载工况。"
                    ),
                    table_block("deck_load_summary"),
                ]
            },
            "water_depth": {
                "blocks": [
                    paragraph(
                        f"现阶段模板按项目基础信息暂记水深为 {frontend_data['water_depth']}，正式值建议以后续 SACS 文件读取的分析水深为准。"
                    ),
                    paragraph(
                        "设计高水位、低水位和天文潮位等参数可在 SACS 文件解析完成后自动写入正文，同时保留前端手工修订能力。"
                    ),
                ]
            },
            "wave": {
                "blocks": [
                    paragraph("波浪参数章节预留给 1 年、10 年、100 年等代表性重现期的波高和周期输入。"),
                    paragraph("当前表格使用模板占位数据演示写入方式，后续可直接替换为 SACS 文件解析结果。"),
                    table_block("wave_parameters"),
                ]
            },
            "current": {
                "blocks": [
                    paragraph("海流章节用于汇总不同重现期或不同水深位置的流速取值，并为环境组合提供输入。"),
                    paragraph("若后续 SACS 输入文件中包含多组工况，可按统一格式扩展为多行自动输出。"),
                    table_block("current_profile"),
                ]
            },
            "wind": {
                "blocks": [
                    paragraph("风参数章节承接平台所在海域的设计风速与方向信息，用于设计水平和极限强度分析。"),
                    paragraph("当前先保留模板占位行，待 SACS 文件接入后可自动写入各重现期参考风速。"),
                    table_block("wind_parameters"),
                ]
            },
            "dynamic_response": {
                "blocks": [
                    paragraph("结构动力响应章节预留用于展示结构周期、主振型和动力放大系数。"),
                    paragraph("该部分通常由 SACS 动力分析结果直接生成，当前表格仅用于验证模板联通性。"),
                    table_block("dynamic_response_summary"),
                ]
            },
            "basic_load_cases": {
                "blocks": self._build_basic_load_case_blocks(sacs_data)
            },
            "analysis_model": {
                "blocks": [
                    paragraph(
                        f"当前分析输入文件状态：{sacs_data['source_status']}；文件编码为 {sacs_data['encoding']}。"
                    ),
                    paragraph(
                        f"已读取总行数 {sacs_data['total_lines']}，非空行 {sacs_data['non_empty_lines']}；已提取基本载荷工况 {sacs_data['basic_load_case_count']} 条。"
                    ),
                    paragraph(self._top_cards_text(sacs_data)),
                ]
            },
            "appendix_c": {
                "blocks": [
                    paragraph("附录C 预留用于输出桩基数据、输入文件关键片段、模型参数和程序摘要。"),
                    paragraph(
                        f"当前状态：{sacs_data['source_status']}；已读取总行数 {sacs_data['total_lines']}，非空行 {sacs_data['non_empty_lines']}。"
                    ),
                    *[paragraph(line) for line in sacs_data.get("preview_lines", [])[:8]],
                ]
            },
            "conclusions": {
                "blocks": [
                    paragraph("结论章节用于汇总各校核模块的控制结果，并形成是否满足改建要求的最终判断。"),
                    paragraph(
                        f"在 SACS 文件尚未接入前，模板先保留结构化结论占位；后续可自动整合杆件利用系数、桩基安全系数、RSR 和疲劳寿命等核心指标。"
                    ),
                    paragraph("当前建议保持“前端补充工程判断 + 程序自动汇总控制值”的组合输出模式。"),
                ]
            },
            "recommendations": {
                "blocks": [
                    paragraph("建议章节面向后续详细设计、施工组织和现场实施阶段，用于承接工程师的专业建议。"),
                    paragraph("该部分默认由前端录入维护，但也可以在程序中根据关键校核结论自动补充标准化建议语句。"),
                ]
            },
            "appendix_a": {
                "blocks": [
                    paragraph("附录A 预留用于插入长期波浪分布图、环境统计图表或由 SACS 文件解析导出的环境输入摘要。"),
                    paragraph("当前模板已完成附录章节占位，后续可根据图片或数据文件自动扩展内容。"),
                ]
            },
            "appendix_b": {
                "blocks": [
                    paragraph("附录B 建议承接新增井槽改造方案、布置图说明、构件清单和施工示意。"),
                    paragraph("该附录以项目前端维护的图纸说明和方案描述为主，后续可追加图像或表格附件。"),
                ]
            },
        }

        for section in self._load_structure_sections():
            if section["id"] in sections or section["id"] in TEMPLATE_FIXED_SECTION_IDS:
                continue
            sections[section["id"]] = {"blocks": self._default_section_blocks(section["id"], section["heading"], sacs_data)}

        return sections

    # 对应表格生成
    def _build_tables(self, frontend_data: dict, sacs_data: dict) -> dict:
        return {
            "history": {
                "columns": ["index", "project_name", "year"],
                "rows": frontend_data["history_rows"],
            },
            "deck_load_summary": {
                "columns": ["item", "value"],
                "rows": [
                    {"item": "现状设备荷载", "value": "依据前端输入确认"},
                    {"item": "改造新增荷载", "value": "待前端录入"},
                    {"item": "施工临时荷载", "value": "待施工方案明确"},
                ],
            },
            "wave_parameters": {
                "columns": ["parameter", "one_year", "hundred_year"],
                "rows": [
                    {"parameter": "有效波高 Hs (m)", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "谱峰周期 Tp (s)", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "主导波向", "one_year": "待导入", "hundred_year": "待导入"},
                ],
            },
            "current_profile": {
                "columns": ["parameter", "one_year", "hundred_year"],
                "rows": [
                    {"parameter": "表层流速 (m/s)", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "中层流速 (m/s)", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "底层流速 (m/s)", "one_year": "待导入", "hundred_year": "待导入"},
                ],
            },
            "wind_parameters": {
                "columns": ["parameter", "one_year", "hundred_year"],
                "rows": [
                    {"parameter": "10m 参考风速 (m/s)", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "主导风向", "one_year": "待导入", "hundred_year": "待导入"},
                    {"parameter": "说明", "one_year": "操作工况", "hundred_year": sacs_data["storm_condition"]},
                ],
            },
            "dynamic_response_summary": {
                "columns": ["mode", "direction", "ts", "tw", "daf"],
                "rows": [
                    {"mode": "一阶", "direction": "X 向平动", "ts": "待导入", "tw": "待导入", "daf": "待导入"},
                    {"mode": "二阶", "direction": "Y 向平动", "ts": "待导入", "tw": "待导入", "daf": "待导入"},
                    {"mode": "三阶", "direction": "扭转", "ts": "待导入", "tw": "待导入", "daf": "待导入"},
                ],
            },
        }

    def _build_basic_data_blocks(self, frontend_data: dict) -> list[dict]:
        file_rows = frontend_data.get("basic_data_files", [])
        if not file_rows:
            return [
                paragraph("当前未从前端文件管理系统接收到基础数据文件，模板暂不展示资料清单。"),
            ]

        return [paragraph(self._format_basic_data_file_line(file_row)) for file_row in file_rows]

    @staticmethod
    def _normalize_basic_data_files(file_rows: Optional[list[dict]]) -> list[dict]:
        normalized_rows = []
        if not file_rows:
            return normalized_rows

        for index, file_row in enumerate(file_rows, start=1):
            normalized_rows.append(
                {
                    "index": str(file_row.get("index", index)),
                    "category": str(file_row.get("category", "基础资料")),
                    "file_name": str(file_row.get("file_name", "未命名文件")),
                    "version": str(file_row.get("version", "未标注")),
                    "source": str(file_row.get("source", "前端文件管理系统")),
                    "upload_date": str(file_row.get("upload_date", "待补充")),
                    "remark": str(file_row.get("remark", "待补充说明")),
                }
            )
        return normalized_rows

    @staticmethod
    def _format_basic_data_file_line(file_row: dict) -> str:
        return (
            f"{file_row['category']}：{file_row['file_name']}；版本 {file_row['version']}；"
            f"来源 {file_row['source']}；登记日期 {file_row['upload_date']}；说明：{file_row['remark']}。"
        )

    def _load_structure_sections(self) -> list[dict[str, str]]:
        if not self.structure_path or not os.path.exists(self.structure_path):
            return []

        root = ET.parse(self.structure_path).getroot()
        sections = []
        for section in root.findall("section"):
            section_id = section.get("id", "").strip()
            heading = section.get("heading", "").strip()
            if section_id and heading:
                sections.append({"id": section_id, "heading": heading})
        return sections

    def _default_section_blocks(self, section_id: str, heading: str, sacs_data: dict) -> list[dict]:
        return [
            paragraph(f"《{heading}》章节已完成模板占位，当前用于 {self._section_purpose(section_id)}。"),
            paragraph(self._section_source_note(section_id, sacs_data)),
        ]

    @staticmethod
    def _section_purpose(section_id: str) -> str:
        if section_id.startswith("ultimate_"):
            return "承接极限强度分析的理论说明、荷载边界和结果摘要"
        if section_id.startswith("fatigue_"):
            return "承接疲劳分析模型、传递函数和寿命评估结果"
        if section_id.startswith("appendix_"):
            return "作为图纸、程序输出和补充说明的扩展附件"

        purposes = {
            "overview": "说明评估目标、使用范围和报告边界",
            "assessment_overview": "概述本次改建评估的任务背景和目标",
            "platform_overview_history": "记录平台概况和历次改造背景",
            "risk_level": "明确平台风险分类和评价边界",
            "assessment_basis": "汇总评估依据、输入资料和适用标准",
            "code_basis": "整理规范、标准和项目专项要求",
            "assessment_process": "说明整体评估流程和工作路径",
            "basic_data": "整理基础资料、检测资料和设计输入",
            "deck_loads": "汇总甲板荷载和改造新增荷载",
            "environment_conditions": "说明环境输入参数的取值范围",
            "water_depth": "说明设计水深和设计水位条件",
            "wave": "汇总波浪参数及重现期取值",
            "current": "汇总海流参数及分层流速",
            "wind": "汇总风速、风向和风载边界",
            "marine_growth": "说明海生物附着参数",
            "splash_zone": "说明飞溅区和腐蚀余量设置",
            "scour": "说明基础冲刷假定",
            "analysis_model": "说明分析模型范围和使用原则",
            "program_coordinate": "说明分析程序和坐标系统定义",
            "structure_modeling": "说明结构建模原则和简化处理",
            "pile_modeling": "说明桩基础建模和土弹簧取值",
            "design_level_analysis": "承接设计水平分析的输入、工况和结果",
            "environment_load_parameters": "说明环境载荷参数和组合策略",
            "dynamic_response": "展示结构周期、振型和动力放大系数",
            "basic_load_cases": "列示基本荷载工况及命名规则",
            "load_combinations": "列示荷载组合及控制组合",
            "analysis_results": "总结设计水平分析结果",
            "nominal_stress_check": "输出杆件名义应力校核结果",
            "effective_length_factor": "说明压杆有效长度系数的取值依据",
            "joint_shear_check": "输出节点冲剪校核结果",
            "pile_check_results": "输出桩基承载校核结果",
            "construction_plan": "说明改造施工思路和实施路径",
            "work_scope": "汇总工作量和主要工程内容",
            "underwater_installation": "描述水下部分施工安排",
            "above_water_installation": "描述水上部分施工安排",
            "conclusion_recommendation": "汇总结论与建议的组织方式",
            "conclusions": "形成最终技术结论",
            "recommendations": "提出后续设计与实施建议",
        }
        return purposes.get(section_id, f"承接《{section_id}》相关内容")

    def _section_source_note(self, section_id: str, sacs_data: dict) -> str:
        if section_id in FRONTEND_SECTION_IDS:
            return "本节优先使用前端录入数据，当前模板已预留基础信息、范围说明和人工结论字段。"
        if section_id in SACS_SECTION_IDS:
            return (
                "本节主要承接 SACS 文件解析结果，后续将自动写入工况、控制值和结论；"
                f"当前状态：{sacs_data['source_status']}。"
            )
        return (
            "本节同时结合前端输入与 SACS 文件结果：文字说明可由前端维护，"
            f"计算结果待 SACS 文件接入后自动填充；当前状态：{sacs_data['source_status']}。"
        )

    @staticmethod
    def _top_cards_text(sacs_data: dict) -> str:
        top_cards = sacs_data.get("top_cards", [])
        if not top_cards:
            return "当前尚无 inp 卡片统计信息。"

        summary = "，".join(f"{item['card']} x {item['count']}" for item in top_cards[:8])
        return f"已识别的高频卡片包括：{summary}。"

    @staticmethod
    def _build_basic_load_case_blocks(sacs_data: dict) -> list[dict]:
        lines = sacs_data.get("basic_load_case_lines", [])
        if lines:
            return [paragraph(line) for line in lines]

        return [
            paragraph("当前尚未从 SACS 输入文件中提取到“基本载荷工况”数据。"),
            paragraph("目标字段为：LOAD  LOAD     ********** DESCRIPTION ***********。"),
        ]

    def _parse_sacs_file(self) -> dict:
        if not self.sacs_source_path:
            return {}

        print(f"准备解析 SACS 文件: {self.sacs_source_path}")
        summary = SACSInpReader(self.sacs_source_path).summarize()
        return {
            "source_file": summary["file_name"],
            "source_status": f"已读取 SACS 输入文件：{summary['file_name']}",
            "encoding": summary["encoding"],
            "total_lines": str(summary["total_lines"]),
            "non_empty_lines": str(summary["non_empty_lines"]),
            "basic_load_case_count": str(summary["basic_load_case_count"]),
            "basic_load_case_lines": summary["basic_load_case_lines"],
            "preview_lines": summary["preview_lines"],
            "top_cards": summary["top_cards"],
        }
