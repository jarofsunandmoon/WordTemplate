"""
基于 Word OOXML 的结构化渲染器。

职责分两层：
1. 对 document/header/footer 这些 XML 文件做 Jinja 渲染；
2. 按 document_structure.xml 的固定章节配置，替换正文段落和表格内容。
"""

import copy
import os
import secrets
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Optional

from jinja2 import Environment, FileSystemLoader, TemplateNotFound, select_autoescape


WORD_NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpsCustomData": "http://www.wps.cn/officeDocument/2013/wpsCustomData",
}

XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
WORD_NAMESPACE = WORD_NAMESPACES["w"]
W14_NAMESPACE = WORD_NAMESPACES["w14"]
NS = {"w": WORD_NAMESPACE}

for prefix, namespace in WORD_NAMESPACES.items():
    ET.register_namespace(prefix, namespace)


def w_tag(name: str) -> str:
    return f"{{{WORD_NAMESPACE}}}{name}"


class XMLWordGenerator:
    def __init__(self, base_docx_path: str, template_dir: str):
        self.base_docx_path = base_docx_path
        self.template_dir = template_dir
        self.env = Environment(
            loader=FileSystemLoader(self.template_dir),
            autoescape=select_autoescape(["xml"]),
        )
        self.document_structure = self._load_document_structure()

    def generate(self, context: dict, output_path: str, render_targets: Optional[list] = None):
        """
        :param context: 注入的数据字典
        :param output_path: 最终生成的文档路径
        :param render_targets: 指定 word 压缩包中需要渲染的 xml 文件列表
        """
        if render_targets is None:
            render_targets = [
                "word/document.xml",
                "word/header1.xml",
                "word/header2.xml",
                "word/footer1.xml",
            ]

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        print("正在打包并渲染文档组件...")

        with zipfile.ZipFile(self.base_docx_path, "r") as zip_read:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_write:
                for item in zip_read.infolist():
                    source_bytes = zip_read.read(item.filename)
                    if item.filename in render_targets:
                        try:
                            rendered_xml = self._render_target(
                                target_name=item.filename,
                                source_bytes=source_bytes,
                                context=context,
                            )
                            zip_write.writestr(item.filename, rendered_xml)
                            print(f"  [OK] 成功渲染: {item.filename}")
                        except Exception as exc:
                            print(f"  [WARN] 渲染 {item.filename} 失败 ({exc})，保留原始内容。")
                            zip_write.writestr(item, source_bytes)
                    else:
                        zip_write.writestr(item, source_bytes)

        print(f"\n文档已生成: {output_path}")

    def _render_target(self, target_name: str, source_bytes: bytes, context: dict) -> bytes:
        source_xml = source_bytes.decode("utf-8")

        if target_name == "word/document.xml":
            rendered_xml = self._render_template_or_fallback("document.xml", context, source_xml)
            root = ET.fromstring(rendered_xml)
            self._replace_section_content(root, context)
            self._remove_table_of_contents(root)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        if target_name == "word/settings.xml":
            root = ET.fromstring(source_xml)
            self._remove_update_fields(root)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        template_name = os.path.basename(target_name)
        rendered_xml = self._render_template_or_fallback(template_name, context, source_xml)
        return rendered_xml.encode("utf-8")

    def _render_template_or_fallback(self, template_name: str, context: dict, fallback: str) -> str:
        try:
            template_path = os.path.join(self.template_dir, template_name)
            if not os.path.exists(template_path) or os.path.getsize(template_path) == 0:
                return fallback
            template = self.env.get_template(template_name)
            return template.render(context)
        except TemplateNotFound:
            return fallback

    def _load_document_structure(self) -> list[dict[str, str]]:
        structure_path = os.path.join(self.template_dir, "document_structure.xml")
        if not os.path.exists(structure_path):
            return []

        root = ET.parse(structure_path).getroot()
        sections = []
        heading_counter: dict[str, int] = {}
        for section in root.findall("section"):
            section_id = section.get("id", "").strip()
            heading = section.get("heading", "").strip()
            level = section.get("level", "2").strip() or "2"
            if section_id and heading:
                heading_counter[heading] = heading_counter.get(heading, 0) + 1
                sections.append(
                    {
                        "id": section_id,
                        "heading": heading,
                        "level": level,
                        "occurrence": str(heading_counter[heading]),
                    }
                )
        return sections

    def _replace_section_content(self, root: ET.Element, context: dict):
        body = root.find("w:body", NS)
        if body is None or not self.document_structure:
            return

        self._ensure_section_skeletons(body)

        body_children = list(body)
        fallback_paragraph_template = self._pick_global_paragraph_template(body_children)
        sections_data = context.get("sections", {})
        tables_data = context.get("tables", {})

        for section_index in range(len(self.document_structure) - 1, -1, -1):
            section_spec = self.document_structure[section_index]
            section_payload = sections_data.get(section_spec["id"])
            if not section_payload:
                continue

            body_children = list(body)
            start_index = self._find_heading_index(
                body_children,
                section_spec["heading"],
                int(section_spec.get("occurrence", "1")),
            )
            if start_index is None:
                continue

            end_index = self._find_section_end_index(body_children, section_index, start_index)
            original_elements = body_children[start_index + 1:end_index]
            replacement_elements = self._build_section_elements(
                section_payload=section_payload,
                original_elements=original_elements,
                tables_data=tables_data,
                fallback_paragraph_template=fallback_paragraph_template,
            )

            for element in original_elements:
                body.remove(element)

            insert_at = list(body).index(body_children[start_index]) + 1
            for element in replacement_elements:
                body.insert(insert_at, element)
                insert_at += 1

    def _ensure_section_skeletons(self, body: ET.Element):
        for section_index, section_spec in enumerate(self.document_structure):
            body_children = list(body)
            start_index = self._find_heading_index(
                body_children,
                section_spec["heading"],
                int(section_spec.get("occurrence", "1")),
            )
            if start_index is not None:
                continue

            insert_at = self._find_insert_index_for_section(body_children, section_index)
            heading_template = self._pick_heading_template(body_children, section_spec.get("level", "2"))
            if heading_template is None:
                continue

            new_heading = copy.deepcopy(heading_template)
            self._sanitize_cloned_element(new_heading)
            self._set_container_text(new_heading, section_spec["heading"])
            body.insert(insert_at, new_heading)

    def _find_insert_index_for_section(self, body_children: list[ET.Element], structure_index: int) -> int:
        for next_spec in self.document_structure[structure_index + 1:]:
            next_index = self._find_heading_index(
                body_children,
                next_spec["heading"],
                int(next_spec.get("occurrence", "1")),
            )
            if next_index is not None:
                return next_index

        for index, child in enumerate(body_children):
            if self._tag_name(child) == "sectPr":
                return index

        return len(body_children)

    def _pick_heading_template(
        self,
        body_children: list[ET.Element],
        level: str,
    ) -> Optional[ET.Element]:
        fallback = None
        for section_spec in self.document_structure:
            start_index = self._find_heading_index(
                body_children,
                section_spec["heading"],
                int(section_spec.get("occurrence", "1")),
            )
            if start_index is None:
                continue

            if fallback is None:
                fallback = body_children[start_index]

            if section_spec.get("level") == level:
                return body_children[start_index]

        return fallback

    def _pick_global_paragraph_template(
        self,
        body_children: list[ET.Element],
    ) -> Optional[ET.Element]:
        for section_index, section_spec in enumerate(self.document_structure):
            start_index = self._find_heading_index(
                body_children,
                section_spec["heading"],
                int(section_spec.get("occurrence", "1")),
            )
            if start_index is None:
                continue

            end_index = self._find_section_end_index(body_children, section_index, start_index)
            template = self._pick_paragraph_template(body_children[start_index + 1:end_index])
            if template is not None:
                return template

        return None

    def _remove_table_of_contents(self, root: ET.Element):
        body = root.find("w:body", NS)
        if body is None:
            return

        body_children = list(body)
        toc_title_index = None
        for index, child in enumerate(body_children):
            if self._tag_name(child) != "p":
                continue
            if self._paragraph_contains_text(child, "目录"):
                toc_title_index = index
                break

        if toc_title_index is None:
            return

        self._trim_toc_title_paragraph(body_children[toc_title_index])

        first_heading_index = None
        for index in range(toc_title_index + 1, len(body_children)):
            child = body_children[index]
            if self._is_toc_anchor_heading(child):
                first_heading_index = index
                break

        if first_heading_index is None:
            return

        for element in body_children[toc_title_index + 1:first_heading_index]:
            body.remove(element)

    @staticmethod
    def _remove_update_fields(root: ET.Element):
        for update_fields in root.findall("w:updateFields", NS):
            root.remove(update_fields)

    @staticmethod
    def _paragraph_contains_text(paragraph: ET.Element, text: str) -> bool:
        return any((text_node.text or "").strip() == text for text_node in paragraph.findall(".//w:t", NS))

    def _trim_toc_title_paragraph(self, paragraph: ET.Element):
        for child in list(paragraph):
            if self._tag_name(child) != "r":
                continue

            has_page_break = child.find("w:br", NS) is not None
            is_toc_title_run = any((text_node.text or "").strip() == "目录" for text_node in child.findall(".//w:t", NS))
            if has_page_break or is_toc_title_run:
                paragraph.remove(child)

    def _is_toc_anchor_heading(self, paragraph: ET.Element) -> bool:
        if self._tag_name(paragraph) != "p":
            return False
        if paragraph.find(".//w:instrText", NS) is not None:
            return False

        for bookmark in paragraph.findall("w:bookmarkStart", NS):
            if bookmark.get(w_tag("name"), "").startswith("_Toc"):
                return True
        return False

    def _find_heading_index(
        self,
        body_children: list[ET.Element],
        heading_text: str,
        occurrence: int = 1,
    ) -> Optional[int]:
        match_count = 0
        for index, child in enumerate(body_children):
            if self._tag_name(child) != "p":
                continue
            if self._paragraph_text(child) == heading_text:
                match_count += 1
                if match_count < occurrence:
                    continue
                return index
        return None

    def _find_section_end_index(
        self,
        body_children: list[ET.Element],
        structure_index: int,
        start_index: int,
    ) -> int:
        for next_spec in self.document_structure[structure_index + 1:]:
            next_index = self._find_heading_index(
                body_children,
                next_spec["heading"],
                int(next_spec.get("occurrence", "1")),
            )
            if next_index is not None and next_index > start_index:
                return next_index

        for index, child in enumerate(body_children):
            if self._tag_name(child) == "sectPr":
                return index

        return len(body_children)

    def _build_section_elements(
        self,
        section_payload: Any,
        original_elements: list[ET.Element],
        tables_data: dict,
        fallback_paragraph_template: Optional[ET.Element] = None,
    ) -> list[ET.Element]:
        blocks = self._normalize_blocks(section_payload)
        if not blocks:
            return []

        paragraph_template = self._pick_paragraph_template(original_elements) or fallback_paragraph_template
        table_templates = [element for element in original_elements if self._tag_name(element) == "tbl"]

        rendered_elements = []
        for block in blocks:
            block_type = block.get("type", "paragraph")
            if block_type == "paragraph":
                rendered_elements.append(
                    self._build_paragraph_element(paragraph_template, str(block.get("text", "")))
                )
                continue

            if block_type == "table":
                table_key = block.get("table_key")
                if not table_key or table_key not in tables_data:
                    continue

                template_index = int(block.get("template_index", 0))
                if template_index >= len(table_templates):
                    continue

                rendered_elements.append(
                    self._build_table_element(table_templates[template_index], tables_data[table_key])
                )

            if block_type == "template_element":
                element_index = int(block.get("element_index", -1))
                if 0 <= element_index < len(original_elements):
                    rendered_elements.append(self._clone_original_element(original_elements[element_index]))

        return rendered_elements

    @staticmethod
    def _normalize_blocks(section_payload: Any) -> list[dict[str, Any]]:
        if isinstance(section_payload, list):
            if all(isinstance(item, str) for item in section_payload):
                return [{"type": "paragraph", "text": item} for item in section_payload]
            return [item for item in section_payload if isinstance(item, dict)]

        if isinstance(section_payload, dict):
            if "blocks" in section_payload:
                return XMLWordGenerator._normalize_blocks(section_payload["blocks"])
            if "paragraphs" in section_payload:
                return [
                    {"type": "paragraph", "text": item}
                    for item in section_payload.get("paragraphs", [])
                ]

        return []

    def _pick_paragraph_template(self, original_elements: list[ET.Element]) -> Optional[ET.Element]:
        for element in original_elements:
            if self._tag_name(element) != "p":
                continue
            if element.find(".//w:drawing", NS) is not None:
                continue
            return element
        return None

    def _build_paragraph_element(self, template: Optional[ET.Element], text: str) -> ET.Element:
        if template is None:
            paragraph = ET.Element(w_tag("p"))
            run = ET.SubElement(paragraph, w_tag("r"))
            text_node = ET.SubElement(run, w_tag("t"))
            text_node.text = text
            return paragraph

        paragraph = copy.deepcopy(template)
        self._sanitize_cloned_element(paragraph)
        self._set_container_text(paragraph, text)
        return paragraph

    def _clone_original_element(self, element: ET.Element) -> ET.Element:
        cloned = copy.deepcopy(element)
        self._sanitize_cloned_element(cloned)
        return cloned

    def _sanitize_cloned_element(self, element: ET.Element):
        removable_tags = {
            "bookmarkStart",
            "bookmarkEnd",
            "commentRangeStart",
            "commentRangeEnd",
            "commentReference",
            "proofErr",
            "permStart",
            "permEnd",
        }

        for parent in list(element.iter()):
            for child in list(parent):
                if self._tag_name(child) in removable_tags:
                    parent.remove(child)

        for paragraph in [element, *element.findall('.//w:p', NS)]:
            paragraph.set(f"{{{W14_NAMESPACE}}}paraId", secrets.token_hex(4).upper())
            paragraph.set(f"{{{W14_NAMESPACE}}}textId", secrets.token_hex(4).upper())

    def _build_table_element(self, template_table: ET.Element, table_payload: dict) -> ET.Element:
        table = copy.deepcopy(template_table)
        rows = table.findall("w:tr", NS)
        if not rows:
            return table

        header_row = rows[0]
        prototype_row = copy.deepcopy(rows[1] if len(rows) > 1 else rows[0])
        for row in rows[1:]:
            table.remove(row)

        columns = table_payload.get("columns", [])
        row_items = table_payload.get("rows", [])

        if not row_items:
            blank_row = copy.deepcopy(prototype_row)
            self._fill_table_row(blank_row, [""] * len(header_row.findall("w:tc", NS)))
            table.append(blank_row)
            return table

        for row_item in row_items:
            new_row = copy.deepcopy(prototype_row)
            values = self._row_values(row_item, columns)
            self._fill_table_row(new_row, values)
            table.append(new_row)

        return table

    @staticmethod
    def _row_values(row_item: Any, columns: list[str]) -> list[str]:
        if isinstance(row_item, dict):
            if columns:
                return [str(row_item.get(column, "")) for column in columns]
            return [str(value) for value in row_item.values()]

        if isinstance(row_item, (list, tuple)):
            return [str(value) for value in row_item]

        return [str(row_item)]

    def _fill_table_row(self, row: ET.Element, values: list[str]):
        for index, cell in enumerate(row.findall("w:tc", NS)):
            cell_value = values[index] if index < len(values) else ""
            paragraphs = cell.findall("w:p", NS)
            if not paragraphs:
                paragraph = ET.SubElement(cell, w_tag("p"))
                run = ET.SubElement(paragraph, w_tag("r"))
                ET.SubElement(run, w_tag("t"))
                paragraphs = [paragraph]

            self._set_container_text(paragraphs[0], cell_value)
            for paragraph in paragraphs[1:]:
                self._set_container_text(paragraph, "")

    def _set_container_text(self, container: ET.Element, text: str):
        text_nodes = container.findall(".//w:t", NS)
        if not text_nodes:
            run = container.find("w:r", NS)
            if run is None:
                run = ET.SubElement(container, w_tag("r"))
            text_nodes = [ET.SubElement(run, w_tag("t"))]

        text_nodes[0].text = text
        if text.startswith(" ") or text.endswith(" "):
            text_nodes[0].set(f"{{{XML_NAMESPACE}}}space", "preserve")
        else:
            text_nodes[0].attrib.pop(f"{{{XML_NAMESPACE}}}space", None)

        for extra_text in text_nodes[1:]:
            extra_text.text = ""
            extra_text.attrib.pop(f"{{{XML_NAMESPACE}}}space", None)

    @staticmethod
    def _paragraph_text(paragraph: ET.Element) -> str:
        return "".join(text_node.text or "" for text_node in paragraph.findall(".//w:t", NS)).strip()

    @staticmethod
    def _tag_name(element: ET.Element) -> str:
        return element.tag.rsplit("}", 1)[-1]
