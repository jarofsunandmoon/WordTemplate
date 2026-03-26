"""
轻量级的“调度中心”，只负责路径配置和调用。
"""

import os

from src.generator import XMLWordGenerator
from src.data_loader import ReportDataLoader


TARGET_SACS_PREFIX = "psilst"
PREFERRED_SACS_FILENAMES = ("psilst.factor", "psilst.inp", "psilst")


def detect_sacs_input_path(base_dir: str) -> str | None:
    env_path = os.environ.get("SACS_INP_PATH", "").strip()
    if env_path:
        return env_path if os.path.exists(env_path) else None

    data_dir = os.path.join(base_dir, "data")
    if not os.path.isdir(data_dir):
        return None

    candidates = [
        entry.path
        for entry in os.scandir(data_dir)
        if entry.is_file() and entry.name.lower().startswith(TARGET_SACS_PREFIX)
    ]
    if not candidates:
        return None

    exact_matches = {
        os.path.basename(path).lower(): path
        for path in candidates
    }
    for file_name in PREFERRED_SACS_FILENAMES:
        if file_name in exact_matches:
            return exact_matches[file_name]

    return max(candidates, key=lambda path: (os.path.getmtime(path), os.path.basename(path).lower()))

def main():
    # 1. 绝对路径配置
    base_dir = os.path.dirname(os.path.abspath(__file__))
    base_docx = os.path.join(base_dir, 'templates', 'base.docx')
    template_dir = os.path.join(base_dir, 'templates')
    structure_path = os.path.join(template_dir, 'document_structure.xml')
    output_docx = os.path.join(base_dir, 'output', '报告_final.docx')
    sacs_inp_path = detect_sacs_input_path(base_dir)
    if sacs_inp_path:
        print(f"检测到 SACS 输入文件: {sacs_inp_path}")
    else:
        print("未检测到可读取的 SACS 输入文件，当前继续使用占位数据。")

    # 2. 提取数据：实例化加载器并获取拼装好的字典
    data_loader = ReportDataLoader(sacs_source_path=sacs_inp_path, structure_path=structure_path)
    context_data = data_loader.get_context_data()

    # 3. 执行生成
    generator = XMLWordGenerator(base_docx, template_dir)
    generator.generate(context=context_data, output_path=output_docx)

if __name__ == "__main__":
    main()
