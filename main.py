"""DOCX文档信息提取工具。

使用OpenAI结构化输出提取文档中的特定字段。支持从单个文档中提取多条记录（例如表格的多行数据），
以及批量处理文件夹中的多个文档。

Features:
    - 单文件处理：提取单个DOCX文档中的结构化信息
    - 批量处理：处理文件夹中的所有DOCX文件
    - 多格式导出：支持JSON和CSV格式输出
    - 多记录提取：自动识别文档中的多行数据

Typical usage example:
    # 处理单个文件
    python main.py document.docx
    python main.py document.docx -o output.json

    # 批量处理文件夹，导出为CSV
    python main.py ./documents -o results.csv

    # 使用自定义API配置
    python main.py document.docx --api-key your-api-key --model gpt-4o
    python main.py document.docx --api-base https://api.openai.com/v1

Environment variables:
    OPENAI_API_KEY: OpenAI API密钥
    OPENAI_API_BASE: OpenAI API基础URL
    OPENAI_MODEL: 使用的模型名称（默认: gpt-4o-2024-08-06）

Note:
    - 文档中的每一行数据都会被提取为一个单独的记录，所有记录会以列表形式返回
    - 批量处理文件夹时，建议使用CSV格式以便整合所有文件的结果
    - CSV输出会包含源文件名，方便追踪每条记录的来源
"""

import argparse
import csv
import json
import os
import sys
from pathlib import Path
from typing import Optional

from docx import Document
from openai import OpenAI
from pydantic import BaseModel, Field


class DocumentFields(BaseModel):
    """文档字段结构化模型。

    Attributes:
        tl_ea: Column 1中的attached protocol - TL EA信息。
        test_standard: Column 2中的测试标准（非网站链接）。
        test_analytes: Column 5中的测试分析物。
        pp_notes: Column 3中的PP备注信息。
        source_link: Column 2中的网站链接（如果有）。
        label_and_symbol: 是否找到标签和符号（yes/no）。
    """
    tl_ea: str = Field(description="Column 1 of attached protocol - TL EA信息")
    test_standard: str = Field(description="Column 2 but not website - 测试标准（非网站链接）")
    test_analytes: str = Field(description="Column 5 - 测试分析物")
    pp_notes: str = Field(description="Column 3 - PP备注信息")
    source_link: Optional[str] = Field(default=None, description="Column 2 website if found - 来源链接（如果有网站）")
    label_and_symbol: str = Field(description="Any label found in this row, just state yes/no - 是否找到标签和符号")


class DocumentExtraction(BaseModel):
    """文档提取结果模型，包含多个记录。

    Attributes:
        records: 从文档中提取的所有记录列表。
    """
    records: list[DocumentFields] = Field(
        description="文档中提取的所有记录，每个记录对应文档中的一行数据"
    )


class DocxExtractor:
    """DOCX文档提取器，用于从Word文档中提取结构化信息。

    该类使用OpenAI的结构化输出功能，从DOCX文档中提取特定字段。

    Attributes:
        client: OpenAI客户端实例。
        model: 使用的OpenAI模型名称。
    """

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4o-2024-08-06",
        api_base: Optional[str] = None
    ):
        """初始化DocxExtractor实例。

        Args:
            api_key: OpenAI API密钥。
            model: 使用的模型名称，默认为 "gpt-4o-2024-08-06"。
            api_base: API基础URL，默认为None（使用OpenAI官方地址）。

        Raises:
            openai.OpenAIError: 如果API密钥无效。
        """
        client_kwargs = {"api_key": api_key}
        if api_base:
            client_kwargs["base_url"] = api_base

        self.client = OpenAI(**client_kwargs)
        self.model = model

    def read_docx(self, file_path: str) -> str:
        """读取docx文件并提取所有文本内容。

        该方法会提取文档中的段落文本和表格内容，并将它们组合成单个字符串。

        Args:
            file_path: DOCX文件的路径（相对或绝对路径）。

        Returns:
            包含文档所有内容的字符串，段落和表格内容会被分开标注。

        Raises:
            FileNotFoundError: 如果指定的文件不存在。
            docx.opc.exceptions.PackageNotFoundError: 如果文件不是有效的DOCX格式。
        """
        doc = Document(file_path)
        
        # 提取所有段落文本
        paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
        
        # 提取表格内容
        tables_content = []
        for table in doc.tables:
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                if any(row_data):  # 只添加非空行
                    tables_content.append(" | ".join(row_data))
        
        # 合并所有内容
        full_text = "\n".join(paragraphs)
        if tables_content:
            full_text += "\n\n=== 表格内容 ===\n" + "\n".join(tables_content)
        
        return full_text
    
    def extract_fields(self, text: str) -> DocumentExtraction:
        """使用OpenAI结构化输出API提取文档字段。

        该方法将文档文本发送到OpenAI API，使用结构化输出功能提取预定义的字段。
        一个文档可能包含多行数据，每行都会被提取为一个DocumentFields记录。

        Args:
            text: 从DOCX文档中提取的文本内容。

        Returns:
            包含所有提取记录的DocumentExtraction实例。

        Raises:
            openai.APIError: 如果API调用失败。
            openai.RateLimitError: 如果超出API速率限制。
        """
        completion = self.client.responses.parse(
            model=self.model,
            input=[
                {
                    "role": "system",
                    "content": """你是一个专业的文档信息提取助手。请从提供的文档内容中提取以下字段：

1. TL EA: 提取Column 1中的attached protocol信息
2. Test standard: 提取Column 2中的非网站内容（测试标准）
3. Test analytes: 提取Column 5中的测试分析物信息
4. PP notes: 提取Column 3中的备注信息
5. Source link: 如果Column 2中有网站链接，提取它；否则返回null
6. Label and symbol: 检查该行是否有任何标签，如果找到就返回"yes"，否则返回"no"

重要提示：
- 文档中可能包含多行数据（例如表格的多行）
- 请为每一行数据创建一个单独的记录
- 将所有记录放在records列表中返回
- 请仔细分析文档内容，准确提取这些信息。"""
                },
                {
                    "role": "user",
                    "content": f"请从以下文档内容中提取所有行的信息：\n\n{text}"
                }
            ],
            text_format=DocumentExtraction,
        )

        return completion.output_parsed

    @staticmethod
    def export_to_csv(extractions: list[tuple[str, DocumentExtraction]], output_path: str) -> None:
        """将多个提取结果导出为CSV文件。

        Args:
            extractions: 包含 (文件名, DocumentExtraction) 元组的列表。
            output_path: CSV输出文件路径。

        Raises:
            IOError: 如果无法写入文件。
        """
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)

            # 写入表头
            writer.writerow([
                'Source File',
                'TL EA',
                'Test Standard',
                'Test Analytes',
                'PP Notes',
                'Source Link',
                'Label and Symbol'
            ])

            # 写入每个文件的所有记录
            for filename, extraction in extractions:
                for record in extraction.records:
                    writer.writerow([
                        filename,
                        record.tl_ea,
                        record.test_standard,
                        record.test_analytes,
                        record.pp_notes,
                        record.source_link or '',
                        record.label_and_symbol
                    ])

    def process_file(self, file_path: str, output_path: Optional[str] = None) -> DocumentExtraction:
        """处理DOCX文件并提取结构化信息。

        该方法是主要的工作流程方法，它读取DOCX文件、提取字段，并可选地将结果保存到JSON文件。
        处理进度和结果会打印到标准输出。文档中可能包含多行数据，每行都会被提取。

        Args:
            file_path: 输入的DOCX文件路径（相对或绝对路径）。
            output_path: 可选的输出JSON文件路径。如果提供，结果将被保存为JSON格式。

        Returns:
            包含所有提取记录的DocumentExtraction实例。

        Raises:
            FileNotFoundError: 如果输入文件不存在。
            openai.APIError: 如果OpenAI API调用失败。
            IOError: 如果无法写入输出文件。
        """
        print(f"正在读取文件: {file_path}")
        text = self.read_docx(file_path)

        print(f"文档内容长度: {len(text)} 字符")
        print("\n正在使用OpenAI提取结构化信息...")

        extraction = self.extract_fields(text)

        print("\n提取完成！")
        print("=" * 80)
        print(f"共提取 {len(extraction.records)} 条记录\n")

        for idx, record in enumerate(extraction.records, 1):
            print(f"记录 #{idx}")
            print("-" * 80)
            print(f"  TL EA:           {record.tl_ea}")
            print(f"  Test standard:   {record.test_standard}")
            print(f"  Test analytes:   {record.test_analytes}")
            print(f"  PP notes:        {record.pp_notes}")
            print(f"  Source link:     {record.source_link}")
            print(f"  Label & symbol:  {record.label_and_symbol}")
            print()

        print("=" * 80)

        # 如果指定了输出路径，保存为JSON
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(extraction.model_dump(), f, ensure_ascii=False, indent=2)
            print(f"\n结果已保存到: {output_path}")

        return extraction


def parse_args() -> argparse.Namespace:
    """解析命令行参数。

    Returns:
        包含解析后参数的Namespace对象。
    """
    parser = argparse.ArgumentParser(
        description="从DOCX文档中提取结构化信息，支持单文件或文件夹批量处理",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 处理单个文件
  %(prog)s document.docx
  %(prog)s document.docx -o output.json

  # 处理文件夹中的所有DOCX文件，导出为CSV
  %(prog)s ./documents -o results.csv

  # 使用自定义API配置
  %(prog)s document.docx --api-key your-api-key --model gpt-4o
  %(prog)s document.docx --api-base https://api.openai.com/v1

环境变量:
  OPENAI_API_KEY     OpenAI API密钥（如果未通过 --api-key 指定）
  OPENAI_API_BASE    OpenAI API基础URL（如果未通过 --api-base 指定）
  OPENAI_MODEL       OpenAI模型名称（如果未通过 --model 指定，默认: gpt-4o-2024-08-06）
        """
    )

    parser.add_argument(
        "input_path",
        type=str,
        help="输入的DOCX文件路径或包含DOCX文件的文件夹路径"
    )

    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="输出文件路径（.json 或 .csv 格式）。处理文件夹时建议使用 .csv 格式"
    )

    parser.add_argument(
        "--api-key",
        type=str,
        default=None,
        help="OpenAI API密钥（优先于环境变量OPENAI_API_KEY）"
    )

    parser.add_argument(
        "--api-base",
        type=str,
        default=None,
        help="OpenAI API基础URL（优先于环境变量OPENAI_API_BASE）"
    )

    parser.add_argument(
        "--model",
        type=str,
        default=None,
        help="使用的模型名称（优先于环境变量OPENAI_MODEL，默认: gpt-4o-2024-08-06）"
    )

    parser.add_argument(
        "--json",
        action="store_true",
        help="以JSON格式输出到标准输出"
    )

    return parser.parse_args()


def main():
    """CLI工具主入口函数。

    解析命令行参数，验证输入，并处理DOCX文件以提取结构化信息。
    支持单文件处理和文件夹批量处理。结果会打印到标准输出，并可选地保存到JSON或CSV文件。

    Returns:
        int: 退出代码（0表示成功，1表示失败）。
    """
    args = parse_args()

    # 获取API配置（优先使用命令行参数，其次是环境变量）
    api_key = args.api_key or os.getenv("OPENAI_API_KEY")
    api_base = args.api_base or os.getenv("OPENAI_API_BASE")
    model = args.model or os.getenv("OPENAI_MODEL") or "gpt-4o-2024-08-06"

    if not api_key:
        print("错误: 未提供OpenAI API密钥", file=sys.stderr)
        print("请通过以下方式之一提供API密钥：", file=sys.stderr)
        print("  1. 使用 --api-key 参数", file=sys.stderr)
        print("  2. 设置环境变量 OPENAI_API_KEY", file=sys.stderr)
        print("     示例: export OPENAI_API_KEY='your-api-key-here'", file=sys.stderr)
        return 1

    # 验证输入路径
    input_path = Path(args.input_path)
    if not input_path.exists():
        print(f"错误: 路径不存在 - {args.input_path}", file=sys.stderr)
        return 1

    # 创建提取器实例
    extractor = DocxExtractor(
        api_key=api_key,
        model=model,
        api_base=api_base
    )

    try:
        # 判断是文件还是文件夹
        if input_path.is_file():
            # 处理单个文件
            if input_path.suffix.lower() not in ['.docx', '.doc']:
                print(f"警告: 文件可能不是DOCX格式 - {args.input_path}", file=sys.stderr)

            extraction = extractor.process_file(str(input_path), args.output)

            # 如果指定了 --json 标志，输出JSON格式
            if args.json:
                print("\n" + "=" * 80)
                print("JSON输出:")
                print(json.dumps(extraction.model_dump(), ensure_ascii=False, indent=2))

        elif input_path.is_dir():
            # 处理文件夹中的所有DOCX文件
            docx_files = list(input_path.glob("*.docx")) + list(input_path.glob("*.doc"))

            if not docx_files:
                print(f"错误: 文件夹中没有找到DOCX文件 - {args.input_path}", file=sys.stderr)
                return 1

            print(f"找到 {len(docx_files)} 个DOCX文件")
            print("=" * 80)

            extractions = []
            for idx, docx_file in enumerate(docx_files, 1):
                print(f"\n[{idx}/{len(docx_files)}] 处理文件: {docx_file.name}")
                print("-" * 80)

                try:
                    extraction = extractor.process_file(str(docx_file))
                    extractions.append((docx_file.name, extraction))
                except Exception as e:
                    print(f"警告: 处理文件 {docx_file.name} 时出错: {e}", file=sys.stderr)
                    continue

            # 保存结果
            if args.output:
                output_path = Path(args.output)
                if output_path.suffix.lower() == '.csv':
                    # 导出为CSV
                    DocxExtractor.export_to_csv(extractions, str(output_path))
                    print(f"\n所有结果已保存到CSV文件: {output_path}")
                elif output_path.suffix.lower() == '.json':
                    # 导出为JSON
                    all_data = {
                        "files": [
                            {
                                "filename": filename,
                                "records": extraction.model_dump()["records"]
                            }
                            for filename, extraction in extractions
                        ]
                    }
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(all_data, f, ensure_ascii=False, indent=2)
                    print(f"\n所有结果已保存到JSON文件: {output_path}")
                else:
                    print(f"警告: 不支持的输出格式 {output_path.suffix}，请使用 .csv 或 .json", file=sys.stderr)

            # 如果指定了 --json 标志，输出JSON格式到标准输出
            if args.json:
                all_data = {
                    "files": [
                        {
                            "filename": filename,
                            "records": extraction.model_dump()["records"]
                        }
                        for filename, extraction in extractions
                    ]
                }
                print("\n" + "=" * 80)
                print("JSON输出:")
                print(json.dumps(all_data, ensure_ascii=False, indent=2))

        else:
            print(f"错误: 输入路径既不是文件也不是文件夹 - {args.input_path}", file=sys.stderr)
            return 1

        return 0

    except FileNotFoundError as e:
        print(f"错误: 文件未找到 - {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"错误: 处理过程中出错 - {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())