# DOCX文档信息提取工具

使用OpenAI的结构化输出功能从DOCX文档中提取特定字段。

## 功能特点

- 📄 读取DOCX文档内容（包括段落和表格）
- 🤖 使用OpenAI GPT-4o进行智能信息提取
- 📊 结构化输出，确保数据格式一致
- 💾 支持JSON格式导出

## 提取字段

根据文档模板提取以下字段：

| 字段 | 说明 |
|------|------|
| TL EA | Column 1的attached protocol信息 |
| Test standard | Column 2的测试标准（非网站） |
| Test analytes | Column 5的测试分析物 |
| PP notes | Column 3的备注信息 |
| Source link | Column 2的网站链接（如果有） |
| Label and symbol | 是否有标签（yes/no） |

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 设置OpenAI API密钥

```bash
export OPENAI_API_KEY='your-api-key-here'
```

或者在Windows上：
```cmd
set OPENAI_API_KEY=your-api-key-here
```

### 2. 运行提取程序

```python
from extract_docx import DocxExtractor

# 初始化提取器
extractor = DocxExtractor(api_key="your-api-key")

# 处理文件
fields = extractor.process_file(
    file_path="your_document.docx",
    output_path="extracted_data.json"
)

# 访问提取的字段
print(fields.tl_ea)
print(fields.test_standard)
print(fields.test_analytes)
```

### 3. 命令行使用

直接修改 `extract_docx.py` 中的 `input_file` 变量，然后运行：

```bash
python extract_docx.py
```

## 示例输出

```json
{
  "tl_ea": "Protocol XYZ-123",
  "test_standard": "ISO 9001:2015",
  "test_analytes": "pH, Temperature, Moisture",
  "pp_notes": "Sample tested under standard conditions",
  "source_link": "https://example.com/standard",
  "label_and_symbol": "yes"
}
```

## 技术栈

- **OpenAI API**: 使用GPT-4o模型进行智能提取
- **python-docx**: 读取DOCX文档
- **Pydantic**: 数据验证和结构化输出

## 注意事项

- 确保使用的OpenAI模型支持结构化输出（如 gpt-4o-2024-08-06）
- API调用会产生费用，请注意控制使用
- 首次运行需要联网下载依赖包

