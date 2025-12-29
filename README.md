# PPTX 翻译工具

一个使用 AWS Bedrock 将 PowerPoint 演示文稿翻译成简体中文的 Python 工具。

## 功能特点

- 自动翻译 PPTX 文件中的所有文本内容
- 支持幻灯片和备注的翻译
- 保留原始格式、换行符和空白字符
- 保留公司名称、产品名称、技术术语、专有名词、数字和 URL
- 批量处理文本以提高效率

## 前置要求

- Python 3.x
- AWS 账户并配置好凭证
- 访问 AWS Bedrock 服务的权限

## 安装

1. 克隆或下载此项目

2. 安装依赖：
```bash
pip install boto3
```

3. 配置 AWS 凭证（如果尚未配置）：
```bash
aws configure
```

## 使用方法

基本用法：
```bash
python3 translate_pptx.py input.pptx output.pptx
```

启用调试模式：
```bash
python3 translate_pptx.py input.pptx output.pptx --debug
```

### 参数说明

- `input.pptx`: 输入的 PowerPoint 文件路径
- `output.pptx`: 输出的翻译后文件路径
- `--debug`: （可选）启用调试日志，显示详细的翻译过程

## 工作原理

1. 解压 PPTX 文件（本质上是 ZIP 压缩包）
2. 提取幻灯片和备注中的所有文本元素
3. 批量调用 AWS Bedrock 模型进行翻译
4. 将翻译后的文本替换回原始 XML 结构
5. 重新打包生成新的 PPTX 文件

## 自定义配置

### 更换模型

默认使用 Claude Sonnet 4.5，如需使用其他 Bedrock 模型，修改 `translate_pptx.py` 中的 `model_id`：

```python
def call_llm(prompt):
    model_id = "us.anthropic.claude-sonnet-4-20250514-v1:0"  # 修改此处
    # 其他可选模型示例：
    # model_id = "amazon.nova-pro-v1:0"
```

### 更换区域

默认使用 `us-west-2` 区域，可在代码开头修改：

```python
bedrock_client = boto3.client("bedrock-runtime", region_name="us-west-2")
```

## 注意事项

- 翻译过程会保持原始文本的结构和格式
- 已经是简体中文或空白的文本将保持不变
- 确保 AWS 账户有足够的 Bedrock 使用配额

本项目仅供学习和参考使用。
