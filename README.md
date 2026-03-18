# 空管请示 Word 模板生成器

这个目录用于生成“固定格式”的请示类 Word（`.docx`）文件：用你提供的 `.dotx` 模板套版，保证标题、正文、各级标题、落款等格式与模板一致。

## 快速开始

你提供的空管请示模板已复制到项目里：`atc_request_word/templates/空管请示模板.dotx`。

1. 准备内容（参考 `sample_content.json`）
2. 套用模板生成 docx：

```bash
python atc_request_word/generate_from_template.py \
  --template atc_request_word/templates/空管请示模板.dotx \
  --content atc_request_word/sample_content.json \
  --out atc_request_word/out.docx
```
