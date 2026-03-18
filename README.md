# 空管请示 Word 模板生成器

这个目录用于生成“固定格式”的请示类 Word（`.docx`）文件：用你提供的 `.dotx` 模板套版，保证标题、正文、各级标题、落款等格式与模板一致。

## 快速开始

你提供的空管请示模板已复制到项目里：`atc_request_word/templates/空管请示模板.dotx`。

1. 准备内容（统一放在 `content/` 目录下，例如 `content/content_采购油机爬梯_请示.json`）
2. 套用模板生成 docx：

```bash
python atc_request_word/generate_from_template.py \
  --template atc_request_word/templates/空管请示模板.dotx \
  --content atc_request_word/content/content_采购油机爬梯_请示.json \
  --out atc_request_word/out.docx

约定：后续所有请示内容 JSON 默认存放在 `content/` 目录，不再放在项目根目录。
```
