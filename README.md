# Word to WordPress WXR

An OpenAI Codex skill for converting Chinese SEO articles in Word documents
into reviewable and validated WordPress WXR draft imports for the Style3D Help
site.

一个 OpenAI Codex 技能，用于将 Word 文档中的中文 SEO 文章转换为可审查和验证的 WordPress WXR 草稿导入，供 Style3D 帮助中心使用。

## Contents / 内容

- `SKILL.md`: workflow and content-review requirements. / 工作流和内容审查要求。
- `scripts/generate_wxr.py`: Word-to-WXR converter and validator. / Word 到 WXR 转换器和验证器。
- `scripts/schedule_wxr.py`: creates a separate scheduled-import WXR. / 创建单独的定时导入 WXR。
- `agents/openai.yaml`: Codex skill metadata. / Codex 技能元数据。

## Use in Codex

## 在 Codex 中使用

Install this directory as a Codex skill, then invoke
`$word-to-wordpress-wxr` in a task. The skill documentation describes the
required dry-run, category approval, content audit, and validation workflow.

将此目录安装为 Codex 技能，然后在任务中调用 `$word-to-wordpress-wxr`。技能文档描述了所需的测试运行、类别批准、内容审计和验证工作流。

The converter does not modify input Word documents or publish/import posts.

转换器不会修改输入的 Word 文档或发布/导入帖子。

