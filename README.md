# Word to WordPress WXR

An OpenAI Codex skill for converting Chinese SEO articles in Word documents
into reviewable and validated WordPress WXR draft imports for the Style3D Help
site.

## Contents

- `SKILL.md`: workflow and content-review requirements.
- `scripts/generate_wxr.py`: Word-to-WXR converter and validator.
- `scripts/schedule_wxr.py`: creates a separate scheduled-import WXR.
- `agents/openai.yaml`: Codex skill metadata.

## Use in Codex

Install this directory as a Codex skill, then invoke
`$word-to-wordpress-wxr` in a task. The skill documentation describes the
required dry-run, category approval, content audit, and validation workflow.

The converter does not modify input Word documents or publish/import posts.
