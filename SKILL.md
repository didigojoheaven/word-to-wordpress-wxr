---
name: word-to-wordpress-wxr
description: Convert Chinese SEO articles in Word .docx files, folders, or ZIP archives into reviewable and validated WordPress WXR/XML drafts for the Style3D Help site, then optionally create a separate scheduled-post WXR. Use for WordPress imports requiring Yoast metadata, category confirmation, SEO quality review, or timed release plans. Never use this skill to publish posts or import XML; retain original source files and make any approved content revisions only in reviewed copies.
---

# Word To WordPress WXR

Create a reviewable WordPress draft-import or scheduled-post WXR file. Do not
publish posts, import XML, or rewrite source text.

## Workflow

1. Locate the input `.docx` files, a folder, or a ZIP archive. Article filenames
   use the form `L-26-0253.Article Title.docx`.
2. Run `scripts/generate_wxr.py --dry-run`. Review its CSV for categories and
   SEO warnings before creating any XML.
3. Ask the user to approve each pending category. Store approvals in an
   overrides JSON object keyed by article ID, then run the converter again.
4. Use `--fail-on-seo-warnings` when all titles, descriptions, and keywords
   must clear the SEO review before XML is generated.
5. Run validation on the generated XML before handing it over.
6. When a release plan is requested, create and validate a separate scheduled
   WXR. Keep the draft-import WXR unchanged.
7. Tell the user to import the final XML in WordPress at Tools > Import > WordPress.

## Intro Images

When the user asks for one image before every article body, use only public URLs
for images already in the WordPress media library. Do not put local file paths
in WXR, upload media, publish posts, or change the original Word files.

Enable the standard 53-image `Frame-1.png` to `Frame-53.png` Style3D media
pool with `--use-style3d-frame-pool`. Supply one persistent
`--intro-image-state` file for the whole publication stream. Images are chosen
uniformly from unused entries, removed after selection, and cannot repeat until
the full pool has been exhausted; the state advances only after WXR validation
passes. The generated `<output>.intro-images.csv` records every assignment.

For another approved media pool, use `--intro-image-pool` with a JSON URL array
or `{ "images": ["https://..."] }`. Require user approval of the media pool
before generation. Image mode is off by default, so normal WXR behavior is
unchanged.

## Text Review Before Conversion

Run a content audit before creating any WXR. The converter writes
`<output>.content-audit.csv` and blocks XML when a configured competitor name
is found. The default list covers common apparel-3D competitors; provide the
client-approved list with `--competitor-list` when it is available.

For every flagged article, perform this sequence before rerunning conversion:

1. Report each competitor name, its quoted context, and the affected article.
2. Keep the original Word file unchanged. Create a reviewed copy in a separate
   folder; remove the name and rewrite the surrounding passage only as needed
   for coherent, logical Style3D-oriented prose. Do not change the article's
   subject, scope, or unsupported factual claims.
3. Review the whole revised article semantically for indirect competitor
   promotion, comparative claims, invented capabilities, outdated statements,
   and unsupported statistics. Preserve neutral factual content; remove or
   qualify claims that cannot be supported by the supplied source material.
4. Report every changed passage as `original -> revised`, state why it changed,
   and identify any authority or accuracy concern. Do not state that a claim is
   verified unless the user supplied a reliable source for it.
5. Rerun the audit against the reviewed-copy folder. Generate WXR only when it
   has no competitor-name matches and the user has received the change report.

Use a JSON array or `{ "competitors": ["..."] }` for the optional list:

```json
["CLO 3D", "Marvelous Designer", "Browzwear"]
```

## Category Rules

- Use `Solution` for clear Style3D, AI fashion, 3D+AI fashion, or explicit
  implementation and solution topics.
- Use `Industry knowledge` for non-brand definitions, properties, tutorials,
  principles, and other knowledge explainers.
- Mark mixed or unclear titles as pending for user approval.

## Run The Converter

Use the bundled Python runtime when available. The input may be a single Word
file, a ZIP archive, or a folder. The converter never changes source files.

```powershell
& <python> scripts/generate_wxr.py `
  --input "C:\Articles" `
  --output "C:\Articles\wordpress-drafts.xml" `
  --site-url "https://help-zh.style3d.com" `
  --category-overrides "C:\Articles\category-overrides.json" `
  --competitor-list "C:\Articles\competitors.json" `
  --use-style3d-frame-pool `
  --intro-image-state "C:\Articles\style3d-frame-rotation.json"
```

Use `--dry-run` for the first pass. The adjacent `.review.csv` records the
category decision plus keyword, SEO title length, description length, and any
SEO warnings. The converter never changes source files. To block output until
every SEO warning is resolved in the source Word fields, add
`--fail-on-seo-warnings`.

Defaults target the Style3D Help site, author `weichanghua` (ID `3`), and
brand suffix `Style3D`. Override these only when the user gives different site
settings.

## Output Requirements

- Never change an original source file. Text revision is allowed only in a
  separately identified reviewed copy after the content-audit workflow.
- Render body section headings as unbolded `h2` elements with `font-size: 20px`.
- Render standalone numeric and Chinese-numeric subheadings as bold paragraphs
  with `font-size: 16px`, not as headings.
- Preserve basic inline bold, italic, underline, lists, and tables in normal
  body paragraphs.
- Read explicit keyword, SEO title, and SEO description fields when present.
- Create posts as `draft` with Yoast focus keyword, title, and description
  metadata.

## Validate

```powershell
& <python> scripts/generate_wxr.py --validate-only "C:\Articles\wordpress-drafts.xml"
```

For an image-enabled WXR, add `--require-intro-images` to validation. Confirm
that the adjacent intro-image CSV has one assignment per post before handoff.

Do not claim success until validation passes and the category review CSV has no
pending records. SEO warnings require explicit user approval unless
`--fail-on-seo-warnings` was used and passed.

## Schedule A Separate Import

Only schedule after the draft WXR passes validation and the user gives a future
start time, interval, and timezone. This command never overwrites its input.

```powershell
& <python> scripts/schedule_wxr.py `
  --input "C:\Articles\wordpress-drafts.xml" `
  --output "C:\Articles\wordpress-scheduled.xml" `
  --start "2026-08-10 09:00" `
  --interval-minutes 1440 `
  --timezone "Asia/Shanghai"
```
