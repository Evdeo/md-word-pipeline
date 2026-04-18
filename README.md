# md-to-docx

Word documents are a mess. Open the same file on two machines and something looks different. Use someone else's template and the styles are wrong. Save through a different version of Word and the spacing shifts. I spent more time fixing formatting than writing.

I tried LaTeX and Typst — both are great, you write once and it always looks the same. But at work Word is the standard. People need to open it, edit it, send it back. A PDF doesn't cut it.

So I write in Markdown, keep everything about the document's appearance in a config file, and treat Word as the output format.

---

## What it does

**Consistent output.** Fonts, heading sizes, colours, table styles, margins, headers, footers — all in `configs/<name>.yaml`, shared across projects. Same config, same document, every time.

**Auto-numbered figures and tables.** Tag a caption with `{#fig-arch}` and reference it as `[Figure 1](#fig-arch)`. Reorder the figures — the numbers update.

**Variables.** Put `{{project.name}}` or `{{document.version}}` anywhere. The values live in YAML. Change one, rebuild, every instance updates.

**Handling reviewed files.** You send a Word file, someone edits it, sends it back. Normally you're manually hunting for changes and hoping you don't overwrite your source. The review workflow here handles it properly.

---

## Getting started

```bash
pip install -r lib/requirements.txt
python md.py showcase
```

The first command opens a browser with the live preview of the bundled
showcase project. Edit any `.md` / `.yaml` file in `projects/showcase/`
and the preview rebuilds within a few seconds. Ctrl+C stops it.

Requires Python 3.10+ and Microsoft Word (for the preview's
docx → PDF step).

### Everything else (rare)

```
python md.py                          pick project, then action, from a menu
python md.py <project>                live preview (the 95% case)
python md.py <project> -build         build once → projects/<project>/output/
python md.py <project> -diff          section-diff vs output/received/
python md.py <project> -open          open the built docx in Word
python md.py <project> -import FILE   extract a docx into this project
python md.py -new <name>              scaffold a new empty project
```

`-b`, `-d`, `-o`, `-i` are short aliases.

---

## Creating a project

```bash
python md.py -new my-doc
```

Makes `projects/my-doc/` from `projects/_template/`. It refuses if the
name is taken — never overwrites. Edit the YAML files in VS Code, then
`python md.py my-doc` to start the preview.

To bootstrap from an existing Word file:

```bash
python md.py -new my-doc
python md.py my-doc -import path/to/existing.docx
```

`-import` refuses to run if `content.md` is already populated, so you
can't clobber a live project by accident.

---

## Project layout

```
my-project/
├── project.yaml            ← names which config/ preset to use
├── content.md              ← write here (or several *.md files)
├── 00-frontpage.md         ← optional cover page
├── document-info.yaml      ← title, author, version, revision history
├── properties.yaml         ← custom variables
├── revisions.yaml          ← revision-history table
├── images/
└── output/
    ├── document.docx
    ├── received/           ← drop reviewed files here
    └── review_report.html
```

The shared layout config lives at the repo root under `configs/` — e.g.
`configs/default.yaml`. A project says which config to use via its
`project.yaml`:

```yaml
config: default                 # required: references configs/<name>.yaml
output: my-doc.docx             # optional: override the output filename
title_override: "Draft 3"       # optional: override document-info.yaml title
```

---

## Review workflow

1. Build and send the document
2. Someone edits it and sends it back
3. Drop the file in `output/received/`
4. Run `python md.py <project> -diff`

Builds a fresh copy from your current source, diffs it section by section against the received file, opens an HTML report.

Both versions shown side by side, rendered to look like the actual document:

- ✅ Identical sections — collapsed, expand if you want to check
- 🟡 Changed — open, both sides visible
- 🔴 Removed — shows what was there
- 🟢 Added — shows what's new
- ⬆️ Moved — shown in new position

Tree overview at the top, click anything to jump there.

---

## Markdown

### Merged cells

```markdown
| Region | <<      | Sales |
|--------|---------|-------|
| EMEA   | UK      | 142   |
| ^^     | Germany | 98    |
```

`<<` — consumed by the cell to its left (colspan)
`^^` — consumed by the cell above (rowspan)

You can optionally use `{cs=2}` and `{rs=2}` on the anchor cell to be explicit, but `<<` and `^^` alone are sufficient.

Column widths: add `{col-widths="20%,50%,30%"}` after the table.

### Images

```markdown
![Alt](image.png){.medium}
```

`.xs` 20% · `.small` 30% · `.medium` 50% · `.large` 75% · `.xl` 100%

Side by side:

```markdown
:::figures
![Left](a.png)
![Right](b.png)
:::

*Figure 1: Caption. {#fig-1}*
```

### Image overlays (arrows, callouts, annotations)

For user-guide screenshots where you want arrows, boxes, and text
callouts on top of a base image — all kept consistent across builds and
editable in Word afterwards:

```markdown
:::overlay {#fig-login width=medium}
![Login screen](screens/login.png)
::arrow    from=20%,30% to=50%,35% color=#FF0000 stroke=2
::rect     at=10%,20% size=30%,10% color=#FF0000 stroke=2
::ellipse  at=55%,45% size=12%,12% color=#FF0000 stroke=2
::callout  at=60%,40% size=22%,10% text="Click here" color=#0000FF fill=#FFFF99
:::
```

Coordinates are percent-of-image, so shapes scale with the image if you
resize the base. Available shapes:

| Kind       | Required attrs                      | Optional attrs                    |
|------------|-------------------------------------|-----------------------------------|
| `arrow`    | `from=X%,Y% to=X%,Y%`               | `color` `stroke`                  |
| `rect`     | `at=X%,Y% size=W%,H%`               | `color` `stroke` `fill`           |
| `ellipse`  | `at=X%,Y% size=W%,H%`               | `color` `stroke` `fill`           |
| `callout`  | `at=X%,Y% size=W%,H%`               | `color` `stroke` `fill` `text`    |

The base image is embedded via python-docx and then wrapped in a native
Word `wpg:wgp` group drawing alongside `wps:wsp` shapes — reviewers can
drag an arrow in Word, the review importer will pick up the new
position on the next round-trip.

Requires Word 2010 or later. Very old Word versions won't render the
group and will show only the base image.

### Cross-references

```markdown
*Figure 1: System overview. {#fig-1}*

*Table: Feature status. {#tbl-features}*
```

Reference them anywhere in the document:

```markdown
As shown in [***Figure***](#fig-1), the system has three layers.

As shown in [***Table***](#tbl-features), all features are implemented.
```

### Alerts

```markdown
> [NOTE] Worth knowing.
> [TIP] Easier way to do it.
> [WARNING] This could go wrong.
> [CAUTION] This will cause problems.
```

### Variables

`properties.yaml`:
```yaml
project:
  name: "Falcon"
  version: "1.4"
```

```markdown
{{project.name}} version {{project.version}}
```

Unknown placeholders are left literal and a warning is logged so typos
like `{{projec.name}}` don't silently produce empty text. To put a
literal `{{` into the output, escape it with a backslash: `\{{literal}}`.

### Appendix

```markdown
:::appendix
# A. Reference {.nonumber}

## A.1. First subsection {.nonumber}

Headings become A, A.1, A.2, B, B.1 etc.
:::
```

`{.nonumber}` is required on each appendix heading to suppress the normal numbered heading style.

### Page breaks and spacing

```markdown
---                  ← horizontal rule (line across the page)
+++                  ← page break
:::space{lines=2}    ← blank lines
:::space{pt=24}      ← exact points
```

---

## Defaults

Edit `configs/default.yaml` in VS Code. Every project that says
`config: default` in its `project.yaml` rebuilds against the updated
values next time you run the preview or `-build`. For alternate
presets, drop a `configs/<name>.yaml` next to it and reference it
from `project.yaml`.

---

## Running the tests

```bash
pip install -r lib/requirements.txt
pytest tests/ -v
```

Covers: bullet/ordered/nested lists, table merge markers (`<<` / `^^`)
and column widths, heading numbering and appendix mode, image size
classes, property substitution, section diff (identical / changed /
moved / added / removed), overlay parser and XML emission, overlay
round-trip through docx, and an end-to-end deterministic build of the
showcase project.

---

## Licence

MIT
