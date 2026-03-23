# md-to-docx

Word documents are a mess. Open the same file on two machines and something looks different. Use someone else's template and the styles are wrong. Save through a different version of Word and the spacing shifts. I spent more time fixing formatting than writing.

I tried LaTeX and Typst — both are great, you write once and it always looks the same. But at work Word is the standard. People need to open it, edit it, send it back. A PDF doesn't cut it.

So I write in Markdown, keep everything about the document's appearance in a config file, and treat Word as the output format.

---

## What it does

**Consistent output.** Fonts, heading sizes, colours, table styles, margins, headers, footers — all in `config.yaml`. Same config, same document, every time.

**Auto-numbered figures and tables.** Tag a caption with `{#fig-arch}` and reference it as `[Figure 1](#fig-arch)`. Reorder the figures — the numbers update.

**Variables.** Put `{{project.name}}` or `{{document.version}}` anywhere. The values live in YAML. Change one, rebuild, every instance updates.

**Handling reviewed files.** You send a Word file, someone edits it, sends it back. Normally you're manually hunting for changes and hoping you don't overwrite your source. The review workflow here handles it properly.

---

## Getting started

```bash
python run.py
```

Dependencies install on first run. Requires Python 3.10+ and Word to open the output.

---

## Creating a project

Three options:

1. **Minimal template** — bare bones
2. **Full template** — complete example with every feature
3. **Import from Word** — converts an existing `.docx` to Markdown *(rough around the edges, some manual cleanup needed)*

---

## Project layout

```
my-project/
├── input/
│   ├── content.md          ← write here
│   ├── config.yaml         ← controls how the document looks
│   ├── document-info.yaml  ← title, author, version, revision history
│   └── properties.yaml     ← custom variables
└── output/
    ├── document.docx
    ├── received/           ← drop reviewed files here
    └── review_report.html
```

---

## Inside a project

- **Build** — generates the Word file
- **Export** — copies it somewhere
- **Open document** — opens it in Word
- **Open in VS Code** — opens the project folder
- **Review changes** — compares a received file against your source
- **Edit document info** — title, author, version, classification
- **Edit properties** — custom variables

---

## Review workflow

1. Build and send the document
2. Someone edits it and sends it back
3. Drop the file in `output/received/`
4. Select **Review changes**

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

The main menu has a **Change defaults** option for the config applied to every new project. Edit the YAML directly, copy settings from a project, or go field by field and pick what to keep.

---

## Licence

MIT
