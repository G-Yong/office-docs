# Agent Instructions for office-docs

This repository is a **skills library**. Each subdirectory of `skills/` contains
a `SKILL.md` whose YAML frontmatter has a `name` and a `description` that starts
with "Use when ...". The description states the *triggering conditions* for the
skill.

## How to use these skills

1. When a task involves reading or writing Microsoft Office files (Word `.docx`/
   `.doc`, Excel `.xlsx`/`.xls`/`.csv`) on Windows, scan the skills below.
2. If a skill's `description` matches, **read its full `SKILL.md` before acting**
   — do not work from the description alone.
3. Follow the skill's patterns exactly, especially the open → try/finally →
   release COM lifecycle.

## Skill index

| Skill | Trigger summary |
|-------|-----------------|
| `office-docs-overview` | Routing + decide COM vs library + prerequisites |
| `word-com-powershell` | Word read/write/edit/generate/PDF via COM |
| `excel-com-powershell` | Excel cells/ranges/formulas/sheets/PDF via COM |
| `office-com-cleanup` | Releasing COM objects; leaked/locked/hung processes |

## Non-negotiable rules when generating Office automation

- **Start at `office-docs-overview`** if unsure which approach fits.
- **Always release COM objects** and `Quit()` the app inside a `finally` block.
  Read `office-com-cleanup`.
- **Always use absolute paths** for `Open`/`SaveAs`.
- **Office COM collections are 1-based.**
- **Do not run Office COM in CI / as a service account.** Recommend a library
  instead (see `office-docs-overview/references/non-com-alternatives.md`).
- Set `Visible = $false` and disable `DisplayAlerts` to avoid hidden modal
  dialogs that hang the script.
