# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
Formats Romanian academic/conference papers (.docx) according to strict conference formatting rules using Claude AI for section detection and python-docx for output generation.

## Running
```bash
python formatter.py input.docx [-o output.docx] [--dry-run] [--show-classification]
```
- `--dry-run`: classify paragraphs without generating output (useful for debugging classification)
- `--show-classification`: print full paragraph-by-paragraph classification before generating output
- `--model`: override the Claude model used for classification (default: claude-sonnet-4-5-20250514)

There is also an older rule-based formatter (no Claude dependency):
```bash
python format_docx.py input.docx [-o output.docx] [--authors NAME ...] [--title-en TEXT] [--abstract-en TEXT]
```

## Architecture

### `formatter.py` — Main pipeline (Claude-powered)
Three-stage pipeline:
1. **Extract** (`extract_paragraphs`): Walks body XML directly (not `doc.paragraphs`) to get paragraphs AND table content in document order. Resolves auto-numbering from the numbering part XML to prepend bullet/number prefixes. Returns a list of element dicts with text, style, alignment, bold/italic flags, indent values, and table grid data.
2. **Classify** (`classify_with_claude`): Sends all paragraphs (with formatting metadata hints like bold, centered, style) to Claude via the Anthropic API. Claude returns a JSON array mapping each paragraph index to one of ~20 section types (defined in `SECTION_TYPES`). The system prompt contains all classification rules.
3. **Build** (`build_formatted_document`): Creates a fresh template via `create_template()`, then iterates through paragraphs applying the classified section type to choose the correct style and spacing. Spacing between sections (blank lines) is managed by tracking `prev_type` and section-transition flags.

### `create_template.py` — Template generator
Called by `formatter.py` at the start of each run. Creates `template_conference.docx` with all custom styles pre-configured (Author, Figure Caption, Table Caption, Bibliography Header, Bibliography Entry, etc.). Sets document-level defaults via direct XML manipulation of `w:docDefaults`. Returns the template path.

### `format_docx.py` — Legacy standalone formatter
Rule-based section detection using heuristics (searches for "REZUMAT", "BIBLIOGRAFI", centered bold text, etc.). Does not use Claude. Builds a new document from scratch by copying text content. Does not handle tables-in-document-order or auto-numbering.

## Key Dependencies
- python-docx, anthropic, python-dotenv, lxml
- Requires `ANTHROPIC_API_KEY` in `.env`

## Conference Formatting Rules
- **Page**: ISO B5 (176x250mm), margins 20mm all sides, header/footer 12.7mm
- **Font**: Times New Roman, Size 11, single spacing throughout
- **TAB indent**: 12.7mm (first-line indent for body text)
- **Title**: bold, 12pt, UPPERCASE, center + 6 blank lines after
- **Authors**: normal, 11pt, center + 1 blank line after
- **Rezumat label**: bold, 11pt, justify, indent 1 TAB
- **Rezumat text**: italic, 11pt, justify, indent 1 TAB + 2 blank lines after
- **Body text**: normal, 11pt, justify, indent 1 TAB
- **Headings**: bold, 11pt, hanging indent (number at margin, text at TAB)
- **English title**: normal, 12pt, UPPERCASE, center + 2 blank lines before/after
- **Abstract label**: bold, 11pt, indent 1 TAB
- **Abstract text**: italic, 11pt, justify + 4 blank lines after
- **Bibliografie header**: bold, 12pt, center, 3 blank lines before
- **Figure captions**: normal, 9pt, center
- **Table captions**: bold, 10pt, center

## Important Implementation Details
- Tables are handled specially: single-cell tables (often used for abstract boxes) have their paragraphs extracted inline; multi-cell data tables store the full grid in `table_data` and are reconstructed as Word tables in the output.
- The Claude classification prompt truncates paragraph text to 200 chars and includes formatting hints (`[bold, centered, style:Heading1]`) to aid classification without sending full content.
- Blank line spacing is entirely controlled by section-transition logic in `build_formatted_document`, not by preserving source empty paragraphs (those are classified as "empty" and skipped).
