# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Ruby gem for reading and writing Microsoft Excel XLS files (BIFF5/BIFF8 binary format). Does **not** support XLSX. Runtime dependencies: `ruby-ole`, `logger`, `bigdecimal`.

## Common Commands

```bash
# Install dependencies
bundle install

# Run tests
bundle exec ruby test/suite.rb

# Lint (StandardRB)
bundle exec rake standard

# Auto-fix lint issues
bundle exec standardrb --fix

# Build gem
bundle exec rake build
```

There is no way to run a single test file independently — all tests are loaded via `test/suite.rb`. To run a subset, you can run `bundle exec ruby -Ilib test/<file>.rb` for files that don't depend on the suite loader.

## Architecture

**Three-tier data model**: `Workbook` → `Worksheet` → `Row`

- **`lib/spreadsheet.rb`** — Entry point. `Spreadsheet.open(path)` reads a file, `Spreadsheet.writer(path)` creates a writer. Global `client_encoding` setting (default UTF-8).
- **`lib/spreadsheet/workbook.rb`** — Container for worksheets, formats, fonts, and color palette.
- **`lib/spreadsheet/worksheet.rb`** — Container for rows/columns. Includes `Enumerable`. Access cells via `worksheet[row][col]` or `worksheet.cell(row, col)`.
- **`lib/spreadsheet/row.rb`** — Extends `Array`. Each cell is an element. Tracks per-cell formats and notifies its worksheet of changes via `updater` DSL.
- **`lib/spreadsheet/format.rb`** / **`font.rb`** — Cell styling (borders, colors, alignment, fonts). Attributes defined via the `Datatypes` metaprogramming DSL (`boolean`, `enum`, `colors` macros in `datatypes.rb`).

**Excel binary layer** (`lib/spreadsheet/excel/`):

- `reader.rb` + `reader/biff8.rb` / `biff5.rb` — Binary XLS parser. Decodes opcodes, RK values, Shared String Table (SST), codepages.
- `writer/workbook.rb` + `writer/worksheet.rb` + `writer/biff8.rb` — Binary XLS generator.
- `internals/biff8.rb` / `biff5.rb` — Opcode constants and record definitions.
- `Excel::Row` overrides `Row#[]` to convert Excel date serial numbers to Ruby `Date`/`DateTime`.

**Legacy**: `lib/parseexcel/` provides backward compatibility with the old ParseExcel API.

## Testing

- Framework: **Test::Unit** (not RSpec)
- Test files in `test/` mirror the lib structure
- `test/integration.rb` is the main integration test suite (~1500 lines)
- Test data (sample XLS files) in `test/data/`
- SimpleCov enabled for coverage (requires >= 0.22 for Ruby 3.3+ compatibility)

## Code Style

- **StandardRB** — no custom RuboCop config. Run `bundle exec rake standard` to check.
- Ruby 2.6+ minimum (uses `[x..]` syntax). Ruby 3.3.0 in `.ruby-version`.
- CI tests against Ruby 2.6 through 3.4, JRuby, and ruby-head.
