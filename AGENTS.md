# rexcell

Rust CLI + eframe/egui GUI for updating Excel (`.xlsx`) files. Reads a target workbook and a reference/analysis workbook, matches rows by key column, copies/accumulates values, applies Excel formulas, and writes results.

## Binaries (defined in Cargo.toml)

- `cargo run --bin rexcell -- <args>` — CLI (`src/main.rs`)
- `cargo run --bin gui` — GUI (`src/gui.rs`)
- Library entrypoint used by both: `excell::execute(&common::Config)` in `src/lib/excell.rs`; modules declared in `src/lib/mod.rs` (`common`, `excell`, `range_ops`, `range_types`).

## Verification

- `cargo check` — typecheck (fast)
- `cargo build` — debug build
- `cargo build --release --target x86_64-pc-windows-gnu` — Windows release build used by the packager metadata (`[package.metadata.packager]` → NSIS installer; `wix/main.wxs` is an empty placeholder)
- No test suite, no CI workflows, no pre-commit config, no `opencode.json`, no `rust-toolchain` file

## CLI commands & flags

Command is passed with `-c/--command`; values:
- `cmd-list-sheets` — list sheets in target file
- `cmd-filter-sheets` — unique rows from `-s` col into new sheet `-n`, accumulate `-a` col on duplicates
- `cmd-update-sheets` — copy `-v` col from reference file `-r`/sheet `-e` keyed by `-k` into target `-d` cols
- `cmd-autocomplete-sheets` — filter + merge analysis data (`-a` analysis file, `-o` table) into target

Common flags (short = long, defaults in `src/lib/common.rs`):
- `-t/--tgt-file` — target Excel file (default `""`)
- `-u/--tgt-sheets` — comma-separated sheets to update (default `""`)
- `-s/--tgt-src-col` — source column, e.g. `C` (default `C`)
- `-d/--tgt-dest-col` — destination columns, comma-separated, e.g. `B,F` (default `B,F`)
- `-c/--tgt-accum-col` — accumulate column (default `E`); note `-c` is also used for `--command`
- `-r/--ref-file` — reference Excel file (default `""`)
- `-e/--ref-sheet` — reference sheet name (default `""`)
- `-k/--ref-col-key` — key column in reference (default `B`)
- `-v/--ref-col-value` — value column in reference (default `C`)
- `-n/--new-sheet-name` — new sheet name (default `Prices`)
- `-i/--inplace` — overwrite input file instead of creating `_new.xlsx`
- `-a/--anlysis-file` — analysis Excel file (note the typo in the long flag name)
- `-o/--anlysis-table` — analysis sheet name (note the typo)

Output: unless `-i` is given, results are written to `<target>_new.xlsx` (`NEW_FILE_SUFFIX = "_new.xlsx"`). Sheet names and cell content may be Cyrillic; string comparison (`cmp_strs`) lowercases and strips whitespace.

## Hardcoded limits

`src/lib/common.rs`: `MAX_ROW = 5000`, `MAX_COL = 8`. Rows/columns beyond these are not scanned.

## Test fixtures

`Ref_Files/` (gitignored) contains `.xlsx` fixtures used in `src/main.rs` example comments, e.g. `Test_Twins.xlsx`, `Test_Excell_T1.xlsx`, `Test_Twins_Analysis.xlsx`, `Test_Range_Basic/Merged/Multiline.xlsx`.

## Logging

`fern` logger at `Debug` level to stdout. GUI runs are typically launched as `cargo run --bin gui >OUT`; `OUT` is a captured log file in the repo root (not committed).
