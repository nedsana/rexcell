use std::fs;
use std::path::{Path, PathBuf};

use rexcell::{common, excell, range_ops};

// cargo runs integration tests with CWD = crate root
fn reference_file(name: &str) -> PathBuf {
    Path::new("tests/fixtures/reference").join(name)
}

fn expected_file(name: &str) -> PathBuf {
    Path::new("tests/fixtures/expected").join(name)
}

fn actual_file(name: &str) -> PathBuf {
    Path::new("tests/fixtures/actual").join(name)
}


fn actual_dir(command: &str) -> PathBuf {
    Path::new("tests/fixtures/actual").join(command)
}

fn expected_path(command: &str, tgt_base: &str) -> PathBuf {
    Path::new("tests/fixtures/expected")
        .join(format!("{}__{}_new.xlsx", command, tgt_base))
}

fn copy_reference(command: &str, file: &str) -> PathBuf {
    let dir = actual_dir(command);
    fs::create_dir_all(&dir).expect("create actual dir");
    let src = reference_file(file);
    let dst = dir.join(file);
    fs::copy(&src, &dst).expect("copy reference fixture into actual dir");
    dst
}

fn command_str(command: common::Command) -> &'static str {
    match command {
        common::Command::CmdListSheets => "cmd-list-sheets",
        common::Command::CmdFilterSheets => "cmd-filter-sheets",
        common::Command::CmdUpdateSheets => "cmd-update-sheets",
        common::Command::CmdAutocompleteSheets => "cmd-autocomplete-sheets",
        common::Command::CmdUndefined => "cmd-undefined",
    }
}

fn base_config() -> common::Config {
    common::Config {
        command: common::COMMAND_DEFAULT,
        tgt_file: common::TGT_DEFAULT_EXCEL_FILE.to_string(),
        tgt_upd_table: common::TGT_DEFAULT_TABLE.to_string(),
        tgt_src_col: common::TGT_DEFAULT_SRC_COL.to_string(),
        tgt_acc_col: common::TGT_DEFAULT_ACC_COL.to_string(),
        tgt_dest_col: common::TGT_DEFAULT_DST_COL.to_string(),
        tgt_calcs: String::new(),
        ref_file: common::REF_DEFAULT_EXCEL_FILE.to_string(),
        ref_table: common::REF_DEFAULT_TABLE.to_string(),
        ref_col_key: "B".to_string(),
        ref_col_value: "C".to_string(),
        new_sheet_name: common::TGT_DEFAULT_NEW_SHEET_NAME.to_string(),
        inplace: false,
        analysis_file: common::ANA_DEFAULT_EXCEL_FILE.to_string(),
        analysis_table: common::ANA_DEFAULT_SHEET.to_string(),
        analysis_col_srch: common::ANA_DEFAULT_COL_SRCH.to_string(),
        analysis_col_term: common::ANA_DEFAULT_COL_TERM.to_string(),
        analysis_cols_cp_src: common::ANA_DEFAULT_COLS_CP_SRC.to_string(),
        analysis_srch_pat: common::ANA_DEFAULT_SRCH_PAT.to_string(),
        analysis_term_pat: common::ANA_DEFAULT_TERM_PAT.to_string(),
    }
}

/// Execute rexcell on `cfg` (whose `tgt_file` must already live inside the
/// per-command actual dir) and return the path of the produced `_new.xlsx`.
fn run(cfg: &mut common::Config) -> Result<PathBuf, String> {
    let dir = actual_dir(command_str(cfg.command));
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;

    let tgt = Path::new(&cfg.tgt_file);
    if !tgt.starts_with(&dir) {
        return Err(format!(
            "run(): tgt_file must be inside {}: got {}",
            dir.display(),
            cfg.tgt_file
        ));
    }
    let Some(stem) = tgt.file_stem().and_then(|s| s.to_str()) else {
        return Err(format!("run(): no file stem in tgt_file: {}", cfg.tgt_file));
    };
    let out = dir.join(format!("{}_new.xlsx", stem));
    let _ = fs::remove_file(&out); // drop stale output from earlier runs

    excell::execute(cfg).map_err(|e| format!("execute failed: {}", e))?;
    if !out.exists() {
        return Err(format!("output missing: {}", out.display()));
    }
    Ok(out)
}

/// Cell-by-cell comparison of two workbooks (sheets, then cells up to
/// MAX_ROW/MAX_COL). Never byte-wise: xlsx metadata differs between writes.
fn compare_workbooks(actual: &Path, expected: &Path) -> Result<(), String> {
    let a = umya_spreadsheet::reader::xlsx::read(actual)
        .map_err(|e| format!("read {}: {}", actual.display(), e))?;
    let b = umya_spreadsheet::reader::xlsx::read(expected)
        .map_err(|e| format!("read {}: {}", expected.display(), e))?;

    let a_sheets = a.get_sheet_collection();
    let b_sheets = b.get_sheet_collection();
    let a_names: Vec<String> = a_sheets.iter().map(|s| s.get_name().to_string()).collect();
    let b_names: Vec<String> = b_sheets.iter().map(|s| s.get_name().to_string()).collect();
    if a_names != b_names {
        return Err(format!(
            "mismatch: {} vs {}: sheet names {:?} != {:?}",
            actual.display(),
            expected.display(),
            a_names,
            b_names
        ));
    }

    let mut diffs: Vec<String> = Vec::new();
    'outer: for (asheet, bsheet) in a_sheets.iter().zip(b_sheets.iter()) {
        let name = asheet.get_name().to_string();
        for col in 1..=common::MAX_COL {
            for row in 1..=common::MAX_ROW {
                if diffs.len() >= 10 {
                    break 'outer;
                }
                let acell = asheet.get_cell((col, row));
                let bcell = bsheet.get_cell((col, row));
                let a_formula = acell.map(|c| c.is_formula()).unwrap_or(false);
                let b_formula = bcell.map(|c| c.is_formula()).unwrap_or(false);

                if a_formula || b_formula {
                    let a_f = acell
                        .filter(|c| c.is_formula())
                        .map(|c| c.get_formula().to_string());
                    let b_f = bcell
                        .filter(|c| c.is_formula())
                        .map(|c| c.get_formula().to_string());
                    if a_f != b_f {
                        diffs.push(format!(
                            "{}!{}{}: actual={:?} expected={:?} (formula)",
                            name,
                            range_ops::index_to_column(col),
                            row,
                            a_f,
                            b_f
                        ));
                    }
                    continue;
                }

                let a_val = acell.map(|c| c.get_value().to_string()).unwrap_or_default();
                let b_val = bcell.map(|c| c.get_value().to_string()).unwrap_or_default();
                if a_val.is_empty() && b_val.is_empty() {
                    continue;
                }
                let equal = match (a_val.parse::<f64>(), b_val.parse::<f64>()) {
                    (Ok(an), Ok(bn)) => {
                        (an - bn).abs() <= 1e-6 * 1f64.max(an.abs()).max(bn.abs())
                    }
                    _ => range_ops::cmp_strs(&a_val, &b_val),
                };
                if !equal {
                    diffs.push(format!(
                        "{}!{}{}: actual='{}' expected='{}'",
                        name,
                        range_ops::index_to_column(col),
                        row,
                        a_val,
                        b_val
                    ));
                }
            }
        }
    }

    if diffs.is_empty() {
        Ok(())
    } else {
        Err(format!(
            "mismatch: {} vs {}:\n{}",
            actual.display(),
            expected.display(),
            diffs.join("\n")
        ))
    }
}

/// With REXCELL_RECORD set: record `actual` as the golden. Otherwise compare
/// and panic with the diff message (naming both files) on mismatch.
fn record_or_compare(actual: &Path, expected: &Path) {
    if std::env::var("REXCELL_RECORD").is_ok() {
        if let Some(parent) = expected.parent() {
            fs::create_dir_all(parent).expect("create expected dir");
        }
        fs::copy(actual, expected).expect("record golden file");
        eprintln!("WARNING: REXCELL_RECORD set — recorded {} as golden", expected.display());
        return;
    }
    compare_workbooks(actual, expected).unwrap_or_else(|e| panic!("{}", e));
}

/// Build a workbook in memory via `build` and write it to `path`.
fn write_workbook(path: &Path, build: impl FnOnce(&mut umya_spreadsheet::Spreadsheet)) {
    let mut book = umya_spreadsheet::new_file_empty_worksheet();
    build(&mut book);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).expect("create parent dir for workbook");
    }
    umya_spreadsheet::writer::xlsx::write(&book, path).expect("write workbook");
}

// ----------------------------------------------------------------------
// Corner cases: workbooks built on the fly, specific cells asserted.
// ----------------------------------------------------------------------

fn sheet_value(book: &umya_spreadsheet::Spreadsheet, sheet: &str, col: u32, row: u32) -> String {
    book.get_sheet_by_name(sheet)
        .expect("sheet not found")
        .get_value((col, row))
        .to_string()
}

fn prices_keys(out: &Path) -> Vec<String> {
    let book = umya_spreadsheet::reader::xlsx::read(out).expect("read output");
    let prices = book.get_sheet_by_name("Prices").expect("Prices sheet missing");
    (1..=common::MAX_ROW)
        .map(|r| prices.get_value((3, r)).to_string())
        .filter(|v| !v.is_empty())
        .collect()
}

fn update_cfg(target: &Path, inplace: bool) -> common::Config {
    let mut cfg = base_config();
    cfg.command = common::Command::CmdUpdateSheets;
    cfg.tgt_file = target.to_string_lossy().into_owned();
    cfg.tgt_upd_table = "S1".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_dest_col = "D".to_string();
    cfg.ref_file = target.to_string_lossy().into_owned(); // same workbook
    cfg.ref_table = "R".to_string();
    cfg.ref_col_key = "B".to_string();
    cfg.ref_col_value = "C".to_string();
    cfg.inplace = inplace;
    cfg
}

// #[test]
fn corner_case() {
    let path = actual_dir("corner_case").join("case.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((2, 2)).set_value("бетон "); // different case + trailing space
        r.get_cell_mut((3, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    let out = run(&mut update_cfg(&path, false)).expect("execute failed");
    let book = umya_spreadsheet::reader::xlsx::read(&out).expect("read output");
    assert_eq!(sheet_value(&book, "S1", 4, 2), "123");
}

// #[test]
fn corner_no_match() {
    let path = actual_dir("corner_no_match").join("nomatch.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((2, 2)).set_value("Няма");
        r.get_cell_mut((3, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    let err = run(&mut update_cfg(&path, false)).expect_err("must fail");
    assert!(
        err.contains(common::MESSAGE_NO_KEY_VALUE_MAPPING),
        "unexpected error: {err}"
    );
}

// #[test]
fn corner_numeric() {
    let path = actual_dir("corner_numeric").join("numeric.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((2, 2)).set_value("Бетон");
        r.get_cell_mut((3, 2)).set_value_number(12.5);
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    let out = run(&mut update_cfg(&path, false)).expect("execute failed");
    let book = umya_spreadsheet::reader::xlsx::read(&out).expect("read output");
    let cell = book
        .get_sheet_by_name("S1")
        .unwrap()
        .get_cell((4, 2))
        .expect("D2 missing");
    assert_eq!(cell.get_data_type(), "n");
    assert_eq!(cell.get_value_number(), Some(12.5));
}

// #[test]
fn corner_filter_accumulate() {
    let path = actual_dir("corner_filter_accumulate").join("accum.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((1, 1)).set_value_number(1.0);
        s1.get_cell_mut((3, 1)).set_value("A1");
        s1.get_cell_mut((5, 1)).set_value_number(2.0);
        s1.get_cell_mut((1, 2)).set_value_number(2.0);
        s1.get_cell_mut((3, 2)).set_value("A1");
        s1.get_cell_mut((5, 2)).set_value_number(3.0);
        book.add_sheet(s1).unwrap();
    });

    let mut cfg = base_config();
    cfg.command = common::Command::CmdFilterSheets;
    cfg.tgt_file = path.to_string_lossy().into_owned();
    cfg.tgt_upd_table = "S1".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_acc_col = "E".to_string();
    cfg.new_sheet_name = "Prices".to_string();

    let out = run(&mut cfg).expect("execute failed");
    let book = umya_spreadsheet::reader::xlsx::read(&out).expect("read output");
    let prices = book.get_sheet_by_name("Prices").expect("Prices sheet missing");

    let mut found = 0u32;
    let mut acc = 0.0f64;
    for r in 1..=10u32 {
        if range_ops::cmp_strs(&sheet_value(&book, "Prices", 3, r), "A1") {
            found += 1;
            acc = prices
                .get_cell((5, r))
                .and_then(|c| c.get_value_number())
                .unwrap_or(0.0);
        }
    }
    assert_eq!(found, 1, "expected exactly one accumulated 'A1' row");
    assert!((acc - 5.0).abs() < 1e-9, "expected accumulated 5.0, got {acc}");
}

// #[test]
fn corner_inplace() {
    let dir = actual_dir("corner_inplace");
    let path = dir.join("inplace_in.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Ключ");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((2, 2)).set_value("ключ ");
        r.get_cell_mut((3, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    // inplace writes back to the input, so execute directly (no run()).
    excell::execute(&update_cfg(&path, true)).expect("execute failed");

    let book = umya_spreadsheet::reader::xlsx::read(&path).expect("read input in place");
    assert_eq!(sheet_value(&book, "S1", 4, 2), "123");
    assert!(
        !dir.join("inplace_in_new.xlsx").exists(),
        "inplace must not create a _new.xlsx"
    );
}

// #[test]
fn corner_max_row() {
    // Rows 1..4999 share key "K", row 5000 = K5000 (last scanned), rows
    // 5001/5002 = K5001/K5002 must be ignored (MAX_ROW = 5000).
    let path = actual_dir("corner_max_row").join("big.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        for r in 1..=5002u32 {
            let key = if r < 5000 { "K".to_string() } else { format!("K{r}") };
            s1.get_cell_mut((1, r)).set_value_number(r as f64);
            s1.get_cell_mut((3, r)).set_value(key);
            s1.get_cell_mut((5, r)).set_value_number(1.0);
        }
        book.add_sheet(s1).unwrap();
    });

    let mut cfg = base_config();
    cfg.command = common::Command::CmdFilterSheets;
    cfg.tgt_file = path.to_string_lossy().into_owned();
    cfg.tgt_upd_table = "S1".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_acc_col = "E".to_string();
    cfg.new_sheet_name = "Prices".to_string();

    let out = run(&mut cfg).expect("execute failed");
    let keys = prices_keys(&out);
    assert!(keys.iter().any(|k| k == "K"), "expected accumulated key K in Prices");
    assert!(keys.iter().any(|k| k == "K5000"), "expected K5000 in Prices");
    assert!(!keys.iter().any(|k| k == "K5001"), "K5001 must not be processed (MAX_ROW)");
    assert!(!keys.iter().any(|k| k == "K5002"), "K5002 must not be processed (MAX_ROW)");
}

// #[test]
fn harness_self_check() {
    let p = reference_file("Test_Twins.xlsx");
    compare_workbooks(&p, &p).expect("self-compare must be equal");
}

// ----------------------------------------------------------------------
// Golden tests: run each command over reference fixtures and compare
// the output against committed golden files. Record mode:
// REXCELL_RECORD=1 cargo test --test functional
// ----------------------------------------------------------------------

// #[test]
fn list_sheets_twins() {
    let p = reference_file("Test_Twins.xlsx");
    let names = excell::get_worksheet_names(&p).expect("list sheets");
    assert!(names.contains("Ф200"), "missing sheet 'Ф200' in: {}", names);

    // CmdListSheets writes no file, so execute directly (no run()).
    let mut cfg = base_config();
    cfg.command = common::Command::CmdListSheets;
    cfg.tgt_file = p.to_string_lossy().into_owned();
    excell::execute(&cfg).expect("execute failed");
}

// #[test]
fn filter_twins() {
    let mut cfg = base_config();
    cfg.command = common::Command::CmdFilterSheets;
    cfg.tgt_file = copy_reference("cmd-filter-sheets", "Test_Twins.xlsx")
        .to_string_lossy()
        .into_owned();
    cfg.tgt_upd_table = "Ф200,Ф160,Ф 125,Ф 90".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_acc_col = "E".to_string();
    cfg.tgt_dest_col = "E".to_string();
    cfg.new_sheet_name = "Prices".to_string();

    let out = run(&mut cfg).expect("execute failed");
    record_or_compare(&out, &expected_path("cmd-filter-sheets", "Test_Twins"));
}

// #[test]
fn update_twins() {
    let mut cfg = base_config();
    cfg.command = common::Command::CmdUpdateSheets;
    cfg.tgt_file = copy_reference("cmd-update-sheets", "Test_Twins.xlsx")
        .to_string_lossy()
        .into_owned();
    // Ruling (plan said ref_table "Ед. Цени"): that sheet does not exist in
    // the current Test_Twins.xlsx. The four sheets are twins with identical
    // work names in col C, so sync quantities (col E) from Ф200 to the rest.
    cfg.tgt_upd_table = "Ф160,Ф 125,Ф 90".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_dest_col = "E".to_string();
    // reference file is read in place, never copied
    cfg.ref_file = reference_file("Test_Twins.xlsx").to_string_lossy().into_owned();
    cfg.ref_table = "Ф200".to_string();
    cfg.ref_col_key = "C".to_string();
    cfg.ref_col_value = "E".to_string();

    let out = run(&mut cfg).expect("execute failed");
    record_or_compare(&out, &expected_path("cmd-update-sheets", "Test_Twins"));
}

// #[test]
fn autocomplete_t1() {
    let analysis = reference_file("Test_Excell_T1_Analysis.xlsx");
    let analysis_before = fs::read(&analysis).expect("read analysis file");

    let mut cfg = base_config();
    cfg.command = common::Command::CmdAutocompleteSheets;
    cfg.tgt_file = copy_reference("cmd-autocomplete-sheets", "Test_Excell_T1.xlsx")
        .to_string_lossy()
        .into_owned();
    cfg.tgt_upd_table = "КСС_ОП1".to_string();
    cfg.tgt_src_col = "C".to_string();
    cfg.tgt_dest_col = "B,F".to_string();
    cfg.tgt_acc_col = "E".to_string();
    cfg.tgt_calcs = "G=E*F".to_string();
    cfg.new_sheet_name = "Prices".to_string();
    // analysis file is read in place, never copied
    cfg.analysis_file = analysis.to_string_lossy().into_owned();
    cfg.analysis_table = "Analysis".to_string();
    cfg.analysis_col_srch = "A".to_string();
    cfg.analysis_col_term = "B".to_string();
    cfg.analysis_cols_cp_src = "A,G".to_string();
    cfg.analysis_srch_pat = "Позиция: *, *Основание:(.*)".to_string();
    cfg.analysis_term_pat = "Общо".to_string();

    let out = run(&mut cfg).expect("execute failed");
    assert_eq!(
        fs::read(&analysis).expect("re-read analysis file"),
        analysis_before,
        "analysis file must stay untouched"
    );
    record_or_compare(&out, &expected_path("cmd-autocomplete-sheets", "Test_Excell_T1"));
}





//delete all files from fixtures/expected
fn clear_dir_content(path_dir: &str) -> Result<&str, String>
{
    if let Ok(path_dir) =  fs::read_dir(path_dir)
    {
        for entry in path_dir 
        {
            let path = entry.expect("dir entry").path();
            if path.is_file() {
                fs::remove_file(&path).expect("remove file");
            }
        }
    }
    else 
    {
        return Err(format!("Failed to find folder '{}'! Can't delete content!", path_dir));
    }
    Ok(path_dir)
}

//Execute functional tests for "Autocomplete" command over Test_Excell.*.xlsx and Test_Range_.*.xlsx files.
fn test_autocomplete(afile_name: &str, tfile_name: &str) -> bool
{
    let mut ret = false;

    // let path_ref_dir = "tests/fixtures/reference";
    // let path_exp_dir = "tests/fixtures/expected";

    let afile_path = reference_file(afile_name);
    let mut tfile_path = reference_file(tfile_name);

    if afile_path.exists() && tfile_path.exists()
    {
        let mut act_file_path = actual_file(tfile_name);
        let exp_file_path = expected_file(tfile_name);

        let afile = afile_path.to_string_lossy().into_owned();
        let efile = exp_file_path.to_string_lossy().into_owned();

        let mode_record = std::env::var("MODE_RECORD").is_ok();
        if true == mode_record
        {
            if exp_file_path.exists() && let Err(err) = fs::remove_file(&exp_file_path)
            {
                eprintln!("Failed to delete '{}'! {}", exp_file_path.as_path().to_string_lossy(), err);
                return ret;
            }
            act_file_path = exp_file_path;
        }

        if let Err(err) = fs::copy(&tfile_path, &act_file_path)
        {
            eprintln!("Failed to copy '{}' to '{}'! {}", tfile_path.as_path().to_string_lossy(), act_file_path.as_path().to_string_lossy(), err);
            return ret;
        }
        tfile_path = act_file_path;

        let tfile = tfile_path.to_string_lossy().into_owned();

        if let Ok(assheet) =  umya_spreadsheet::reader::xlsx::read(afile_path)
        {
            let awsheets = excell::get_worksheet_names_list(&assheet);
            if awsheets.len() == 1
            {
                if let Ok(ssheet) = umya_spreadsheet::reader::xlsx::read(tfile_path)
                {
                    let swsheets = excell::get_worksheet_names_string(&ssheet);

                    let mut cfg = base_config();
                    cfg.command = common::Command::CmdAutocompleteSheets;
                    cfg.tgt_file = tfile.clone();
                    cfg.tgt_upd_table = swsheets.clone();
                    cfg.tgt_src_col = "C".to_string();
                    cfg.tgt_dest_col = "B,F".to_string();
                    cfg.tgt_acc_col = "E".to_string();
                    cfg.tgt_calcs = "G=E*F".to_string();
                    cfg.new_sheet_name = "Prices".to_string();
                    // analysis file is read in place, never copied
                    cfg.analysis_file = afile.clone();
                    cfg.analysis_table = awsheets[0].to_string();
                    cfg.analysis_col_srch = "A".to_string();
                    cfg.analysis_col_term = "B".to_string();
                    cfg.analysis_cols_cp_src = "A,G".to_string();
                    cfg.analysis_srch_pat = "Позиция: *, *Основание:(.*)".to_string();
                    cfg.analysis_term_pat = "Общо".to_string();
                    //use inplace update, as we've already set correct path for tgt_file
                    cfg.inplace = true;

                    if let Err(err) = excell::execute(&cfg)
                    {
                        eprintln!("Failed to process file:{} sheets:{}! {}", tfile, swsheets, err);
                    }
                    else if true  == mode_record
                    {
                        println!("Created {}!", tfile);
                        ret = true;
                    }
                    else
                    {
                        if let Err(err) = compare_workbooks(Path::new(&tfile), Path::new(&efile))
                        {
                            eprintln!("[ERROR] {}! {}", tfile, err);
                        }
                        else 
                        {
                            println!("[OK] {}!", tfile);
                            ret = true;    
                        }
                    }
                }
            }
        }
    }
    ret
}

#[test]
fn test_create_autocompleted_files_0() 
{
    let afile_name = "Test_Excell_T1_Analysis.xlsx";
    let ref_files = ["Test_Excell_T0.xlsx", "Test_Excell_ref.xlsx", "Test_Excell_T1_small.xlsx", "Test_Excell_T1.xlsx", 
                                "Test_Range_Basic.xlsx", "Test_Range_Multiline.xlsx", "Test_Range_Merged.xlsx"];
    // let ref_files = ["Test_Excell_T0.xlsx"];
    for ref_file in &ref_files {
        if false == test_autocomplete(afile_name, ref_file)
        {
            panic!("Failed to execute test_autocomplete({}, {})!", afile_name, ref_file);
        }
    }
}

#[test]
fn test_create_autocompleted_files_1() 
{
    let afile_name = "Test_Twins_Analysis.xlsx";
    let ref_files = ["Test_Twins.xlsx"];
    for ref_file in &ref_files {
        if false == test_autocomplete(afile_name, ref_file)
        {
            panic!("Failed to execute test_autocomplete({}, {})!", afile_name, ref_file);
        }
    }
}