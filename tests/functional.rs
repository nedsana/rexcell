use std::fs;
use std::path::{Path, PathBuf};

use rexcell::{common, excell, range_ops};

static LOGGER: std::sync::Once = std::sync::Once::new();

fn enable_logging() {
    if std::env::var_os("TEST_LOGGING").is_none() { return; }
    LOGGER.call_once(|| {
        let _ = fern::Dispatch::new()
            .format(|out, message, _record| out.finish(*message))
            .level(log::LevelFilter::Debug)
            .chain(std::io::stdout())
            .apply();
    });
}

fn create_folder(folder: &str)
{
    fs::create_dir_all(folder)
        .unwrap_or_else(|e| panic!("Failed to create '{}': {}", folder, e));
}

// cargo runs integration tests with CWD = crate root
fn reference_file(name: &str) -> PathBuf {
    let dir = "tests/fixtures/reference";
    // tests/fixtures/reference is part of the project, so if it is missing there 
    // is something wrong with the git repo access. Don't create the folder!
    // create_folder(dir); 
    Path::new(dir).join(name)
}

fn expected_file(name: &str) -> PathBuf {
    let dir = "tests/fixtures/expected";
    create_folder(dir);
    Path::new(dir).join(name)
}

fn actual_file(name: &str) -> PathBuf {
    let dir = "tests/fixtures/actual";
    create_folder(dir);
    Path::new(dir).join(name)
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

/// Execute rexcell on `cfg` (whose `tgt_file` must already live inside the actual dir) 
/// and return the path of the produced `_new.xlsx`.
fn run(cfg: &mut common::Config) -> Result<PathBuf, String>
{
    let Some(stem) = Path::new(&cfg.tgt_file).file_stem().and_then(|s| s.to_str()) else
    {
        return Err(format!("run(): Failed to read target file name from '{}'!", cfg.tgt_file));
    };

    let out = actual_file(&format!("{}_new.xlsx", stem));
    let _ = fs::remove_file(&out); // drop the old test file

    excell::execute(cfg).map_err(|e| format!("execute failed: {}", e))?;

    if !out.exists() {
        return Err(format!("run(): Missing: {}", out.display()));
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

//Execute functional tests for "Autocomplete" command over Test_Excell.*.xlsx and Test_Range_.*.xlsx files.
fn test_autocomplete(afile_name: &str, tfile_name: &str) -> bool
{
    let mode_record = std::env::var("MODE_RECORD").is_ok();
    let mut ret = false;

    let afile_path = reference_file(afile_name);
    let mut tfile_path = reference_file(tfile_name);

    if afile_path.exists() && tfile_path.exists()
    {
        let act_file_path = actual_file(tfile_name);
        let exp_file_path = expected_file(tfile_name);

        let afile = afile_path.to_string_lossy().into_owned();
        let efile = exp_file_path.to_string_lossy().into_owned();

        if let Err(err) = fs::copy(&tfile_path, &act_file_path)
        {
            eprintln!("Failed to copy '{}' to '{}'! {}", tfile_path.as_path().to_string_lossy(), act_file_path.as_path().to_string_lossy(), err);
            return ret;
        }

        tfile_path = act_file_path; //the file we will be working on

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
                        // Update reference/golden is written only after a successful run,
                        if let Err(err) = fs::copy(Path::new(&tfile), &exp_file_path)
                        {
                            eprintln!("Failed to copy '{}' to '{}'! {}", tfile, exp_file_path.as_path().to_string_lossy(), err);
                        }
                        else
                        {
                            println!("Created {}!", exp_file_path.as_path().to_string_lossy());
                            ret = true;
                        }
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

/* ----------------------------------------------------------------------------------------------- */

#[test]
fn corner_case() 
{
    enable_logging();

    let path = actual_file("corner_case.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((3, 2)).set_value("бетон "); // different case + trailing space
        r.get_cell_mut((4, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    let out = run(&mut update_cfg(&path, false)).expect("execute failed");
    let book = umya_spreadsheet::reader::xlsx::read(&out).expect("read output");
    assert_eq!(sheet_value(&book, "S1", 4, 2), "123");
}

#[test]
fn corner_no_match() 
{
    enable_logging();

    let path = actual_file("corner_no_match.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((3, 2)).set_value("Няма"); //missing entry
        r.get_cell_mut((4, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    let err = run(&mut update_cfg(&path, false)).expect_err("must fail");
    assert!(
        err.contains(common::MESSAGE_NO_KEY_VALUE_MAPPING),
        "unexpected error: {err}"
    );
}

#[test]
fn corner_numeric() 
{
    enable_logging();

    let path = actual_file("corner_numeric.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Бетон");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((3, 2)).set_value("Бетон");
        r.get_cell_mut((4, 2)).set_value_number(12.5);
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

#[test]
fn corner_inplace() 
{
    enable_logging();

    let path = actual_file("corner_inplace.xlsx");
    write_workbook(&path, |book| {
        let mut s1 = umya_spreadsheet::Worksheet::default();
        s1.set_name("S1");
        s1.get_cell_mut((3, 2)).set_value("Ключ");
        let mut r = umya_spreadsheet::Worksheet::default();
        r.set_name("R");
        r.get_cell_mut((3, 2)).set_value("ключ "); // different case + trailing space
        r.get_cell_mut((4, 2)).set_value("123");
        book.add_sheet(s1).unwrap();
        book.add_sheet(r).unwrap();
    });

    // inplace writes back to the input, so execute directly (no run()).
    excell::execute(&update_cfg(&path, true)).expect("execute failed");

    let book = umya_spreadsheet::reader::xlsx::read(&path).expect("read input in place");
    assert_eq!(sheet_value(&book, "S1", 4, 2), "123");
    assert!(
        !actual_file("corner_inplace_new.xlsx").exists(),
        "inplace must not create a _new.xlsx"
    );
}

#[test]
fn corner_filter_accumulate() 
{
    enable_logging();

    let path = actual_file("corner_filter_accumulate.xlsx");
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

// EXPECTED TO FAIL: MAX_ROW boundary bug, keep until fixed!
#[test]
fn corner_max_row() 
{
    enable_logging();
    // Rows 1..4999 share key "K", row 5000 = K5000 (last scanned), rows
    // 5001/5002 = K5001/K5002 must be ignored (MAX_ROW = 5000).
    let path = actual_file("corner_max_row.xlsx");
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

#[test]
fn test_autocomplete_0() 
{
    enable_logging();

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
fn test_autocomplete_1() 
{
    enable_logging();

    let afile_name = "Test_Twins_Analysis.xlsx";
    let ref_files = ["Test_Twins.xlsx"];
    for ref_file in &ref_files {
        if false == test_autocomplete(afile_name, ref_file)
        {
            panic!("Failed to execute test_autocomplete({}, {})!", afile_name, ref_file);
        }
    }
}