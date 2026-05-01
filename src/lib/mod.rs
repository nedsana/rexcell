use umya_spreadsheet::*;
use std::collections::HashMap;
pub mod common;

pub fn column_to_index(col: &str) -> u32 {
    let mut index = 0;
    for c in col.chars() {
        index = index * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    index
}

pub fn index_to_column(mut index: u32) -> String {
    let mut col = String::new();
    while index > 0 {
        index -= 1;
        let remainder = (index % 26) as u8;
        col.push((b'A' + remainder) as char);
        index /= 26;
    }
    col.chars().rev().collect()
}

pub fn get_ref_map_by_indexes(sheet: &Worksheet, col_key: u32, col_value: u32) -> HashMap<String, String> {
    let mut ref_map: HashMap<String, String> = HashMap::new();

    for row in 1..=sheet.get_highest_row() {
        let cell_key = sheet.get_value((col_key, row));
        let cell_value = sheet.get_value((col_value, row));

        if !cell_value.is_empty() && !cell_key.is_empty() {
            ref_map.insert(cell_value.clone(), cell_key.clone());
        }
    }

    ref_map
}

pub fn get_ref_map_by_strings(sheet: &Worksheet, col_key: &String, col_value: &String) -> HashMap<String, String> {
    get_ref_map_by_indexes(sheet, column_to_index(col_key),column_to_index(col_value))
}

pub fn apply_key_value_data_by_indexes(
    sheet: &mut Worksheet,
    extra_sheet: &mut Worksheet,
    ref_map: &HashMap<String, String>,
    src_col: u32,
    dest_col: u32,
) -> Result<(Vec<String>, Vec<String>), String> {
    let mut res = (Vec::new(), Vec::new());
    let max_row = sheet.get_highest_row();
    for row in 1..=max_row {
        let cell_value = sheet.get_value((src_col, row));
        if !cell_value.is_empty() {
            if let Some(value) = ref_map.get(&cell_value) {
                sheet.get_cell_mut((dest_col, row)).set_value(value.clone());

                res.0.push(format!("[Col:{} Raw:{}]: Updated '{}' in '{}'!", 
                        index_to_column(src_col), row, cell_value, sheet.get_name()));
            } else {
                
                res.1.push(format!("[Col:{} Raw:{}]: Unable to find '{}' in '{}'! Adding to sheet {}!", 
                            index_to_column(src_col), row, cell_value, sheet.get_name(), extra_sheet.get_name()));

                let max_col = sheet.get_highest_column();
                let next_row = extra_sheet.get_highest_row() + 1;
                for col in 1..=max_col {
                    let cell_value = sheet.get_value((col, row));
                    if !cell_value.is_empty() {
                        extra_sheet.get_cell_mut((col, next_row)).set_value(cell_value.clone());
                    }
                }
            }
        }
    }

    if res.0.is_empty(){
        Err(common::MESSAGE_NO_KEY_VALUE_MAPPING.to_string())
    } else {
        Ok(res)
    }
}

pub fn apply_key_value_data_by_strings(
    sheet: &mut Worksheet,
    extra_sheet: &mut Worksheet,
    ref_map: &HashMap<String, String>,
    src_col: &String,
    dest_col: &String,
) -> Result<(Vec<String>, Vec<String>), String> {
    apply_key_value_data_by_indexes(sheet, extra_sheet, ref_map, column_to_index(src_col), column_to_index(dest_col))
}

pub fn get_worksheet_names_list(book: &Spreadsheet) -> Vec<String> {
    let sheets = book.get_sheet_collection();
    sheets.iter().map(|s| s.get_name().to_string()).collect()
}

pub fn get_worksheet_names_string(book: &Spreadsheet) -> String {
    get_worksheet_names_list(book).join(",")
}

pub fn get_worksheet_names(path: &std::path::Path) -> String {
    let bk: Spreadsheet = reader::xlsx::read(path).expect(common::ERROR_CANT_READ_FILE);
    get_worksheet_names_string(&bk) 
}

pub fn execute(cfg: &common::Config) -> Result<usize, String> {
    // Load the reference Excel file
    let ref_path = std::path::Path::new(&cfg.ref_file);
    let rbook: Spreadsheet = reader::xlsx::read(ref_path).expect(common::ERROR_CANT_READ_FILE);

    // Load the update Excel file
    let target_path = std::path::Path::new(&cfg.tgt_file);
    let mut ubook: Spreadsheet = reader::xlsx::read(target_path).expect(common::ERROR_CANT_READ_FILE);

    let mut extra_sheet = Worksheet::default();
    extra_sheet.set_name(common::LABEL_NEW_SHEET.to_string());

    // Get the reference sheet
    let rtbl = rbook.get_sheet_by_name(&cfg.ref_table).expect(common::ERROR_REFERENCE_SHEET_NOT_FOUND);

    // Get the key-value entries from the reference table
    use std::collections::HashMap;
    let ref_map: HashMap<String, String> = get_ref_map_by_strings(&rtbl, &cfg.ref_col_key, &cfg.ref_col_value);

    let mut applied: usize = 0;
    for utbln in cfg.tgt_upd_table.split(',') {
        // Get the update sheet
        let utbl = ubook.get_sheet_by_name_mut(&utbln).expect(common::ERROR_UPDATE_SHEET_NOT_FOUND);

        let r = apply_key_value_data_by_strings(utbl, &mut extra_sheet, &ref_map, &cfg.tgt_src_col, &cfg.tgt_dest_col).expect(common::MESSAGE_NO_KEY_VALUE_MAPPING);
        applied += r.0.len(); 
    }

    ubook.add_sheet(extra_sheet).expect(common::ERROR_FAILED_TO_ADD_SHEET);

    if applied > 0 {
        // Save changes
        if cfg.inplace {
            let _ = writer::xlsx::write(&ubook, target_path).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
        } else {
            let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
            let _ = writer::xlsx::write(&ubook, std::path::Path::new(&new_file)).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
        }
    }

    // format!("Applied {} key-value mapping(s).", applied)

    if applied == 0 {
        Err(common::ERROR_NO_ROWS_UPDATED.to_string())
    } else {
        Ok(applied)
    }
}
