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
    ref_map: &HashMap<String, String>,
    src_col: u32,
    dest_col: u32,
) -> Result<usize, String> {
    let mut applied = 0;

    let max_col = sheet.get_highest_row();
    for row in 1..=max_col {
        let cell_value = sheet.get_value((src_col, row));

        if !cell_value.is_empty() {
            if let Some(value) = ref_map.get(&cell_value) {
                sheet.get_cell_mut((dest_col, row)).set_value(value.clone());
                applied += 1;
            } else {
                println!("[Col:{} Raw:{}]: Unable to find '{}' in '{}'!", index_to_column(src_col), row, cell_value, sheet.get_name());

                // Print all available values in the source column for this row. Lots of them are empty!
                // let values: Vec<String> = (1..=max_col)
                //     .map(|col| sheet.get_value((col, row)).to_string()).collect();
                // println!("Available values in row {}: {}", row, values.join(","));
            }
        }
    }

    if applied == 0 {
        Err(common::MESSAGE_NO_KEY_VALUE_MAPPING.to_string())
    } else {
        Ok(applied)
    }
}

pub fn apply_key_value_data_by_strings(
    sheet: &mut Worksheet,
    ref_map: &HashMap<String, String>,
    src_col: &String,
    dest_col: &String,
) -> Result<usize, String> {
    apply_key_value_data_by_indexes(sheet, ref_map, column_to_index(src_col), column_to_index(dest_col))
}

pub fn get_worksheet_names_list(book: &Spreadsheet) -> Vec<String> {
    let sheets = book.get_sheet_collection();
    sheets.iter().map(|s| s.get_name().to_string()).collect()
}

pub fn get_worksheet_names_string(book: &Spreadsheet) -> String {
    get_worksheet_names_list(book).join(",")
}
