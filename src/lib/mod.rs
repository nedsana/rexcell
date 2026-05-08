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
                
                res.1.push(format!("[Col:{} Raw:{}]: Can't find '{}' in '{}'! Adding to sheet '{}'!", 
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

pub fn create_unique_entries_sheet<F>(
    sheet_in:  &Worksheet, 
    sheet_out: &mut Worksheet,
    filter: Option<F>,
) where F: Fn(&Worksheet, u32, u32, &mut Worksheet) -> bool
{
    let sheet_in_merged_cells = sheet_in.get_merge_cells(); 

    let max_row = sheet_in.get_highest_row();
    let max_col = sheet_in.get_highest_column();
    let mut current_new_row = 1;

    for row in 1..=max_row 
    {
        // Are there any merged cells that include this row?
        let is_merged = sheet_in_merged_cells.iter().any(|range| {
            let start_row = range.get_coordinate_start_row().unwrap();
            let end_row = range.get_coordinate_end_row().unwrap();
            row >= *start_row.get_num() && row <= *end_row.get_num()
        });

        if !is_merged 
        {
            // Execute the filter logic if provided
            let passes_filter = match &filter {
                Some(f) => f(sheet_in, row, 0, sheet_out),
                None => true,
            };

            if passes_filter
            {
                // Copy the data and formatting cell by cell
                for col in 1..=max_col 
                {
                    let o_src_cell = sheet_in.get_cell((col, row));
                    if let Some(src_cell) = o_src_cell 
                    {
                        let cell_value = src_cell.get_value().clone();
                        let cell_style = src_cell.get_style().clone();

                        let dst_cell = sheet_out.get_cell_mut((col, current_new_row));
                        dst_cell.set_value(cell_value);
                        dst_cell.set_style(cell_style);

                        // Copy column width if defined
                        let o_col_dim = sheet_in.get_column_dimension_by_number(&col);
                        if let Some(col_dim) = o_col_dim 
                        {
                            let col_width = col_dim.get_width().clone();
                            sheet_out.get_column_dimension_by_number_mut(&col).set_width(col_width);
                        }
                    }
                }

                // Copy row height if defined
                let o_row_dim = sheet_in.get_row_dimension(&row);
                if let Some(row_dim) = o_row_dim 
                {
                    let row_height = row_dim.get_height().clone();
                    sheet_out.get_row_dimension_mut(&current_new_row).set_height(row_height);
                }

                current_new_row += 1;
            }
        }
        else
        {
            println!("Row {} is part of a merged cell, skipping!", row);
        }
    }
}

pub fn execute(cfg: &common::Config) -> Result<(Vec<String>, Vec<String>), String> {
    // Load the update Excel file
    let target_path = std::path::Path::new(&cfg.tgt_file);
    let mut ubook: Spreadsheet = reader::xlsx::read(target_path).expect(common::ERROR_CANT_READ_FILE);

    let mut extra_sheet = Worksheet::default();
    extra_sheet.set_name(common::LABEL_NEW_SHEET.to_string());

    let mut res:(Vec<String>, Vec<String>) = (Vec::new(), Vec::new());

    if cfg.ref_file.is_empty() 
    {
        let tgt_col = column_to_index(&cfg.tgt_src_col);

        for utbln in cfg.tgt_upd_table.split(',') {
            // Get the update sheet
            let utbl = ubook.get_sheet_by_name_mut(&utbln).expect(common::ERROR_UPDATE_SHEET_NOT_FOUND);

            // create_unique_entries_sheet::<fn(&Worksheet, u32, u32, &mut Worksheet) -> bool>(utbl, &mut extra_sheet, None);

            create_unique_entries_sheet(utbl, &mut extra_sheet, Some(|sheet_in: &Worksheet, row: u32, _col: u32, sheet_out: &mut Worksheet| 
                {
                    let o_src_cell = sheet_in.get_cell((tgt_col, row));
                    if let Some(src_cell) = o_src_cell 
                    {
                        let src_cell_value = src_cell.get_value();

                        // Check if the value already exists in the output sheet
                        let max_row_out = sheet_out.get_highest_row();
                        for row_out in 1..=max_row_out 
                        {
                            let o_dst_cell = sheet_out.get_cell((tgt_col, row_out));

                            if let Some(dst_cell) = o_dst_cell 
                            {
                                let dst_cell_value = dst_cell.get_value();

                                // println!("DST [row:{} col:{}] '{}' <-> SRC [row:{} col:{}] '{}'", row_out, tgt_col, dst_cell_value, row, tgt_col, src_cell_value);

                                if dst_cell_value == src_cell_value 
                                {
                                    return false; // already exists, don't copy
                                }
                            }
                        }
                        return true;
                    }
                    false
                })
            ); //create_unique_entries_sheet

        }

        res.0.push("SOME DUMMY CONTENT!".to_string());

        
        /*
        let col_id = column_to_index(&cfg.tgt_src_col);
        for utbln in cfg.tgt_upd_table.split(',') 
        {
            // Get the update sheet
            let utbl = ubook.get_sheet_by_name_mut(&utbln).expect(common::ERROR_UPDATE_SHEET_NOT_FOUND);
            let max_row = utbl.get_highest_row();
            for row in 1..=max_row 
            {
                println!("{}: Row:{} has {} columns!", utbln, row, utbl.get_highest_column().to_string());

                let cell_value = utbl.get_value((col_id, row));
                if cell_value.len() > 0 
                {
                    // find cell_value in extra_sheet and if not found, copy the whole row to extra_sheet
                    let mut found = false;
                    let e_max_row = extra_sheet.get_highest_row();
                    if 0 != e_max_row
                    {
                        for erow in 1..=extra_sheet.get_highest_row() 
                        {
                            let cell_key = extra_sheet.get_value((col_id, erow));
                            if cell_key.len() > 0 && cell_value == cell_key 
                            {
                                let r = format!("{}:[Row:{} Col:{}]: Already in extra sheet:{}", extra_sheet.get_name(), erow, col_id, cell_key);
                                res.1.push(r);

                                found = true;
                                break; // found the value in extra_sheet, break the loop
                            }
                        }
                    }
                    
                    if !found
                    {
                        let next_row = e_max_row+1;
                        extra_sheet.get_cell_mut((col_id, next_row)).set_value(cell_value.clone());

                        let r = format!("{}:[Row:{} Col:{}]: adding to extra sheet:{}", utbln, next_row, col_id, cell_value);
                        res.0.push(r);
                    }
                }
            }
        }
        */
    }
    else
    {
        // Load the reference Excel file
        let ref_path = std::path::Path::new(&cfg.ref_file);
        let rbook: Spreadsheet = reader::xlsx::read(ref_path).expect(common::ERROR_CANT_READ_FILE);

        // Get the reference sheet
        let rtbl = rbook.get_sheet_by_name(&cfg.ref_table).expect(common::ERROR_REFERENCE_SHEET_NOT_FOUND);

        // Get the key-value entries from the reference table
        use std::collections::HashMap;
        let ref_map: HashMap<String, String> = get_ref_map_by_strings(&rtbl, &cfg.ref_col_key, &cfg.ref_col_value);

        for utbln in cfg.tgt_upd_table.split(',') {
            // Get the update sheet
            let utbl = ubook.get_sheet_by_name_mut(&utbln).expect(common::ERROR_UPDATE_SHEET_NOT_FOUND);

            let r = apply_key_value_data_by_strings(utbl, 
                                                                                &mut extra_sheet, 
                                                                                &ref_map, 
                                                                                &cfg.tgt_src_col, 
                                                                                &cfg.tgt_dest_col).expect(common::MESSAGE_NO_KEY_VALUE_MAPPING);
            res.0.extend(r.0);
            res.1.extend(r.1); 
        }
    }

    ubook.add_sheet(extra_sheet).expect(common::ERROR_FAILED_TO_ADD_SHEET);

    if res.0.len() > 0 {
        // Save changes
        if cfg.inplace {
            let _ = writer::xlsx::write(&ubook, target_path).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
        } else {
            let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
            let _ = writer::xlsx::write(&ubook, std::path::Path::new(&new_file)).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
        }
        Ok(res)
    }
    else {
        Err(common::ERROR_NO_ROWS_UPDATED.to_string())
    }
}
