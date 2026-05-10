// use clap::error;
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

pub fn get_worksheet_names(path: &std::path::Path) -> Result<String, String> {
    let result = reader::xlsx::read(path);
    match result {
        Ok(bk) => Ok(get_worksheet_names_string(&bk)),
        Err(err) => Err(format!("{}: {}", err, path.display())),
    }   
}

pub fn create_unique_entries_sheet<FRow, FCol, FCell>(
    sheet_in:  &Worksheet, 
    sheet_out: &mut Worksheet,
    filter_row:  Option<FRow>,
    filter_col:  Option<FCol>,
    filter_cell: Option<FCell>,
) -> bool 
where FRow:  Fn(&Worksheet, u32,      &mut Worksheet) -> bool,
      FCol:  Fn(&Worksheet, u32,      &mut Worksheet) -> bool,
      FCell: Fn(&Worksheet, u32, u32, &mut Worksheet) -> bool
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
            let mut added_col = false;
            // Execute per row filter logic, if provided. 
            let mut passes_filter = match &filter_row 
            {
                Some(f) => f(sheet_in, row, sheet_out),
                None => true,
            };

            if passes_filter
            {
                // Copy the data and formatting cell by cell
                for col in 1..=max_col 
                {
                    // Execute per col filter logic, if provided.
                    passes_filter = match &filter_col {
                        Some(f) => f(sheet_in, col, sheet_out),
                        None => true,
                    };

                    if passes_filter 
                    {
                        // Execute per row and col filter logic, if provided.
                        passes_filter = match &filter_cell {
                            Some(f) => f(sheet_in, row, col, sheet_out),
                            None => true,
                        };

                        if passes_filter
                        {
                            let o_src_cell = sheet_in.get_cell((col, row));
                            if let Some(src_cell) = o_src_cell 
                            {
                                let cell_value = src_cell.get_value().clone();
                                let cell_style = src_cell.get_style().clone();

                                let dst_cell = sheet_out.get_cell_mut((col, current_new_row));
                                dst_cell.set_value(cell_value);
                                dst_cell.set_style(cell_style);
                                added_col = true;

                                // Copy column width if defined
                                let o_col_dim = sheet_in.get_column_dimension_by_number(&col);
                                if let Some(col_dim) = o_col_dim 
                                {
                                    let col_width = col_dim.get_width().clone();
                                    sheet_out.get_column_dimension_by_number_mut(&col).set_width(col_width);
                                }
                            }
                            else
                            {
                                added_col = false;
                            }
                        }
                    }
                }

                if added_col
                {
                    // Copy row height if defined
                    let o_row_dim = sheet_in.get_row_dimension(&row);
                    if let Some(row_dim) = o_row_dim 
                    {
                        let row_height = row_dim.get_height().clone();
                        sheet_out.get_row_dimension_mut(&current_new_row).set_height(row_height);
                    }
                }

                current_new_row += 1;
            }
        }
        else
        {
            println!("Row {} is part of a merged cell, skipping!", row);
        }
    }
    true
}

pub fn filter_sheet_by_col_and_accum(
    sheet_in:  &Worksheet, 
    sheet_out: &mut Worksheet,
    col_filter: &String,
    col_accum: &String
) -> bool
{
    let tgt_col = column_to_index(col_filter);
    let quantity_col = column_to_index(col_accum);

    create_unique_entries_sheet(sheet_in, sheet_out, Some(|sheet_in: &Worksheet, row: u32, sheet_out: &mut Worksheet| 
        {
            let o_src_cell = sheet_in.get_cell((tgt_col, row));
            if let Some(src_cell) = o_src_cell 
            {
                let src_cell_value = src_cell.get_value();
                // println!("======================================");
                // Check if the value already exists in the output sheet
                let max_row_out = sheet_out.get_highest_row();
                for row_out in 1..=max_row_out 
                {
                    let o_dst_cell = sheet_out.get_cell((tgt_col, row_out));

                    if let Some(dst_cell) = o_dst_cell 
                    {
                        let dst_cell_value = dst_cell.get_value();

                        if dst_cell_value == src_cell_value 
                        {
                            // println!("  <FOUND> DST [row:{} col:{}] '{}' <-> SRC [row:{} col:{}] '{}'", row_out, tgt_col, dst_cell_value, row, tgt_col, src_cell_value);

                            //the entry is found, but we have to update the cell with quantity
                            let mut q_cell_value_src = 0.0;
                            let o_q_cell_src = sheet_in.get_cell((quantity_col, row));
                            if let Some(q_cell_src) = o_q_cell_src
                            {
                                q_cell_value_src = q_cell_src.get_value().parse::<f32>().unwrap_or(0.0);
                            }

                            let q_cell_dst = sheet_out.get_cell_mut((quantity_col, row_out));
                            let q_cell_value_dst = q_cell_dst.get_value().parse::<f32>().unwrap_or(0.0) + q_cell_value_src;
                            q_cell_dst.set_value(q_cell_value_dst.to_string());

                            return false; // already exists, don't copy
                        }
                        else
                        {
                            // println!("<MISSING> DST [row:{} col:{}] '{}' <-> SRC [row:{} col:{}] '{}'", row_out, tgt_col, dst_cell_value, row, tgt_col, src_cell_value);
                        }
                    }
                }
                return true;
            }
            false
        }),
None::<fn(&Worksheet, u32, &mut Worksheet) -> bool>,
None::<fn(&Worksheet, u32, u32, &mut Worksheet) -> bool>,
    )
}

pub fn execute(cfg: &common::Config) -> Result<(Vec<String>, Vec<String>), String> 
{
    let mut res_error: String = String::new();
    let mut res_success:(Vec<String>, Vec<String>) = (Vec::new(), Vec::new());

    // Load the update Excel file
    let target_path = std::path::Path::new(&cfg.tgt_file);
    let result = reader::xlsx::read(target_path);
    let mut ubook = match result
    {
        Ok(bk) => bk,
        Err(err) => {
            return Err(format!("{}:{} {}", common::ERROR_CANT_READ_FILE, target_path.display(), err));
        }
    };

    let mut extra_sheet = Worksheet::default();
    extra_sheet.set_name(common::LABEL_NEW_SHEET.to_string());

    match cfg.command 
    {
        common::Command::CmdListSheets => 
        {
            let result = get_worksheet_names(std::path::Path::new(&cfg.tgt_file));
            match result 
            {
                Ok(names) => 
                {
                    if names.len() > 0 
                    {
                        res_success.0.push(names);
                    } 
                    else 
                    {
                        return Err(format!("{} {}", common::NO_SHEETS_FOUND, cfg.tgt_file));
                    }
                }
                Err(err) => {
                    return Err(format!("{} {}", common::ERROR_CANT_READ_FILE, err));
                }
            }
        },

        common::Command::CmdFilterSheets => 
        {
            let tgt_col = column_to_index(&cfg.tgt_src_col);
            let quantity_col = tgt_col + 2; //think how to pass it as a parameter

            for utbln in cfg.tgt_upd_table.split(',') 
            {
                // Get the update sheet
                let result = ubook.get_sheet_by_name_mut(&utbln);
                let utbl = match result
                {
                    Some(tbl) => tbl,
                    None => {
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };
                // create_unique_entries_sheet::<fn(&Worksheet, u32, u32, &mut Worksheet) -> bool>(utbl, &mut extra_sheet, None);
                let r = filter_sheet_by_col_and_accum(utbl, &mut extra_sheet, &cfg.tgt_src_col, &index_to_column(quantity_col));
                if !r 
                {
                    res_error = format!("{}:{}", common::ERROR_FAILED_FILTER_SHEET, utbln);
                    break;
                }
                else
                {
                    res_success.0.push(format!("{} '{}'", common::FILTERED_SHEET, utbln));
                }
            }
        },

        common::Command::CmdUpdateSheets => 
        {
            // Load the reference Excel file
            let ref_path = std::path::Path::new(&cfg.ref_file);
            let result = reader::xlsx::read(ref_path);
            let mut rbook = match result
            {
                Ok(bk) => bk,
                Err(err) => {
                    return Err(format!("{}:{} {}", common::ERROR_CANT_READ_FILE, ref_path.display(), err));
                }
            };        

            // Get the reference sheet
            let result = rbook.get_sheet_by_name_mut(&cfg.ref_table);
            let rtbl = match result
            {
                Some(tbl) => tbl,
                None => {
                    return Err(format!("{}:{}", common::ERROR_REFERENCE_SHEET_NOT_FOUND, cfg.ref_table));
                }
            };

            // Get the key-value entries from the reference table
            use std::collections::HashMap;
            let ref_map: HashMap<String, String> = get_ref_map_by_strings(&rtbl, &cfg.ref_col_key, &cfg.ref_col_value);

            for utbln in cfg.tgt_upd_table.split(',') 
            {
                // Get the update sheet
                let result = ubook.get_sheet_by_name_mut(&utbln);
                let utbl = match result
                {
                    Some(tbl) => tbl,
                    None => {
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };

                let result = apply_key_value_data_by_strings(utbl, 
                                                                                            &mut extra_sheet, 
                                                                                            &ref_map, 
                                                                                            &cfg.tgt_src_col, 
                                                                                            &cfg.tgt_dest_col);
                let r = match result {
                    Ok(r) => r,                
                    Err(e) => {
                        return Err(format!("{}:{}", common::MESSAGE_NO_KEY_VALUE_MAPPING, e));
                    }
                };

                res_success.0.extend(r.0);
                res_success.1.extend(r.1); 
            }
        },

        _ => 
        {
            res_error = format!("{}:{:?}", common::ERROR_INVALID_COMMAND, cfg.command);
        },
    }

    if cfg.command == common::Command::CmdFilterSheets || cfg.command == common::Command::CmdUpdateSheets
    {
        //Add the extra sheet to the book
        let result = ubook.add_sheet(extra_sheet);
        if let Err(err) = result
        {
            return Err(format!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err));
        }; 
    }

    // Save the changes if there are any successful updates, otherwise return the error message
    if res_success.0.len() > 0 
    {
        if cfg.command == common::Command::CmdFilterSheets || cfg.command == common::Command::CmdUpdateSheets
        {
            // Save changes
            if cfg.inplace 
            {
                let result = writer::xlsx::write(&ubook, target_path);
                if let Err(err) = result 
                {
                    return Err(format!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, target_path.display(), err));
                }
            } 
            else 
            {
                let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                let result = writer::xlsx::write(&ubook, std::path::Path::new(&new_file));
                if let Err(err) = result 
                {
                    return Err(format!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, new_file, err));
                }
            }
        }
        Ok(res_success)
    }
    else 
    {
        Err(format!("{} {}", common::ERROR_NO_ROWS_UPDATED.to_string(), res_error))
    }
}
