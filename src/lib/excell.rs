use std::process;
// use clap::error;
use umya_spreadsheet::*;
use std::collections::HashMap;
use crate::range_types::*;

use super::common;
use super::range_ops;

//to do: make these constants configurable
const MAX_COL: u32 = 8;
const MAX_ROW: u32 = 1000;

pub fn get_ref_map_by_indexes(sheet: &Worksheet, col_key: u32, col_value: u32) -> HashMap<String, String> {
    let mut ref_map: HashMap<String, String> = HashMap::new();

    for row in 1..=MAX_ROW /*sheet.get_highest_row()*/ {
        let cell_key = sheet.get_value((col_key, row));
        let cell_value = sheet.get_value((col_value, row));

        if !cell_value.is_empty() && !cell_key.is_empty() {
            ref_map.insert(cell_value.clone(), cell_key.clone());
        }
    }

    ref_map
}

pub fn get_ref_map_by_strings(sheet: &Worksheet, col_key: &String, col_value: &String) -> HashMap<String, String> {
    get_ref_map_by_indexes(sheet, range_ops::column_to_index(col_key),range_ops::column_to_index(col_value))
}

pub fn apply_formulas(
    rtbl: &Worksheet,
    utbl: &mut Worksheet,
    col_key: u32,
)
{
    let utbl_max_row = MAX_ROW; //utbl.get_highest_row();
    let rtbl_max_row = MAX_ROW; //rtbl.get_highest_row();

    for rtbl_row in 1..=rtbl_max_row //loop over the reference table rows
    {
        let rtbl_key_value = rtbl.get_value((col_key, rtbl_row));
        
        if !rtbl_key_value.is_empty()
        {
            for utbl_row in 1..=utbl_max_row //loop over the update table rows
            {
                let utbl_key_value = utbl.get_value((col_key, utbl_row));

                if !utbl_key_value.is_empty() && range_ops::cmp_strs(&utbl_key_value, &rtbl_key_value) 
                {
                    // let utbl_name = utbl.get_name().to_string();
                    let utbl_max_col = MAX_COL; //utbl.get_highest_column();
                    for utbl_col in 1..=utbl_max_col
                    {
                        let ucell = utbl.get_cell_mut((utbl_col, utbl_row));
                        if ucell.is_formula()
                        {
                            let formula: String = ucell.get_formula().to_string();
                            // println!("Found formula({}) in '{} {}{}'", formula, utbl_name, index_to_column(utbl_col), utbl_row);
                            ucell.set_value("");
                            ucell.set_formula(formula);
                        }
                    }
                }
            }
        }
    }
}

pub fn reset_formulas(
    utbl: &mut Worksheet,
)
{
    let utbl_max_row = MAX_ROW; //utbl.get_highest_row();
    let utbl_max_col = MAX_COL; //utbl.get_highest_column();
    for utbl_col in 1..=utbl_max_col //loop over the update table rows
    {
        for utbl_row in 1..=utbl_max_row //loop over the update table rows
        {
            let ucell = utbl.get_cell_mut((utbl_col, utbl_row));
            if ucell.is_formula()
            {
                let formula: String = ucell.get_formula().to_string();
                // println!("Found formula({}) in '{} {}{}'", formula, utbl_name, index_to_column(utbl_col), utbl_row);
                ucell.set_value("");
                ucell.set_formula(formula);
            }
        }
    }
}

pub fn apply_key_value_data_by_indexes(
    rtbl: &Worksheet,
    utbl: &mut Worksheet,
    col_key: u32,
    col_upd: u32,
) -> Result<(Vec<String>, Vec<String>), String> 
{
    // println!("apply_key_value_data_by_indexes(rtbl:{} utbl:{} col_key:{} col_upd:{})", rtbl.get_name(), utbl.get_name(), col_key, col_upd);

    let mut res = (Vec::new(), Vec::new());
    
    let utbl_max_row = MAX_ROW; //utbl.get_highest_row();
    let rtbl_max_row = MAX_ROW; //rtbl.get_highest_row();

    for utbl_row in 1..=utbl_max_row //loop over the update table rows
    {
        let utbl_key_value = utbl.get_value((col_key, utbl_row));

        if !utbl_key_value.is_empty() 
        {
            let mut found = false;
            
            for rtbl_row in 1..=rtbl_max_row //loop over the reference table rows
            {
                let rtbl_key_value = rtbl.get_value((col_key, rtbl_row));

                if range_ops::cmp_strs(&utbl_key_value, &rtbl_key_value) 
                {
                    let rtbl_upd_value = rtbl.get_value((col_upd, rtbl_row));
                    let rtbl_upd_cell = rtbl.get_cell((col_upd, rtbl_row));

                    if let Some(upd_cell) = rtbl_upd_cell 
                    {
                        let dst_cell = utbl.get_cell_mut((col_upd, utbl_row));

                        // println!("dst_cell({}{}).get_data_type()={}", index_to_column(col_upd), utbl_row, upd_cell.get_data_type());

                        if upd_cell.get_data_type() == "n" && let Some(num) = upd_cell.get_value_number()
                        {
                            // println!("dst_cell({}{}).set_value_number({})", index_to_column(col_upd), utbl_row, num);
                            dst_cell.set_value_number(num);
                        } 
                        else 
                        {
                            // println!("dst_cell({}{}).set_value({})", index_to_column(col_upd), utbl_row, rtbl_upd_value);
                            dst_cell.set_value(rtbl_upd_value.clone());
                        }
                    } 
                    else 
                    {
                        utbl.get_cell_mut((col_upd, utbl_row)).set_value(rtbl_upd_value.clone());
                    }

                    res.0.push(format!("Updated '{} {}{}' with '{}' from '{} {}{}'!", 
                                        utbl.get_name(), range_ops::index_to_column(col_upd), utbl_row, rtbl_upd_value,
                                        rtbl.get_name(), range_ops::index_to_column(col_upd), rtbl_row));

                    found = true;
                    
                    break;
                }
            }

            if !found
            {
                res.1.push(format!("Can't find '{} {}{}' '{}' in '{}'!", utbl.get_name(), range_ops::index_to_column(col_upd), 
                                    utbl_row, utbl_key_value, rtbl.get_name()));
            }
        }
    }

    if res.0.is_empty()
    {
        Err(common::MESSAGE_NO_KEY_VALUE_MAPPING.to_string())
    } 
    else 
    {
        reset_formulas(utbl);
        Ok(res)
    }
}

pub fn apply_key_value_data_by_strings(
    rtbl: &Worksheet,
    utbl: &mut Worksheet,
    col_key: &String,
    cols_upd: &String,
) -> Result<(Vec<String>, Vec<String>), String> 
{
    if cols_upd.len() == 0 
    {
        return Err(common::ERROR_DEST_COL_NOT_DEFINED.to_string());
    }

    let mut res = (Vec::new(), Vec::new());
    for col_upd in cols_upd.split(',') 
    {
        let result = apply_key_value_data_by_indexes(rtbl, utbl, 
                                                                range_ops::column_to_index(col_key), 
        range_ops::column_to_index(col_upd));

        match result {
            Ok((mut updated, mut not_found)) => 
            {
                res.0.append(&mut updated);
                res.1.append(&mut not_found);   
            }
            Err(err) => 
            {
                return Err(format!("{}", err));
            }
        }  
    }
    Ok(res)
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

/**
 * Find a matching range in the sheet.
 * Returns Some(range) when a matching range is found, otherwise None.
 */
fn find_range_in_sheet<'a>(range: &'a dyn IRange, sheet: &'a Worksheet, cmp_cols: &'a Vec<u32>) -> Option<RangeType<'a>>
{
    let max_row = MAX_ROW; //sheet.get_highest_row();
    let max_col = MAX_COL; //sheet.get_highest_column();
    let iter_sheet = range_ops::IterRow::new(sheet, max_row, max_col);
    for it in iter_sheet 
    {
        if it.compare_range(range, false, None, Some(cmp_cols.clone()))
        {
            return Some(it);
        }
    }
    None
}

/**
 * Filter the table. Collect only unique items found in col_filter and accumulate the content from col_accum.
 * sheet_in: source sheet, from which we read
 * sheet_out: destination sheet, to which we write
 * col_filter: the filtering is based on the content of this column
 * col_accum: when, we find item in col_filter, which is aleady present in sheet_out, we accumulate the data from col_accum
 */
pub fn filter_sheet_by_col_and_accum(
    sheet_in:  &Worksheet, 
    sheet_out: &mut Worksheet,
    col_filter: &String,
    cols_accum: &String
) -> bool 
{
    let mut res = false;

    let cmp_cols: Vec<u32> = col_filter.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();
    let acc_cols: Vec<u32> = cols_accum.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();

    let max_row = MAX_ROW; //sheet_in.get_highest_row();
    let max_col = MAX_COL; //sheet_in.get_highest_column();

    // let merged_cells = sheet_in.get_merge_cells();

    let iter_sheet = range_ops::IterRow::new(sheet_in, max_row, max_col);
    
    println!("[filter_sheet_by_col_and_accum] Staring loop: for it in iter_sheet!");

    for it in iter_sheet 
    {
        let mut it_range = it.get_range().clone();

        println!("[filter_sheet_by_col_and_accum] Processing range '{}:[{}]'!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));

        if let Some(found_range) = find_range_in_sheet(&it, sheet_out, &cmp_cols)
        { //accumulating
            let found_range_clone = found_range.get_range().clone();
            drop(found_range);

            println!("[filter_sheet_by_col_and_accum] Range {} already exists in sheet {}! Accumulating data!", range_ops::range_to_string(it.get_range()), sheet_out.get_name());

            if range_ops::accumulate_ranges(sheet_in, it.get_range(), sheet_out, &found_range_clone, None, Some(acc_cols.clone()))
            {
                println!("[filter_sheet_by_col_and_accum] Accumulated in-range '{}' to out-range '{}'!", range_ops::range_to_string(it.get_range()), range_ops::range_to_string(&found_range_clone));
            }
        }
        else
        { //appending
            if let RangeType::Multiline(_) = it 
            {
                println!("[filter_sheet_by_col_and_accum] >>>>>> Range {} does not exist in sheet {}! Find largest multiline section! <<<<<<<", range_ops::range_to_string(it.get_range()), sheet_out.get_name());
                
                let (brow_it, erow_it, bcol_it, ecol_it, rows_it, cols_it) = range_ops::range_bounds(it.get_range());

                // Find the largest multiline range in 'sheet_in', which matches the first line of the current range.
                let multiline_iter_sheet = range_ops::IterRow::new(sheet_in, max_row, max_col);
                for mlit in multiline_iter_sheet 
                {
                    if let RangeType::Multiline(_) = mlit
                    {
                        let (brow_mlit, erow_mlit, bcol_mlit, ecol_mlit, rows_mlit, cols_mlit) = range_ops::range_bounds(mlit.get_range());
                        if (brow_it == brow_mlit) && (erow_it == erow_mlit) && (bcol_it == bcol_mlit) && (ecol_it == ecol_mlit) && (rows_it == rows_mlit) && (cols_it == cols_mlit)
                        {
                            println!("[filter_sheet_by_col_and_accum] Inspecting same multiline range {} in sheet {}! Skipping!", range_ops::range_to_string(mlit.get_range()), sheet_out.get_name());
                        }
                        else
                        {
                            let it_flr   = range_ops::make_range_from_indexes(bcol_it,   brow_it,     ecol_it, brow_it);
                            let mlit_flr: Range = range_ops::make_range_from_indexes(bcol_mlit, brow_mlit, ecol_mlit, brow_mlit);
                            
                            println!("[filter_sheet_by_col_and_accum] Comparing first line range {} with multiline range {} in sheet {}!", 
                                range_ops::range_to_string(&it_flr), 
                                range_ops::range_to_string(&mlit_flr), 
                                sheet_out.get_name());

                            if range_ops::comapre_ranges(it.get_sheet(), &it_flr, mlit.get_sheet(), &mlit_flr, false, None, Some(cmp_cols.clone())) && 
                                rows_it < rows_mlit
                            {
                                println!("[filter_sheet_by_col_and_accum] Found matching multiline range {} in sheet {} with more rows!", range_ops::range_to_string(mlit.get_range()), sheet_out.get_name());

                                it_range = mlit.get_range().clone();
                            }
                        }
                    }
                }

                //append the largest multiline range to the output sheet
                res = range_ops::append_range(sheet_in, &it_range, sheet_out);

                println!("[filter_sheet_by_col_and_accum] Appended range {}:[{}] to {}: {}", sheet_in.get_name(), range_ops::range_to_string(&it_range), sheet_out.get_name(), res);

                //reset it_range to the original range for the next iteration
                it_range = it.get_range().clone();
            }
            else
            {
                res = range_ops::append_range(sheet_in, &it_range, sheet_out);

                println!("[filter_sheet_by_col_and_accum] Appended range {}:[{}] to {}: {}", sheet_in.get_name(), range_ops::range_to_string(&it_range), sheet_out.get_name(), res);
            }
        } //appending
        println!("[filter_sheet_by_col_and_accum]========================================================");
        // process::exit(1);
        // return res;
    }
    println!("[filter_sheet_by_col_and_accum] Finisher loop, exiting");
    return res;
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
            return Err(format!("{}:'{}' {}", common::ERROR_CANT_READ_TGT_FILE, target_path.display(), err));
        }
    };

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
                    return Err(format!("{}:'{}' {}", common::ERROR_CANT_READ_TGT_FILE, cfg.tgt_file, err));
                }
            }
        },

        common::Command::CmdFilterSheets => 
        {
            let mut fotbl = Worksheet::default();
            fotbl.set_name(cfg.new_sheet_name.clone());

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

                // Create new table with unique values from cfg.tgt_src_col.When repetition is found, accumulate the values in cfg.tgt_dest_col.
                let r = filter_sheet_by_col_and_accum(utbl, &mut fotbl, &cfg.tgt_src_col, &cfg.tgt_dest_col);
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

            //Add the extra sheet to the book
            let result = ubook.add_sheet(fotbl);
            if let Err(err) = result
            {
                return Err(format!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err));
            }; 
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
                    return Err(format!("{}:'{}' {}", common::ERROR_CANT_READ_REF_FILE, ref_path.display(), err));
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

            // Perform the update for each update sheet
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
                
                let result = apply_key_value_data_by_strings(rtbl, utbl, &cfg.tgt_src_col, &cfg.tgt_dest_col);

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
