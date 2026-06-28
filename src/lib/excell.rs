// use clap::error;
use umya_spreadsheet::*;
use std::collections::HashMap;
use crate::range_types::IRange;
use crate::range_types::IRangeMut;
use crate::lib_impl::range_ops::LendingIterator;

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
 * Copy all rows, which don't contain merged cells, from sheet_in to sheet_out. 
 * Further filtering can be provided via the filter_* arguments.
 * sheet_in: source sheet, from which we read
 * sheet_out: destination sheet, to which we write
 * filter_row: filter lambda, applied per row
 * filter_col: filter lambda, applied per column
 * filter_cell: filter lambda, applied per cell
 */
pub fn create_unique_entries_sheet<FRow, FCol, FCell>(
    sheet_in:  &Worksheet, 
    sheet_out: &mut Worksheet,
    filter_row:  Option<FRow>,
    _filter_col:  Option<FCol>,
    _filter_cell: Option<FCell>,
) -> bool 
where FRow:  Fn(&dyn IRange, &mut Worksheet) -> bool,
      FCol:  Fn(&dyn IRange, &mut Worksheet) -> bool, //is this needed
      FCell: Fn(&dyn IRange, &mut Worksheet) -> bool
{
    let mut res = false;

    let max_row = MAX_ROW; //sheet_in.get_highest_row();
    let max_col = MAX_COL; //sheet_in.get_highest_column();

    let mut current_new_row = sheet_out.get_highest_row()+1;

    let iter_sheet = range_ops::IterRow::new(sheet_in, max_row, max_col);
    let merged_cells = sheet_in.get_merge_cells();
    for it in iter_sheet 
    {
        let it_range = it.get_range();
        let passes_filter_row = match &filter_row 
        {
            Some(f) => f(&it, sheet_out),
            None => true,
        };
/*
        // To do ... how to use these filters? Do we need them at all? Maybe we can just apply them to the whole row/col/cell and not to the range, which is defined by the iterator?
        // let passes_filter_col = match &filter_col 
        // {
        //     Some(f) => f(sheet_in, &it_range, sheet_out),
        //     None => true,
        // };

        // let passes_filter_cell = match &filter_cell 
        // {
        //     Some(f) => f(sheet_in, &it_range, sheet_out),
        //     None => true,
        // };
*/
        if passes_filter_row
        {
            let current_old_row = current_new_row;

            let rbeg = *it_range.get_coordinate_start_row().unwrap().get_num();
            let rend = *it_range.get_coordinate_end_row().unwrap().get_num();
            let cbeg = *it_range.get_coordinate_start_col().unwrap().get_num();
            let cend = *it_range.get_coordinate_end_col().unwrap().get_num();

            let mut added_col = false;

            for row in rbeg..=rend 
            {
                added_col = false;
                for col in cbeg..=cend 
                {
                    //copy to the output sheet all rows, which are defined by the range.
                    if let Some(src_cell) = sheet_in.get_cell((col, row)) 
                    {
                        let o_rich_text = src_cell.get_cell_value().get_raw_value().get_rich_text();
                        let cell_value = src_cell.get_value().clone();
                        let cell_style = src_cell.get_style().clone();
                        let cell_data_type = src_cell.get_data_type().to_string();
                        // let cell_value = src_cell.get_formatted_value().clone();

                        let dst_cell = sheet_out.get_cell_mut((col, current_new_row));
                        
                        // Preserve data types when copying cells
                        if cell_data_type == "n" && let Some(num) = src_cell.get_value_number() 
                        {
                            // println!("dst_cell({}{}).set_value_number({})", range_ops::index_to_column(col), current_new_row, num);
                            dst_cell.set_value_number(num);
                        } 
                        else 
                        {
                            if let Some(rich_text) = o_rich_text 
                            {
                                dst_cell.set_rich_text(rich_text.clone());
                            } 
                            else 
                            {
                                dst_cell.set_value(cell_value);
                            }
                            // println!("dst_cell({}{}).set_value({})", range_ops::index_to_column(col), current_new_row, dst_cell.get_value());
                        }
                        
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
                }

                if added_col
                {
                    current_new_row += 1;
                }
            }

            if added_col
            {
                res = true;
            }

            //apply merged cells formatting to the output sheets. To do: extend if we have formated content of the merger cells!
            if let Some(mrgcells) = merged_cells.iter().find(|range| { range_ops::is_range_in_range(range, &it_range) })
            {
                let mrbeg = mrgcells.get_coordinate_start_row().unwrap().get_num();
                let mrend = mrgcells.get_coordinate_end_row().unwrap().get_num();
                let mcbeg = mrgcells.get_coordinate_start_col().unwrap().get_num();
                let mcend = mrgcells.get_coordinate_end_col().unwrap().get_num(); 
                let mrlen = mrend - mrbeg + 1;

                let mrange = range_ops::make_range_from_indexes(*mcbeg, current_old_row, *mcend, current_old_row+mrlen-1);

                println!("[create_unique_entries_sheet] Range [{}] contains merged cells [{}]", 
                        range_ops::range_to_string(&it_range), mrange.get_range());
                   
                sheet_out.add_merge_cells(mrange.get_range());
            } 
        }
        else 
        {
            res = false;
        }

        println!("[create_unique_entries_sheet]========================================================");
    }
    return res;
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
    let tgt_col = range_ops::column_to_index(col_filter);

    create_unique_entries_sheet(sheet_in, sheet_out, Some(|range_src: &dyn IRange, sheet_dst: &mut Worksheet| 
        {
            let range_in = range_src.get_range();

            println!("[create_unique_entries_sheet] Check if range '{}' is present in the output sheet!", range_ops::range_to_string(range_src.get_range()));

            if !range_ops::is_col_in_range(tgt_col, &range_src.get_range())
            {
                println!("[create_unique_entries_sheet] Input Range [{}] does not contain target column {}!", range_ops::range_to_string(range_src.get_range()), col_filter);
                return false;
            }

            let mut appended = true;
            let max_row = MAX_ROW; //sheet_in.get_highest_row();
            let max_col = MAX_COL; //sheet_in.get_highest_column();

            let mut iter_sheet_dst = range_ops::IterRowMut::new(sheet_dst, max_row, max_col);

            while let Some(mut it) = iter_sheet_dst.next() 
            {
                let it_range_out = it.get_range().clone();
                let it_sheet_out = it.get_sheet_mut();
                
                if range_ops::is_col_in_range(tgt_col, &it_range_out)
                {
                    // range_ops::print_range_cells_1(it.get_sheet(), it.get_range(), Some(12));
                    
                    let allowed_cols: Vec<u32> = col_filter.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();

                    if range_ops::comapre_ranges(sheet_in, range_in, it_sheet_out, &it_range_out, false, None, Some(allowed_cols))
                    {
                        let allowed_cols: Vec<u32> = cols_accum.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();

                        if range_ops::accumulate_ranges(sheet_in, range_in, it_sheet_out, &it_range_out, None, Some(allowed_cols))
                        {
                            appended = false;
                            println!("[create_unique_entries_sheet] Accumulated in-range '{}' to out-range '{}'!", range_ops::range_to_string(range_in), range_ops::range_to_string(&it_range_out));
                        }
                    }
                    else
                    {
                        println!("[create_unique_entries_sheet] in-range [{}] differs from out-range [{}]!", range_ops::range_to_string(range_in), range_ops::range_to_string(&it_range_out));
                    }
                }
                else
                {
                    println!("[create_unique_entries_sheet] out-range [{}] does not contain target column {}!", range_ops::range_to_string(it.get_range()), col_filter);
                }
            }

            if appended
            {
                println!("[create_unique_entries_sheet] Appending in-range '{}' to the output sheet!", range_ops::range_to_string(range_src.get_range()));
            }
            else
            {
                println!("[create_unique_entries_sheet] in-range [{}] is already present in the output sheet!", range_ops::range_to_string(range_src.get_range()));
            }

            appended
        }),
        None::<fn(&dyn IRange, &mut Worksheet) -> bool>,
        None::<fn(&dyn IRange, &mut Worksheet) -> bool>,
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

                // Just create new table with unique values
                // let r = create_unique_entries_sheet(utbl, &mut fotbl, 
                //     None::<fn(&Worksheet, u32, &mut Worksheet) -> bool>,
                //     None::<fn(&Worksheet, u32, &mut Worksheet) -> bool>,
                //     None::<fn(&Worksheet, u32, u32, &mut Worksheet) -> bool>);

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
