use std::process;
use umya_spreadsheet::*;
use std::collections::HashMap;
use crate::range_types::*;

use super::common;
use super::range_ops;

pub fn get_ref_map_by_indexes(sheet: &Worksheet, col_key: u32, col_value: u32) -> HashMap<String, String> {
    let mut ref_map: HashMap<String, String> = HashMap::new();

    for row in 1..=common::MAX_ROW /*sheet.get_highest_row()*/ {
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
    let utbl_max_row = common::MAX_ROW; //utbl.get_highest_row();
    let rtbl_max_row = common::MAX_ROW; //rtbl.get_highest_row();

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
                    let utbl_max_col = common::MAX_COL; //utbl.get_highest_column();
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
    let utbl_max_row = common::MAX_ROW; //utbl.get_highest_row();
    let utbl_max_col = common::MAX_COL; //utbl.get_highest_column();
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
    
    let utbl_max_row = common::MAX_ROW; //utbl.get_highest_row();
    let rtbl_max_row = common::MAX_ROW; //rtbl.get_highest_row();

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
pub fn find_range_in_sheet<'a>(range: &'a dyn IRange, sheet: &'a Worksheet, cmp_cols: &'a Vec<u32>) -> Option<RangeType<'a>>
{
    let iter_sheet = range_ops::IterRow::new(sheet, common::MAX_ROW, common::MAX_COL);
    for it in iter_sheet 
    {
        // if it.compare_range(range, false, None, Some(cmp_cols.clone()))
        if it.contains(range, None, Some(cmp_cols.clone()))
        {
            return Some(it);
        }
    }
    None
}

pub fn clear_worksheet(sheet: &mut Worksheet) 
{
    let (_, max_row) = sheet.get_highest_column_and_row();
    sheet.get_merge_cells_mut().clear();
    if max_row > 0 {
        sheet.remove_row(&1, &max_row);
    }
}

pub fn find_missing_entries(find_where: & dyn IRange, find_what: & dyn IRange, cmp_cols: &Vec<u32>) -> Vec<Range>
{
    let mut res = Vec::new();
    if range_ops::same_types(find_what, find_where)
    {
        let (brow_in, erow_in, _, ecol_in, _, _) = range_ops::range_bounds(find_what.get_range());
        let (brow_it, erow_it, _, _, _, _) = range_ops::range_bounds(find_where.get_range());

        for cmp_col in cmp_cols //to do ... what should happen if we have multiple compare columns
        {
            let bcol_in = *cmp_col;
            let bcol_it = *cmp_col;

            for row_in in (brow_in+1)..=erow_in
            {
                let entry_in = find_what.get_sheet().get_cell_value((bcol_in, row_in)).get_value();

                let mut found_cnt = false;

                for row_it in (brow_it+1)..=erow_it
                {
                    let entry_it = find_where.get_sheet().get_cell_value((bcol_it, row_it)).get_value();

                    // println!("[find_missing_entries] COMPARE {}:[{}:row{} '{}'] to {}:[{}:row{} '{}']!", 
                    //     find_where.get_sheet().get_name(), range_ops::range_to_string(find_where.get_range()), row_it, entry_it,
                    //     find_what.get_sheet().get_name(), range_ops::range_to_string(find_what.get_range()), row_in, entry_in);

                    if range_ops::cmp_strs(&entry_it, &entry_in) 
                    {
                        found_cnt = true;
                        break;
                    }
                }

                if !found_cnt
                {
                    res.push(range_ops::make_range_from_indexes(1, row_in, ecol_in, row_in));
                }
            }
        }
    }
    else
    {
        println!("[find_missing_entries] Types mismatch!"); 
    }
    res
}

/**
 * Scan the workseet to find if there are ranges (Multiline or Merged), with same header, but with more rows than the provided range_out.
 * @return - return temporary Worksheet which contain a single IRange entry (Basic, Merged or Multiline) with all rows which should belong to it.
 */
pub fn make_largest_range<'a>(range_in: &'a dyn IRange, sheet_in: &'a Worksheet, cmp_cols: &'a Vec<u32>) -> Worksheet
{
    let (_, _, _, _, rows_in, cols_in) = range_ops::range_bounds(range_in.get_range());
   
    let mut tmp_sheet = Worksheet::default();
    tmp_sheet.set_name("TMP_SHEET");

    //Add the input range to the temporary sheet. Any rows, which belong to this group will be appened
    if range_ops::append_range(sheet_in, range_in.get_range(), &mut tmp_sheet, false) 
    {
        println!("[make_largest_range] Appended range {}:[{}] to {}", sheet_in.get_name(), range_ops::range_to_string(range_in.get_range()), tmp_sheet.get_name());

        let mut range_tmp = make_range_inst_mut(range_in.get_type(), range_ops::make_range_from_indexes(1, 1, cols_in, rows_in), &mut tmp_sheet);

        let iter_sheet = range_ops::IterRow::new(sheet_in, common::MAX_ROW, common::MAX_COL);
        for it in iter_sheet 
        {
            if range_ops::same_types(&range_tmp, &it)
            {
                let (brow_in, _, _, _, rows_in, _) = range_ops::range_bounds(range_tmp.get_range());
                let (brow_it, _, _, _, rows_it, _) = range_ops::range_bounds(it.get_range());

                for cmp_col in cmp_cols //to do ... what should happen if we have multiple compare columns
                {
                    let bcol_in = *cmp_col;
                    let bcol_it = *cmp_col;

                    //check if the first line of the range_in matches the first line of the current range in the sheets
                    let hdr_in = range_tmp.get_sheet().get_cell_value((bcol_in, brow_in)).get_value();
                    let hdr_it = it.get_sheet().get_cell_value((bcol_it, brow_it)).get_value();

                    if range_ops::cmp_strs(&hdr_in, &hdr_it) 
                    {
                        if rows_it > rows_in 
                        {
                            let missing_entries = find_missing_entries(&range_tmp, &it, cmp_cols);
                            for missing_entry in missing_entries 
                            {
                                let (brow_me, _, _, _, _, _) = range_ops::range_bounds(&missing_entry);

                                if range_ops::append_range(it.get_sheet(), &missing_entry, &mut range_tmp.get_sheet_mut(), true)
                                {
                                    let (br, er, bc, ec, _, _) = range_ops::range_bounds(range_tmp.get_range());

                                    range_tmp = make_range_inst_mut(range_in.get_type(), range_ops::make_range_from_indexes(bc, br, ec, er+1), &mut tmp_sheet);
                                
                                    if range_in.get_type() == IterRowNextKind::Merged
                                    {
                                        let mrange = range_ops::make_range_from_indexes(1, 1, 1, er+1);

                                        println!("[make_largest_range] Appended missing range {}:[{}:'{}'] to {}. Merging {}!", it.get_sheet().get_name(), range_ops::range_to_string(&missing_entry), 
                                            it.get_sheet().get_cell_value((bcol_it, brow_me)).get_value(), range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()));

                                        range_tmp.get_sheet_mut().get_merge_cells_mut().clear();
                                        range_tmp.get_sheet_mut().add_merge_cells(mrange.get_range());
                                    }
                                    else
                                    {
                                        println!("[make_largest_range] Appended missing range {}:[{}:'{}'] to {}", it.get_sheet().get_name(), range_ops::range_to_string(&missing_entry), 
                                            it.get_sheet().get_cell_value((bcol_it, brow_me)).get_value(), range_tmp.get_sheet().get_name());
                                    }
                                }
                                else
                                {
                                    println!("[make_largest_range] Failed to append missing range {}:[{}] to {}!", it.get_sheet().get_name(), 
                                        range_ops::range_to_string(&missing_entry), range_tmp.get_sheet().get_name());
                                }
                            }
                        }
                        else 
                        {
                            // println!("[make_largest_range] {}[{}:'{}'] VS {}[{}:'{}']. KEEPING!", 
                            //         range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()), hdr_in, 
                            //         it.get_sheet().get_name() ,range_ops::range_to_string(it.get_range()), hdr_it);
                        }
                    }
                    else
                    {
                        // println!("[make_largest_range] {}[{}:'{}'] DEFFERENT FROM {}[{}:'{}'].",
                        //         range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()), hdr_in, 
                        //         it.get_sheet().get_name() ,range_ops::range_to_string(it.get_range()), hdr_it);
                    }
                }
            }
            else
            {
            //    println!("[make_largest_range] Types mismatch: {}:{} {}:{}", range_ops::range_to_string(range_tmp.get_range()), range_tmp.get_type_name(), 
            //                                                                 range_ops::range_to_string(it.get_range()), it.get_type_name()); 
            }
        }
    }
    else
    {
        println!("[make_largest_range] Failed to append range {}:[{}] to {}!", sheet_in.get_name(), range_ops::range_to_string(range_in.get_range()), tmp_sheet.get_name());
    }

    tmp_sheet
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

    let max_row = common::MAX_ROW; //sheet_in.get_highest_row();
    let max_col = common::MAX_COL; //sheet_in.get_highest_column();

    let iter_sheet = range_ops::IterRow::new(sheet_in, max_row, max_col);
    for it in iter_sheet 
    {
        let (brow_it, _, _, _, _, _) = range_ops::range_bounds(it.get_range());
        if "n" != it.get_sheet().get_cell_value((1, brow_it)).get_data_type().to_string()
        {
            println!("[filter_sheet_by_col_and_accum] Range {}:[{}] skipping none numeric leading data type!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));
            continue;
        }
        else
        {
            println!("[filter_sheet_by_col_and_accum] Processing range '{}:[{}]'!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));
        }

        loop 
        {
            if let Some(found_range) = find_range_in_sheet(&it, sheet_out, &cmp_cols)
            { //accumulating
                let found_range_clone = found_range.get_range().clone();
                drop(found_range);

                println!("[filter_sheet_by_col_and_accum] Range {} already exists in sheet {}! Accumulating data!", range_ops::range_to_string(it.get_range()), sheet_out.get_name());

                if range_ops::accumulate_ranges(sheet_in, it.get_range(), sheet_out, &found_range_clone, &cmp_cols, &acc_cols)
                {
                    println!("[filter_sheet_by_col_and_accum] Accumulated in-range '{}' to out-range '{}'!", range_ops::range_to_string(it.get_range()), range_ops::range_to_string(&found_range_clone));
                }

                break; //exit the internal loop
            }
            else
            { //appending
                let sheet_largest_range = make_largest_range(&it, sheet_in, &cmp_cols);

                let iter_sheet_largest_range = range_ops::IterRow::new(&sheet_largest_range, max_row, max_col);

                // // temporary file for debugging
                // let mut tmp_ssheet = umya_spreadsheet::new_file(); //DELETE_ME
                // _ = tmp_ssheet.add_sheet(sheet_largest_range.clone()); //DELETE_ME
                // let tmpfname = format!("TMP_SHEET_{}.xlsx", range_ops::range_to_string(it.get_range()));
                // _ = writer::xlsx::write(&tmp_ssheet, std::path::Path::new(&tmpfname)); //DELETE_ME
                // // process::exit(1);

                let mut loop_cnt = 0;
                for it_slr in iter_sheet_largest_range
                {
                    if 0 == loop_cnt
                    {
                        res = range_ops::append_range(it_slr.get_sheet(), &it_slr.get_range(), sheet_out, false);

                        println!("[filter_sheet_by_col_and_accum] Appended range {}:[{}] to {}: {}", it_slr.get_sheet().get_name(), range_ops::range_to_string(it_slr.get_range()), sheet_out.get_name(), res);
                    }
                    loop_cnt += 1;
                }
                if 1 < loop_cnt
                {
                    println!("[filter_sheet_by_col_and_accum] [ERROR] Only one largest range expected! Found {}!", loop_cnt);
                }

                //NOTE: no 'beak' here, because we've appended a range with zeroed numeric cells. The next loop should find this appended range and should update its values properly!
            } //appending            
        }

        println!("[filter_sheet_by_col_and_accum]========================================================");
        // process::exit(1);
        // return res;
    }
    println!("[filter_sheet_by_col_and_accum] Finished loop, exiting");
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
