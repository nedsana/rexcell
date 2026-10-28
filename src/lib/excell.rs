// use std::process;
use umya_spreadsheet::*;
use std::collections::HashMap;
use std::iter::Iterator;
use std::ops::{Deref, DerefMut};
use crate::range_types::*;
use log::{debug, info, warn, error};
use super::common;
use super::range_ops;


pub fn get_ref_map_by_indexes<'a,T>(sheet: &'a WorkSheet<T>, col_key: u32, col_value: u32) -> HashMap<String, String> 
where
    T: Deref<Target = Worksheet>,
{
    let mut ref_map: HashMap<String, String> = HashMap::new();

    for row in 1..=sheet.max_row /*sheet.get_highest_row()*/ {
        let cell_key = sheet.sheet.get_value((col_key, row));
        let cell_value = sheet.sheet.get_value((col_value, row));

        if !cell_value.is_empty() && !cell_key.is_empty() {
            ref_map.insert(cell_value.clone(), cell_key.clone());
        }
    }

    ref_map
}

pub fn get_ref_map_by_strings<'a,T>(sheet: &'a WorkSheet<T>, col_key: &String, col_value: &String) -> HashMap<String, String>
where
    T: Deref<Target = Worksheet>,
{
    get_ref_map_by_indexes(sheet, range_ops::column_to_index(col_key),range_ops::column_to_index(col_value))
}

/**
 * Try to trigger the formula execution
 */
pub fn apply_formulas<'a, R, U>(
    rtbl: &'a WorkSheet<R>,
    utbl: &'a mut WorkSheet<U>,
    col_key: u32,
)
where
    R: Deref<Target = Worksheet>,
    U: DerefMut<Target = Worksheet>,
{
    for rtbl_row in 1..=rtbl.max_row //loop over the reference table rows
    {
        let rtbl_key_value = rtbl.get_value((col_key, rtbl_row));
        
        if !rtbl_key_value.is_empty()
        {
            for utbl_row in 1..=utbl.max_row //loop over the update table rows
            {
                let utbl_key_value = utbl.get_value((col_key, utbl_row));

                if !utbl_key_value.is_empty() && range_ops::cmp_strs(&utbl_key_value, &rtbl_key_value) 
                {
                    for utbl_col in 1..=utbl.max_col
                    {
                        let ucell = utbl.sheet.get_cell_mut((utbl_col, utbl_row));
                        if ucell.is_formula()
                        {
                            let formula: String = ucell.get_formula().to_string();
                            // info!("Found formula({}) in '{} {}{}'", formula, utbl.get_name().to_string(), range_ops::index_to_column(utbl_col), utbl_row);
                            ucell.set_value("");
                            ucell.set_formula(formula);
                        }
                    }
                }
            }
        }
    }
}

pub fn reset_formulas<'a, U>(
    utbl: &'a mut WorkSheet<U>,
)
where
    U: DerefMut<Target = Worksheet>,
{
    for utbl_col in 1..=utbl.max_col //loop over the update table rows
    {
        for utbl_row in 1..=utbl.max_row //loop over the update table rows
        {
            let ucell = utbl.get_cell_mut((utbl_col, utbl_row));
            if ucell.is_formula()
            {
                let formula: String = ucell.get_formula().to_string();
                // info!("Found formula({}) in '{} {}{}'", formula, utbl_name, index_to_column(utbl_col), utbl_row);
                ucell.set_value("");
                ucell.set_formula(formula);
            }
        }
    }
}

pub fn apply_key_value_data_by_indexes<'a, R, U>(
    rtbl: &'a WorkSheet<R>,
    utbl: &'a mut WorkSheet<U>,
    col_key:    u32,
    col_upd:    u32,
    calcs:      &String,
) -> Result<(), String> 
where
    R: Deref<Target = Worksheet>,
    U: DerefMut<Target = Worksheet>,
{
    // info!("rtbl:{} utbl:{} col_key:{} col_upd:{}", rtbl.get_name(), utbl.get_name(), col_key, col_upd);

    let mut found = false;
    
    let utbl_max_row = utbl.max_row;
    let rtbl_max_row = rtbl.max_row;

    for utbl_row in 1..=utbl_max_row //loop over the update table rows
    {
        let utbl_key_value = utbl.get_value((col_key, utbl_row));

        if !utbl_key_value.is_empty() 
        {
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

                        // info!("dst_cell({}{}).get_data_type()={}", index_to_column(col_upd), utbl_row, upd_cell.get_data_type());

                        if upd_cell.get_data_type() == "n" && let Some(num) = upd_cell.get_value_number()
                        {
                            // info!("dst_cell({}{}).set_value_number({})", index_to_column(col_upd), utbl_row, num);
                            dst_cell.set_value_number(num);
                        } 
                        else 
                        {
                            // info!("dst_cell({}{}).set_value({})", index_to_column(col_upd), utbl_row, rtbl_upd_value);
                            dst_cell.set_value(rtbl_upd_value.clone());
                        }
                    } 
                    else 
                    {
                        utbl.get_cell_mut((col_upd, utbl_row)).set_value(rtbl_upd_value.clone());
                    }

                    //apply calculations after the update of the table
                    if let Err(err_msg) = apply_calculations(utbl, utbl_row, &calcs) 
                    {
                        log::error!("Failed to apply formula for row {}: {}", utbl_row, err_msg);
                        return Err(err_msg);
                    }

                    info!("Updated '{} {}{}' with '{}' from '{} {}{}'!", 
                                        utbl.get_name(), range_ops::index_to_column(col_upd), utbl_row, rtbl_upd_value,
                                        rtbl.get_name(), range_ops::index_to_column(col_upd), rtbl_row);

                    found = true;
                    
                    break;
                }
            }

            if !found
            {
                error!("Can't find '{} {}{}' '{}' in '{}'!", utbl.get_name(), range_ops::index_to_column(col_upd), 
                                    utbl_row, utbl_key_value, rtbl.get_name());
            }
        }
    }

    if !found
    {
        error!("{}", common::MESSAGE_NO_KEY_VALUE_MAPPING.to_string());
        return Err(format!("{}", common::MESSAGE_NO_KEY_VALUE_MAPPING.to_string()));
    }
    reset_formulas(utbl);
    Ok(())
}

pub fn apply_key_value_data_by_strings<'a, R, U>(
    rtbl:       &'a WorkSheet<R>,
    utbl:       &'a mut WorkSheet<U>,
    col_key:    &String,
    cols_upd:   &String,
    calcs:      &String,
) -> Result<(), String>
where
    R: Deref<Target = Worksheet>,
    U: DerefMut<Target = Worksheet>,
{
    if cols_upd.len() == 0 
    {
        error!("{}", common::ERROR_DEST_COL_NOT_DEFINED.to_string());
        return Err(common::ERROR_DEST_COL_NOT_DEFINED.to_string());
    }

    for col_upd in cols_upd.split(',') 
    {
        if let Err(err) = apply_key_value_data_by_indexes(rtbl, utbl, range_ops::column_to_index(col_key), range_ops::column_to_index(col_upd), calcs) 
        {
            error!("{}", err);
            return Err(format!("{}", err));
        }
    }
    Ok(())
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
pub fn find_range_in_sheet<'a,T>(range: &'a dyn IRange, wsheet: &'a WorkSheet<T>, cmp_cols: &'a Vec<u32>) -> Option<RangeType<'a>>
where
    T: Deref<Target = Worksheet>,
{
    match range_ops::IterRow::new(wsheet.sheet(), wsheet.max_row(), wsheet.max_col(), 1, true, common::REGEX_MULTILINE, range_ops::Offsets::default())
    {
        Ok(iter_sheet) =>
        {
            for it in iter_sheet 
            {
                // if it.compare_range(range, false, None, Some(cmp_cols.clone()))
                if it.contains(range, None, Some(cmp_cols.clone()))
                {
                    return Some(it);
                }
            }
        },
        Err(err) =>
        {
            error!("Failed to create iterator: {}", err);
        }
    }
    None
}

/** */
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
        if let (Some((brow_in, erow_in, _, ecol_in, _, _)), 
                Some((brow_it, erow_it, _, _, _, _))) = (range_ops::range_bounds(find_what.get_range()), 
                                                          range_ops::range_bounds(find_where.get_range())) 
        {
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

                        // info!("COMPARE {}:[{}:row{} '{}'] to {}:[{}:row{} '{}']!", 
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
    }
    else
    {
        warn!("Types mismatch!"); 
    }
    res
}

/**
 * Scan the workseet to find if there are ranges (Multiline or Merged), with same header, but with more rows than the provided range_out.
 * @return - return temporary Worksheet which contain a single IRange entry (Basic, Merged or Multiline) with all rows which should belong to it.
 */
pub fn make_largest_range<'a,T>(range_in: &'a dyn IRange, sheet_in: &'a WorkSheet<T>, cmp_cols: &'a Vec<u32>, acc_cols: &'a Vec<u32>) -> Worksheet //DELETE_ME: Think of Option/Result return value
where
    T: Deref<Target = Worksheet>,
{
    let mut tmp_sheet = Worksheet::default();
    tmp_sheet.set_name("TMP_SHEET");

    if let Some((_, _, _, _, rows_in, cols_in)) = range_ops::range_bounds(range_in.get_range())
    {
        //Add the input range to the temporary sheet. Any rows, which belong to this group will be appened
        if range_ops::append_range(sheet_in, range_in.get_range(), &mut tmp_sheet, acc_cols) 
        {
            info!("Appended range {}:[{}] to {}", sheet_in.get_name(), range_ops::range_to_string(range_in.get_range()), tmp_sheet.get_name());

            let mut range_tmp = make_range_inst_mut(range_in.get_type(), range_ops::make_range_from_indexes(1, 1, cols_in, rows_in), &mut tmp_sheet);

            match range_ops::IterRow::new(sheet_in, sheet_in.max_row, sheet_in.max_col, 1, true, common::REGEX_MULTILINE, range_ops::Offsets::default())
            {
                Ok(iter_sheet) =>
                {
                    for it in iter_sheet 
                    {
                        if range_ops::same_types(&range_tmp, &it)
                        {
                            if let (Some((brow_in, _, _, _, rows_in, _)), 
                                    Some((brow_it, _, _, _, rows_it, _))) = (range_ops::range_bounds(range_tmp.get_range()), range_ops::range_bounds(it.get_range()))
                            {
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
                                                if let Some((brow_me, _, _, _, _, _)) = range_ops::range_bounds(&missing_entry)
                                                {
                                                    if range_ops::append_range(it.get_sheet(), &missing_entry, &mut range_tmp.get_sheet_mut(), acc_cols)
                                                    {
                                                        if let Some((br, er, bc, ec, _, _)) = range_ops::range_bounds(range_tmp.get_range())
                                                        {
                                                            range_tmp = make_range_inst_mut(range_in.get_type(), range_ops::make_range_from_indexes(bc, br, ec, er+1), &mut tmp_sheet);

                                                            if range_in.get_type() == IterRowNextKind::Merged
                                                            {
                                                                let mrange = range_ops::make_range_from_indexes(1, 1, 1, er+1);

                                                                info!("Appended missing range {}:[{}:'{}'] to {}. Merging {}!", it.get_sheet().get_name(), range_ops::range_to_string(&missing_entry), 
                                                                    it.get_sheet().get_cell_value((bcol_it, brow_me)).get_value(), range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()));

                                                                range_tmp.get_sheet_mut().get_merge_cells_mut().clear();
                                                                range_tmp.get_sheet_mut().add_merge_cells(mrange.get_range());
                                                            }
                                                            else
                                                            {
                                                                info!("Appended missing range {}:[{}:'{}'] to {}", it.get_sheet().get_name(), range_ops::range_to_string(&missing_entry), 
                                                                    it.get_sheet().get_cell_value((bcol_it, brow_me)).get_value(), range_tmp.get_sheet().get_name());
                                                            }
                                                        }
                                                        else
                                                        {
                                                            error!("Can't get append-range dimentions!");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        error!("Failed to append missing range {}:[{}] to {}!", it.get_sheet().get_name(), 
                                                            range_ops::range_to_string(&missing_entry), range_tmp.get_sheet().get_name());
                                                    }
                                                }
                                                else
                                                {
                                                    error!("Can't get missing-range dimentions!");
                                                }
                                            }
                                        }
                                        else 
                                        {
                                            // info!("{}[{}:'{}'] VS {}[{}:'{}']. KEEPING!", 
                                            //         range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()), hdr_in, 
                                            //         it.get_sheet().get_name() ,range_ops::range_to_string(it.get_range()), hdr_it);
                                        }
                                    }
                                    else
                                    {
                                        // info!("{}[{}:'{}'] DEFFERENT FROM {}[{}:'{}'].",
                                        //         range_tmp.get_sheet().get_name(), range_ops::range_to_string(range_tmp.get_range()), hdr_in, 
                                        //         it.get_sheet().get_name() ,range_ops::range_to_string(it.get_range()), hdr_it);
                                    }
                                }
                            }
                            else
                            {
                                error!("Can't get range dimentions!"); 
                            }
                        }
                        else
                        {
                        //    error!("Types mismatch: {}:{} {}:{}", range_ops::range_to_string(range_tmp.get_range()), range_tmp.get_type_name(), 
                        //                                                                 range_ops::range_to_string(it.get_range()), it.get_type_name()); 
                        }
                    }
                },
                Err(err) =>
                {
                    error!("Failed to create iterator: {}", err);
                }
            }
        }
        else
        {
            error!("Failed to append range {}:[{}] to {}!", sheet_in.get_name(), range_ops::range_to_string(range_in.get_range()), tmp_sheet.get_name());
        }
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
    wsheet_in:  &    WorkSheet<&Worksheet>, 
    wsheet_out: &mut WorkSheet<&mut Worksheet>,
    col_filter: &String,
    cols_accum: &String
) -> bool 
{
    let mut res = false;

    let cmp_cols: Vec<u32> = col_filter.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();
    let acc_cols: Vec<u32> = cols_accum.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();

    match range_ops::IterRow::new(wsheet_in.sheet, wsheet_in.max_row, wsheet_in.max_col, 1, true, common::REGEX_MULTILINE, range_ops::Offsets::default())
    {
        Ok(iter_sheet) =>
        {
            for it in iter_sheet 
            {
                if let Some((brow_it, _, _, _, _, _)) = range_ops::range_bounds(it.get_range())
                {
                    if "n" == it.get_sheet().get_cell_value((1, brow_it)).get_data_type().to_string()
                    {
                        info!("Processing range '{}:[{}]'!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));

                        loop 
                        {
                            let wsheet_out = range_ops::effective_bounds_mut(wsheet_out.sheet, 0, 0);

                            if let Some(found_range) = find_range_in_sheet(&it, &wsheet_out, &cmp_cols)
                            { //accumulating
                                let found_range_clone = found_range.get_range().clone();
                                drop(found_range);

                                info!("Range {} already exists in sheet {}! Accumulating data!", range_ops::range_to_string(it.get_range()), wsheet_out.sheet.get_name());

                                if range_ops::accumulate_ranges(wsheet_in.sheet, it.get_range(), wsheet_out.sheet, &found_range_clone, &cmp_cols, &acc_cols)
                                {
                                    info!("Accumulated in-range '{}' to out-range '{}'!", range_ops::range_to_string(it.get_range()), range_ops::range_to_string(&found_range_clone));
                                }

                                break; //exit the internal loop
                            }
                            else
                            { //appending
                                let sheet_largest_range = make_largest_range(&it, wsheet_in, &cmp_cols, &acc_cols);
                                let sheet_largest_range = range_ops::effective_bounds(&sheet_largest_range, 0, 0);

                                match range_ops::IterRow::new(&sheet_largest_range, sheet_largest_range.max_row, sheet_largest_range.max_col, 1, true, common::REGEX_MULTILINE, range_ops::Offsets::default()) 
                                {
                                    Ok(iter_sheet_largest_range) => 
                                    {
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
                                                res = range_ops::append_range(it_slr.get_sheet(), &it_slr.get_range(), wsheet_out.sheet, &acc_cols);

                                                info!("Appended range {}:[{}] to {}: {}", it_slr.get_sheet().get_name(), range_ops::range_to_string(it_slr.get_range()), wsheet_out.sheet.get_name(), res);
                                            }
                                            loop_cnt += 1;
                                        }
                                        if 1 < loop_cnt
                                        {
                                            error!("Only one largest range expected! Found {}!", loop_cnt);
                                            break; //exit the internal loop
                                        }
                                    },
                                    Err(err) => 
                                    {
                                        error!("Failed to create iterator: {}", err);
                                        break; //exit the internal loop
                                    }
                                }
                                //NOTE: no 'beak' here, because we've appended a range with zeroed numeric cells. The next loop should find this appended range and should update its values properly!
                            } //appending
                        }

                    }
                    else
                    {
                        info!("Range {}:[{}] skipping none numeric leading data type!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));
                    }

                    debug!("========================================================");
                    // process::exit(1);
                    // return res;
                }
                else
                {
                    error!("Failed to get range dimenstions!");
                    break;
                }
            }
            info!("Finished filtering loop, exiting");
        }
        Err(err) =>
        {
            error!("Failed to create iterator: {}", err);
        }
    }
    res
}

pub fn convert_to_excel_formula(user_formula: &str, row_num: u32) -> String 
{
    let mut excel_formula = String::with_capacity(user_formula.len() * 2);
    let mut current_var = String::new();

    for c in user_formula.chars() 
    {
        if c.is_ascii_alphabetic() 
        {
            current_var.push(c);
        } 
        else 
        {
            if !current_var.is_empty() 
            {
                excel_formula.push_str(&current_var);
                excel_formula.push_str(&row_num.to_string());
                current_var.clear();
            }
            excel_formula.push(c);
        }
    }

    if !current_var.is_empty() 
    {
        excel_formula.push_str(&current_var);
        excel_formula.push_str(&row_num.to_string());
    }

    excel_formula
}

pub fn apply_calculations(
    sheet:  &mut Worksheet, 
    row:    u32,
    calcs:  &String
) -> Result<(), String> 
{
    let mut res :Result<(), String> = Ok(());
    let fcalcs: Vec<String> = calcs.split(',').map(|s| s.trim().to_string()).collect();
    for scalc in fcalcs
    {
        if scalc.len() > 0
        {
            let eformula = convert_to_excel_formula(&scalc, row);
            if let Some((lhs, rhs)) = eformula.split_once('=') 
            {
                // debug!("LHS:{} = RHS:{}", lhs.trim(), rhs.trim());
                let cell = sheet.get_cell_mut(lhs);
                cell.set_formula(rhs);
            }
            else
            {
                res = Err(format!("Missing sign '='!"));
                break;
            }
        }
    }
    res
}

/**
 * Get the necessary data from the analysis sheet and add it to the filtered sheet
 * analysis_sheet         - the sheet with analysis data
 * analysis_col_srch      - the column, from analysis sheet, we search to find the text from 'analysis_srch_pat' (beg range, holding all rows for this section)
 * analysis_col_term      - the column, from analysis sheet, we search to find the text from 'analysis_term_pat' (end range, holding all rows for this section)
 * analysis_cols_cp_src - the columns (comma separated), we want to copy to the filtered sheet
 * analysis_srch_pat      - the text we use to find the start of the section from analysis sheet
 * analysis_term_pat      - the text we use to find the end of the section from analysis sheet
 * filtered_sheet         - the sheet with filtered data
 * filtered_col_srch      - the column, from filtered sheet, we search to find the text from 'analysis_srch_pat'
 * filtered_cols_copy_dst - the columns (comma separated), we want to copy from the analysis sheet
 * filtered_calculations  - calculations (comma separated), we want to apply to the filtered sheet
 */
pub fn get_anaysis_data<'a, R, U>(
    analysis_sheet:          &'a WorkSheet<R>,
    analysis_col_srch:       &String,
    analysis_col_term:       &String,
    analysis_cols_cp_src:    &String,
    analysis_srch_pat:       &String,
    analysis_term_pat:       &String,
    filtered_sheet:          &'a mut WorkSheet<U>,
    filtered_col_srch:       &String,
    filtered_cols_copy_dst:  &String,
    filtered_calculations:   &String,
) -> bool 
where
    R: Deref<Target = Worksheet>,
    U: DerefMut<Target = Worksheet>,
{
    let mut res: bool = false;

    let aloop_col = range_ops::column_to_index(analysis_col_srch);
    let asrch_col = range_ops::column_to_index(analysis_col_term);

    let acols_copy_src: Vec<u32> = analysis_cols_cp_src.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();
    let fcols_copy_dst: Vec<u32> = filtered_cols_copy_dst.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();

    if acols_copy_src.len() != fcols_copy_dst.len()
    {
        error!("Columns to copy from (len:{}) must be the same as the columns to copy to (len:{})! ", acols_copy_src.len(), fcols_copy_dst.len());
        return res;
    }

    let fsrch_col = range_ops::column_to_index(filtered_col_srch);

    let max_row = filtered_sheet.max_row;
    let max_col = filtered_sheet.max_col;
    match range_ops::IterRowMut::new(filtered_sheet, max_row, max_col, 1, true, common::REGEX_MULTILINE, range_ops::Offsets::default())
    {
        Ok(mut filtered_sheet_it) => 
        {
            while let Some(mut fit) = range_ops::LendingIterator::next(&mut filtered_sheet_it)
            {
                if let Some((ffbr, _, _, _, _, _)) = range_ops::range_bounds(fit.get_range()) //(brow, erow, bcol, ecol, rows, cols)
                {
                    if "n" != fit.get_sheet().get_cell_value((1, ffbr)).get_data_type().to_string()
                    {
                        info!("Range {}:[{}] None numeric leading data type - skipping!", fit.get_sheet().get_name(), range_ops::range_to_string(fit.get_range()));
                        continue;
                    }

                    while let Some(fitr) = Iterator::next(&mut fit) //loop over the rows of the iterator's range
                    {
                        // info!("Tmp Range {}:[{}] sub-range:{}", fit.get_sheet().get_name(), range_ops::range_to_string(fit.get_range()), range_ops::range_to_string(&fitr));

                        if let Some((fbr, _fer, _fbc, _fec, _, _)) = range_ops::range_bounds(&fitr) //(brow, erow, bcol, ecol, rows, cols)
                        {
                            let filtered_cell_value = fit.get_sheet().get_cell_value((fsrch_col, fbr)).get_value();

                            match range_ops::IterRow::new(analysis_sheet, analysis_sheet.max_row, analysis_sheet.max_col, 1, false, analysis_srch_pat, range_ops::Offsets::new(0,0,-1,-1))
                            {
                                Ok(analysis_sheet_it) => 
                                {
                                    let mut found_analysis_entry = false;
                                    let mut found_analysis_value = false;

                                    let mut analysis_data_to_copy: Vec<CellValue> = Vec::new();

                                    for ait in analysis_sheet_it
                                    {
                                        if let Some((abr, aer, _, _, _, _)) = range_ops::range_bounds(ait.get_range()) //(brow, erow, bcol, ecol, rows, cols)
                                        {
                                            let analysis_cell_value = ait.get_sheet().get_cell_value((aloop_col, abr)).get_value();

                                            if range_ops::cmp_strs(&filtered_cell_value, &analysis_cell_value)
                                            {
                                                info!("Found analysis section for '{}:[{}:'{}']'", ait.get_sheet().get_name(), range_ops::coords_to_str(aloop_col, abr), analysis_cell_value);

                                                for ar in (abr..=aer).rev() //loop backwards and get the last analysis_term_pat entry
                                                {
                                                    let cell_value = ait.get_sheet().get_cell_value((asrch_col, ar)).get_value();

                                                    // info!("Scanning '{}:[{}:'{}']'", ait.get_sheet().get_name(), range_ops::coords_to_str(asrch_col, ar), cell_value);

                                                    if range_ops::cmp_strs(analysis_term_pat, &cell_value)
                                                    {
                                                        found_analysis_entry = true;
                                                        found_analysis_value = true;

                                                        for cpcol in &acols_copy_src
                                                        {
                                                            if *cpcol == aloop_col
                                                            {
                                                                analysis_data_to_copy.push(ait.get_sheet().get_cell_value((*cpcol, abr+1)).clone());
                                                            }
                                                            else
                                                            {
                                                                analysis_data_to_copy.push(ait.get_sheet().get_cell_value((*cpcol, ar)).clone());
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            error!("Failed to get analysis range dimensions!");
                                            break;
                                        }
                                    }

                                    if false == found_analysis_entry
                                    {
                                        error!("Failed to find analysis section for '{}:[{}:'{}']'", fit.get_sheet().get_name(), range_ops::coords_to_str(fsrch_col, fbr), filtered_cell_value);
                                    }
                                    else
                                    {
                                        if true == found_analysis_value
                                        {
                                            //copy the values from analysis table to filtered table
                                            for (col_dst, data_to_copy) in fcols_copy_dst.iter().zip(analysis_data_to_copy.iter())
                                            {
                                                let mut s_data_to_copy = data_to_copy.get_value().to_string();
                                                if let Some(last_part) = s_data_to_copy.split(':').last() 
                                                {
                                                    s_data_to_copy = last_part.trim().to_string();
                                                }

                                                info!("Setting value:{} for '{}:[{}:'{}']'", s_data_to_copy, fit.get_sheet().get_name(), range_ops::coords_to_str(*col_dst, fbr), filtered_cell_value);

                                                let dst_cell_desc = fit.get_sheet_mut().get_cell_mut((*col_dst, fbr));

                                                dst_cell_desc.set_value(s_data_to_copy);
                                            }

                                            //apply calculations in the filtered table
                                            if let Err(err_msg) = apply_calculations(fit.get_sheet_mut(), fbr, filtered_calculations) 
                                            {
                                                log::error!("Failed to apply formula for row {}: {}", fbr, err_msg);
                                                return false;
                                            }

                                            res = true;
                                        }
                                        else
                                        {
                                            error!("Failed to find analysis value for '{}:[{}:'{}']'", fit.get_sheet().get_name(), range_ops::coords_to_str(fsrch_col, fbr), filtered_cell_value);
                                        }
                                    }
                                },
                                Err(err) => 
                                {
                                    error!("Failed to create iterator: {}", err);
                                }
                            }
                        }
                        else
                        {
                            error!("Failed to get current filter range dimenstions!");
                            break;
                        }
                    }
                    debug!("========================================================");
                }
                else
                {
                    error!("Failed to get analysis range dimenstions!");
                }
            }
            info!("Finished analysis loop, exiting");
        },
        Err(err) => 
        {
            error!("Failed to create iterator: {}", err);
        }
    }

    return res;
}

pub fn execute(cfg: &common::Config) -> Result<(), String>
{
    let mut res_error: String = String::new();
    let mut count_updated = 0;

    debug!("Cfg:{}", cfg);

    // Load the update Excel file
    let target_path = std::path::Path::new(&cfg.tgt_file);
    let result = reader::xlsx::read(target_path);
    let mut ubook = match result
    {
        Ok(bk) => bk,
        Err(err) => {
            error!("{}:'{}' {}", common::ERROR_CANT_READ_TGT_FILE, target_path.display(), err);
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
                        info!("Found sheets: {}", names);
                        count_updated += 1;
                    }
                    else 
                    {
                        error!("{} {}", common::NO_SHEETS_FOUND, cfg.tgt_file);
                        return Err(format!("{} {}", common::NO_SHEETS_FOUND, cfg.tgt_file));
                    }
                }
                Err(err) => 
                {
                    error!("{}:'{}' {}", common::ERROR_CANT_READ_TGT_FILE, cfg.tgt_file, err);
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
                    None => 
                    {
                        error!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln);
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };

                let     wutbl = range_ops::effective_bounds(utbl, 0, 0);
                let mut wfotbl = range_ops::effective_bounds_mut(&mut fotbl, 0, 0);

                // Create new table with unique values from cfg.tgt_src_col.When repetition is found, accumulate the values in cfg.tgt_acc_col.
                let r = filter_sheet_by_col_and_accum(&wutbl, &mut wfotbl, &cfg.tgt_src_col, &cfg.tgt_acc_col);
                if !r 
                {
                    error!("{}:{}", common::ERROR_FAILED_FILTER_SHEET, utbln);
                    res_error = format!("{}:{}", common::ERROR_FAILED_FILTER_SHEET, utbln);
                    break;
                }
                else
                {
                    info!("{} '{}'", common::FILTERED_SHEET, utbln);
                    count_updated += 1;
                }
            }

            //Add the extra sheet to the book
            let result = ubook.add_sheet(fotbl);
            if let Err(err) = result
            {
                error!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err);
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
                Err(err) => 
                {
                    error!("{}:'{}' {}", common::ERROR_CANT_READ_REF_FILE, ref_path.display(), err);
                    return Err(format!("{}:'{}' {}", common::ERROR_CANT_READ_REF_FILE, ref_path.display(), err));
                }
            };        

            // Get the reference sheet
            let result = rbook.get_sheet_by_name_mut(&cfg.ref_table);
            let rtbl = match result
            {
                Some(tbl) => tbl,
                None => 
                {
                    error!("{}:{}", common::ERROR_REFERENCE_SHEET_NOT_FOUND, cfg.ref_table);
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
                    None => 
                    {
                        error!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln);
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };
                
                let     wrtbl = range_ops::effective_bounds(rtbl, 0, 0);
                let mut wutbl = range_ops::effective_bounds_mut(utbl, 0, 0);

                if let Err(err) = apply_key_value_data_by_strings(&wrtbl, &mut wutbl, &cfg.tgt_src_col, &cfg.tgt_dest_col, &"".to_string())
                {
                    error!("{}:{}", common::MESSAGE_NO_KEY_VALUE_MAPPING, err);
                    return Err(format!("{}:{}", common::MESSAGE_NO_KEY_VALUE_MAPPING, err));
                }
            }
        },

        common::Command::CmdAutocompleteSheets => 
        {
            // Load the analysis Excel file
            let analysis_path = std::path::Path::new(&cfg.analysis_file);
            let result = reader::xlsx::read(analysis_path);
            let abook = match result
            {
                Ok(bk) => bk,
                Err(err) => {
                    error!("{}:'{}' {}", common::ERROR_CANT_READ_ANALYSIS_FILE, analysis_path.display(), err);
                    return Err(format!("{}:'{}' {}", common::ERROR_CANT_READ_ANALYSIS_FILE, analysis_path.display(), err));
                }
            };

            // Get the analysis sheet
            let result = abook.get_sheet_by_name(&cfg.analysis_table);
            let atbl = match result
            {
                Some(tbl) => tbl,
                None => 
                {
                    error!("{}:{}", common::ERROR_ANALYSIS_SHEET_NOT_FOUND, cfg.analysis_table);
                    return Err(format!("{}:{}", common::ERROR_ANALYSIS_SHEET_NOT_FOUND, cfg.analysis_table));
                }
            };

            //create new sheet for the filtered data
            let mut fotbl = Worksheet::default();
            fotbl.set_name(cfg.new_sheet_name.clone());

            //loop over the sheets, where the data is present and adapt the filter sheet
            for utbln in cfg.tgt_upd_table.split(',') 
            {
                // Get the sheet with initial data
                let result = ubook.get_sheet_by_name_mut(&utbln);
                let utbl = match result
                {
                    Some(tbl) => tbl,
                    None => 
                    {
                        error!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln);
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };

                let     wutbl = range_ops::effective_bounds(utbl, 0, 0);
                let mut wfotbl = range_ops::effective_bounds_mut(&mut fotbl, 0, 0);

                //Filter (accumulate or append values) the data from initial sheet to filter sheet
                let r = filter_sheet_by_col_and_accum(&wutbl, &mut wfotbl, &cfg.tgt_src_col, &cfg.tgt_acc_col);
                if !r 
                {
                    error!("{}:{}", common::ERROR_FAILED_FILTER_SHEET, utbln);
                    res_error = format!("{}:{}", common::ERROR_FAILED_FILTER_SHEET, utbln);
                    break;
                }
                else
                {
                    info!("{} '{}'", common::FILTERED_SHEET, utbln);
                    count_updated += 1;
                }
            }

            let     watbl = range_ops::effective_bounds(atbl, 0, 0);
            let mut wfotbl = range_ops::effective_bounds_mut(&mut fotbl, 0, 0);

            //The entries are filtered in the filter sheet. Now get the needed values from analysis table and update the filter sheet.
            if false == get_anaysis_data(
                &watbl, 
                &cfg.analysis_col_srch,                   //&"A".to_string(), 
                &cfg.analysis_col_term,                   //&"B".to_string(), 
                &cfg.analysis_cols_cp_src,                //&"A,G".to_string(), 
                &cfg.analysis_srch_pat,                   //&"Позиция: *, *Основание:(.*)".to_string(),
                &cfg.analysis_term_pat,                   //&"Общо".to_string(), 
                &mut wfotbl, 
                &cfg.tgt_src_col,       //&"C".to_string(), 
                &cfg.tgt_dest_col, //&"B,F".to_string(), 
                &cfg.tgt_calcs)     //&"G=E*F".to_string()
            {
                error!("Failed to process analysis data from {}:{}", cfg.analysis_file, cfg.analysis_table);
                return Err(format!("Failed to process analysis data from {}:{}", cfg.analysis_file, cfg.analysis_table));
            }

            //Now get the data from the filtered sheet and apply it to the initial sheets
            for utbln in cfg.tgt_upd_table.split(',') 
            {
                // Get the update sheet
                let result = ubook.get_sheet_by_name_mut(&utbln);
                let utbl = match result
                {
                    Some(tbl) => tbl,
                    None => 
                    {
                        error!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln);
                        return Err(format!("{}:{}", common::ERROR_UPDATE_SHEET_NOT_FOUND, utbln));
                    }
                };

                let     wfotbl = range_ops::effective_bounds(&fotbl, 0, 0);
                let mut wutbl = range_ops::effective_bounds_mut(utbl, 0, 0);

                if let Err(err) = apply_key_value_data_by_strings(&wfotbl, &mut wutbl, &cfg.tgt_src_col, &cfg.tgt_dest_col, &cfg.tgt_calcs)
                {
                    error!("{}:{}", common::MESSAGE_NO_KEY_VALUE_MAPPING, err);
                    return Err(format!("{}:{}", common::MESSAGE_NO_KEY_VALUE_MAPPING, err));
                }
            }

            //Add the filter sheet to the book
            let result = ubook.add_sheet(fotbl);
            if let Err(err) = result
            {
                error!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err);
                return Err(format!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err));
            };

            //Add the analysis sheet to the book
            let result = ubook.add_sheet(atbl.clone());
            if let Err(err) = result
            {
                error!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err);
                return Err(format!("{}:{}", common::ERROR_FAILED_TO_ADD_SHEET, err));
            };
        },

        _ => 
        {
            error!("{}:{:?}", common::ERROR_INVALID_COMMAND, cfg.command);
            res_error = format!("{}:{:?}", common::ERROR_INVALID_COMMAND, cfg.command);
        },
    }

    // Save the changes if there are any successful updates, otherwise return the error message

    if 0 > count_updated
    {
        error!("{} {}", common::ERROR_NO_ROWS_UPDATED.to_string(), res_error);
        return Err(format!("{} {}", common::ERROR_NO_ROWS_UPDATED.to_string(), res_error))
    }

    if cfg.command == common::Command::CmdFilterSheets || cfg.command == common::Command::CmdUpdateSheets || cfg.command == common::Command::CmdAutocompleteSheets
    {
        let Some(val) = target_path.to_str() else 
        {
            return Err(format!("Target path not found!"))
        };
        let mut outfile = val.to_string();

        // Save changes
        if cfg.inplace 
        {
            let result = writer::xlsx::write(&ubook, target_path);
            if let Err(err) = result 
            {
                error!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, target_path.display(), err);
                return Err(format!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, target_path.display(), err));
            }
        } 
        else 
        {
            let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
            let result = writer::xlsx::write(&ubook, std::path::Path::new(&new_file));
            if let Err(err) = result 
            {
                error!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, new_file, err);
                return Err(format!("{}:{} {}", common::ERROR_UNABLE_TO_WRITE_FILE, new_file, err));
            }
            outfile = new_file;
        }
        info!("Saved file {}!", outfile);
    }
    Ok(())

}

/* ---------------------------------------------------------------------------------------- */

#[cfg(test)]
mod tests {
    use super::*;

    // Sheets are built directly in memory; the functions under test only
    // read/write cells of the given worksheets.
    fn ws(name: &str) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        sheet
    }

    fn put(sheet: &mut Worksheet, col: u32, row: u32, value: &str) {
        sheet.get_cell_mut((col, row)).set_value(value);
    }

    /// A formula cell carrying a stale cached result (like one loaded from
    /// a saved xlsx). set_formula_result_default keeps the formula in place.
    fn put_formula_cached(sheet: &mut Worksheet, col: u32, row: u32, formula: &str, cached: &str) {
        let cell = sheet.get_cell_mut((col, row));
        cell.set_formula(formula);
        cell.set_formula_result_default(cached);
    }

    // ------------------------------------------------------------------
    // get_ref_map_by_indexes / get_ref_map_by_strings
    // ------------------------------------------------------------------

    #[test]
    fn test_0157_get_ref_map_by_indexes_maps_value_to_key() {
        let mut s = ws("RM1");
        put(&mut s, 2, 1, "K1"); put(&mut s, 3, 1, "V1");
        put(&mut s, 2, 2, "K2"); put(&mut s, 3, 2, "V2");
        put(&mut s, 2, 3, "K3"); put(&mut s, 3, 3, "V3");

        let m = get_ref_map_by_indexes(&range_ops::effective_bounds(&s, 0, 0), 2, 3);
        // Documented current behavior: the map is keyed by the VALUE column
        // and holds the KEY column content (reversed vs the parameter names).
        assert_eq!(m.len(), 3);
        assert_eq!(m.get("V1"), Some(&"K1".to_string()));
        assert_eq!(m.get("V2"), Some(&"K2".to_string()));
        assert_eq!(m.get("V3"), Some(&"K3".to_string()));
    }

    #[test]
    fn test_0158_get_ref_map_by_indexes_skips_incomplete_rows() {
        let mut s = ws("RM2");
        put(&mut s, 2, 1, "K1"); // key only
        put(&mut s, 3, 2, "V2"); // value only
        put(&mut s, 2, 3, "K3"); put(&mut s, 3, 3, "V3"); // complete

        let m = get_ref_map_by_indexes(&range_ops::effective_bounds(&s, 0, 0), 2, 3);
        assert_eq!(m.len(), 1);
        assert_eq!(m.get("V3"), Some(&"K3".to_string()));
    }

    #[test]
    fn test_0159_get_ref_map_by_indexes_duplicates() {
        let mut s = ws("RM3");
        // same value, two keys -> last row wins (HashMap insert overwrites)
        put(&mut s, 2, 1, "K1"); put(&mut s, 3, 1, "V");
        put(&mut s, 2, 2, "K2"); put(&mut s, 3, 2, "V");
        // same key, two values -> both entries kept
        put(&mut s, 2, 3, "K3"); put(&mut s, 3, 3, "VA");
        put(&mut s, 2, 4, "K3"); put(&mut s, 3, 4, "VB");

        let m = get_ref_map_by_indexes(&range_ops::effective_bounds(&s, 0, 0), 2, 3);
        assert_eq!(m.len(), 3);
        assert_eq!(m.get("V"), Some(&"K2".to_string()));
        assert_eq!(m.get("VA"), Some(&"K3".to_string()));
        assert_eq!(m.get("VB"), Some(&"K3".to_string()));
    }

    #[test]
    fn test_0160_get_ref_map_by_strings_delegates_to_indexes() {
        let mut s = ws("RM4");
        put(&mut s, 2, 1, "K1"); put(&mut s, 3, 1, "V1");

        let by_idx = get_ref_map_by_indexes(&range_ops::effective_bounds(&s, 0, 0), 2, 3);
        let by_str = get_ref_map_by_strings(&range_ops::effective_bounds(&s, 0, 0), &"B".to_string(), &"C".to_string());
        assert_eq!(by_idx, by_str);
    }

    // ------------------------------------------------------------------
    // apply_formulas / reset_formulas
    //
    // The "reset" of a formula cell is set_value("") + set_formula(f):
    // set_value() wipes the cached raw value (and removes the formula),
    // set_formula() restores the formula but keeps the raw value empty.
    // Observable effect: get_value() becomes empty, formula stays intact.
    // ------------------------------------------------------------------

    #[test]
    fn test_0161_apply_formulas_clears_cache_on_matching_rows_only() {
        let mut rt = ws("AF1R");
        put(&mut rt, 1, 1, "K1");
        put(&mut rt, 1, 2, "K2");

        let mut ut = ws("AF1U");
        put(&mut ut, 1, 1, "K1"); put_formula_cached(&mut ut, 2, 1, "SUM(A1:A2)", "stale1");
        put(&mut ut, 1, 2, "K2"); put_formula_cached(&mut ut, 2, 2, "A2*2", "stale2");
        put(&mut ut, 1, 3, "ZZ"); put_formula_cached(&mut ut, 2, 3, "A3+1", "stale3"); // not in ref
        put_formula_cached(&mut ut, 2, 4, "A4+1", "stale4"); // row without a key
        put_formula_cached(&mut ut, 9, 1, "A9+1", "stale5"); // beyond MAX_COL(8)

        let wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        apply_formulas(&wrt, &mut wut, 1);

        // matching rows: formula kept, cached result cleared
        assert!(ut.get_cell((2, 1)).unwrap().is_formula());
        assert_eq!(ut.get_cell((2, 1)).unwrap().get_formula(), "SUM(A1:A2)");
        assert!(ut.get_value((2, 1)).is_empty());
        assert!(ut.get_cell((2, 2)).unwrap().is_formula());
        assert_eq!(ut.get_cell((2, 2)).unwrap().get_formula(), "A2*2");
        assert!(ut.get_value((2, 2)).is_empty());

        // non-matching rows: cached results intact
        assert_eq!(ut.get_value((2, 3)), "stale3");
        assert_eq!(ut.get_value((2, 4)), "stale4");

        assert_ne!(ut.get_value((9, 1)), "stale5"); 
        assert_eq!(ut.get_cell((9, 1)).unwrap().get_formula(), "A9+1");
    }

    #[test]
    fn test_0162_apply_formulas_matches_keys_case_insensitively() {
        let mut rt = ws("AF2R");
        put(&mut rt, 1, 1, "  Key1 "); // cmp_strs: case- and whitespace-insensitive
        let mut ut = ws("AF2U");
        put(&mut ut, 1, 1, "key1");
        put_formula_cached(&mut ut, 2, 1, "A1+1", "stale");

        let wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        apply_formulas(&wrt, &mut wut, 1);

        assert!(ut.get_cell((2, 1)).unwrap().is_formula());
        assert!(ut.get_value((2, 1)).is_empty()); // matched despite case/spacing
    }

    #[test]
    fn test_0163_apply_formulas_no_reference_keys_touches_nothing() {
        let rt = ws("AF3R"); // empty reference
        let mut ut = ws("AF3U");
        put(&mut ut, 1, 1, "K1");
        put_formula_cached(&mut ut, 2, 1, "A1+1", "stale");

        let wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        apply_formulas(&wrt, &mut wut, 1);

        assert!(ut.get_cell((2, 1)).unwrap().is_formula());
        assert_eq!(ut.get_value((2, 1)), "stale");
    }

    #[test]
    fn test_0164_reset_formulas_clears_all_formula_caches() {
        let mut s = ws("RF1");
        put_formula_cached(&mut s, 1, 1, "A2+1", "stale1");
        put_formula_cached(&mut s, 3, 2, "SUM(A1:B2)", "stale2");
        put(&mut s, 1, 2, "plain");
        put(&mut s, 2, 1, "text");

        let mut ws = range_ops::effective_bounds_mut(&mut s, 0, 0);
        reset_formulas(&mut ws);

        // formulas kept, cached results cleared
        assert!(s.get_cell((1, 1)).unwrap().is_formula());
        assert_eq!(s.get_cell((1, 1)).unwrap().get_formula(), "A2+1");
        assert!(s.get_value((1, 1)).is_empty());
        assert!(s.get_cell((3, 2)).unwrap().is_formula());
        assert_eq!(s.get_cell((3, 2)).unwrap().get_formula(), "SUM(A1:B2)");
        assert!(s.get_value((3, 2)).is_empty());

        // non-formula cells untouched, and not turned into formulas
        assert_eq!(s.get_value((1, 2)), "plain");
        assert_eq!(s.get_value((2, 1)), "text");
        assert!(!s.get_cell((1, 2)).unwrap().is_formula());
    }

    // ------------------------------------------------------------------
    // clear_worksheet
    // ------------------------------------------------------------------

    #[test]
    fn test_0165_clear_worksheet_removes_cells_and_merges() {
        let mut s = ws("CW1");
        put(&mut s, 1, 1, "a");
        put(&mut s, 3, 2, "b");
        put(&mut s, 1, 4, "c");
        s.add_merge_cells("A1:B1");

        clear_worksheet(&mut s);

        assert_eq!(s.get_highest_row(), 0);
        assert_eq!(s.get_highest_column(), 0);
        assert!(s.get_merge_cells().is_empty());
        assert!(s.get_value((1, 1)).is_empty());
        assert!(s.get_value((3, 2)).is_empty());
    }

    #[test]
    fn test_0166_clear_worksheet_empty_sheet_is_noop() {
        let mut s = ws("CW2");
        clear_worksheet(&mut s); // max_row == 0 -> remove_row skipped, no panic
        assert_eq!(s.get_highest_row(), 0);
        assert!(s.get_merge_cells().is_empty());
    }

    // ------------------------------------------------------------------
    // apply_key_value_data_by_indexes / apply_key_value_data_by_strings
    // ------------------------------------------------------------------

    #[test]
    fn test_0167_apply_key_value_data_by_indexes_copies_matched_values() {
        let mut rt = ws("KV1R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "val1");
        put(&mut rt, 1, 2, "K2"); put(&mut rt, 3, 2, "val2");
        let mut ut = ws("KV1U");
        put(&mut ut, 1, 1, "K1");
        put(&mut ut, 1, 2, "K2");
        put(&mut ut, 3, 3, "keep"); // row without a key -> skipped

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).is_ok());

        assert_eq!(ut.get_value((3, 1)), "val1");
        assert_eq!(ut.get_value((3, 2)), "val2");
        assert_eq!(ut.get_value((3, 3)), "keep");
    }

    #[test]
    fn test_0168_apply_key_value_data_by_indexes_numeric_vs_string_copy() {
        let mut rt = ws("KV2R");
        put(&mut rt, 1, 1, "K1");
        rt.get_cell_mut((3, 1)).set_value_number(42.5); // numeric source
        put(&mut rt, 1, 2, "K2");
        put(&mut rt, 3, 2, "text2"); // string source
        let mut ut = ws("KV2U");
        put(&mut ut, 1, 1, "K1");
        put(&mut ut, 1, 2, "K2");

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).is_ok());

        // numeric source cell is copied as a number, not as its text
        assert_eq!(ut.get_cell((3, 1)).unwrap().get_data_type(), "n");
        assert_eq!(ut.get_cell((3, 1)).unwrap().get_value_number(), Some(42.5));
        // string source cell is copied as text
        assert_eq!(ut.get_value((3, 2)), "text2");
    }

    #[test]
    fn test_0169_apply_key_value_data_by_indexes_no_match_returns_error() {
        let mut rt = ws("KV3R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "val1");
        let mut ut = ws("KV3U");
        put(&mut ut, 1, 1, "no-such-key");
        put(&mut ut, 3, 1, "keep");

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        let err = apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).unwrap_err();

        assert_eq!(err, common::MESSAGE_NO_KEY_VALUE_MAPPING);
        assert_eq!(ut.get_value((3, 1)), "keep"); // nothing was written
    }

    #[test]
    fn test_0170_apply_key_value_data_by_indexes_one_match_makes_call_succeed() {
        let mut rt = ws("KV4R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "val1");
        let mut ut = ws("KV4U");
        put(&mut ut, 1, 1, "K1");
        put(&mut ut, 1, 2, "orphan"); // no match in reference

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);
        // Documented current behavior: the `found` flag is never reset per
        // row, so a single matched row makes the whole call succeed and the
        // orphan row is skipped without an Err.
        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).is_ok());
        assert_eq!(ut.get_value((3, 1)), "val1");
        assert!(ut.get_value((3, 2)).is_empty()); // orphan row untouched
    }

    #[test]
    fn test_0171_apply_key_value_data_by_indexes_resets_formula_caches_on_success() {
        let mut rt = ws("KV5R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "val1");

        // success path: reset_formulas runs at the end, wiping cached results
        let mut ut = ws("KV5U");
        put(&mut ut, 1, 1, "K1");
        put_formula_cached(&mut ut, 4, 1, "B1+1", "stale");

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).is_ok());
        assert!(ut.get_cell((4, 1)).unwrap().is_formula());
        assert!(ut.get_value((4, 1)).is_empty());

        // failure path: the early return skips reset_formulas, cache intact
        let mut ut2 = ws("KV5U2");
        put(&mut ut2, 1, 1, "orphan");
        put_formula_cached(&mut ut2, 4, 1, "B1+1", "stale");

        let mut wut2 = range_ops::effective_bounds_mut(&mut ut2, 0, 0);

        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut2, 1, 3, &"".to_string()).is_err());
        assert_eq!(ut2.get_value((4, 1)), "stale");
    }

    #[test]
    fn test_0172_apply_key_value_data_by_indexes_calcs_formula_row_substitution() {
        let mut rt = ws("KV6R");
        put(&mut rt, 1, 2, "K2"); put(&mut rt, 3, 2, "val2"); // reference data in row 2
        let mut ut = ws("KV6U");
        put(&mut ut, 1, 2, "K2");
        let calcs = "D=B*C".to_string(); // user formula, single-letter columns

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &calcs).is_ok());
        // every column letter gets the matched update row appended:
        // "D=B*C" -> "D2=B2*C2" -> formula "B2*C2" written to D2
        let c = ut.get_cell((4, 2)).unwrap();
        assert!(c.is_formula());
        assert_eq!(c.get_formula(), "B2*C2");
    }

    #[test]
    fn test_0173_apply_key_value_data_by_indexes_calcs_missing_equals_errors() {
        let mut rt = ws("KV7R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "val1");
        let mut ut = ws("KV7U");
        put(&mut ut, 1, 1, "K1");
        let calcs = "D B*C".to_string(); // no '=' sign

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        let err = apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &calcs).unwrap_err();
        assert_eq!(err, "Missing sign '='!");
        // Documented current behavior: the value copy happens before the
        // calcs error, so the sheet is left partially updated.
        assert_eq!(ut.get_value((3, 1)), "val1");
    }

    #[test]
    fn test_0174_apply_key_value_data_by_strings_empty_cols_error() {
        let rt = ws("KV8R");
        let mut ut = ws("KV8U");
        
        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        let err = apply_key_value_data_by_strings(&wrt, &mut wut, &"A".to_string(), &"".to_string(), &"".to_string()).unwrap_err();
        assert_eq!(err, common::ERROR_DEST_COL_NOT_DEFINED);
    }

    #[test]
    fn test_0175_apply_key_value_data_by_strings_multiple_columns() {
        let mut rt = ws("KV9R");
        put(&mut rt, 1, 1, "K1");
        put(&mut rt, 2, 1, "vb");
        put(&mut rt, 3, 1, "vc");
        let mut ut = ws("KV9U");
        put(&mut ut, 1, 1, "K1");

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        assert!(apply_key_value_data_by_strings(&wrt, &mut wut, &"A".to_string(), &"B,C".to_string(), &"".to_string()).is_ok());
        assert_eq!(ut.get_value((2, 1)), "vb");
        assert_eq!(ut.get_value((3, 1)), "vc");
    }

    #[test]
    fn test_0176_apply_key_value_data_by_indexes_first_reference_row_wins() {
        let mut rt = ws("KV10R");
        put(&mut rt, 1, 1, "K1"); put(&mut rt, 3, 1, "first");
        put(&mut rt, 1, 2, "K1"); put(&mut rt, 3, 2, "second");
        let mut ut = ws("KV10U");
        put(&mut ut, 1, 1, "K1");

        let     wrt = range_ops::effective_bounds(&rt, 0, 0);
        let mut wut = range_ops::effective_bounds_mut(&mut ut, 0, 0);

        assert!(apply_key_value_data_by_indexes(&wrt, &mut wut, 1, 3, &"".to_string()).is_ok());
        // Documented current behavior: the reference scan breaks on the
        // FIRST matching row, so "first" wins.
        assert_eq!(ut.get_value((3, 1)), "first");
    }

    // ------------------------------------------------------------------
    // get_worksheet_names_list / get_worksheet_names_string /
    // get_worksheet_names
    // ------------------------------------------------------------------

    fn write_xlsx(path: &str, names: &[&str]) {
        let mut book = Spreadsheet::default();
        for &n in names {
            let mut s = Worksheet::default();
            s.set_name(n);
            book.add_sheet(s).unwrap();
        }
        assert!(writer::xlsx::write(&book, std::path::Path::new(path)).is_ok());
    }

    // fn write_sheet_to_xlsx(path: &str, sheet: &Worksheet)  {
    //     let mut book = Spreadsheet::default();
    //     book.add_sheet(sheet.clone()).unwrap();
    //     umya_spreadsheet::writer::xlsx::write(&book, std::path::Path::new(path)).unwrap();
    // }

    #[test]
    fn test_0177_get_worksheet_names_list_and_string() {
        let mut book = Spreadsheet::default();
        for n in ["Alpha", "Beta", "Gamma"] {
            let mut s = Worksheet::default();
            s.set_name(n);
            book.add_sheet(s).unwrap();
        }
        assert_eq!(get_worksheet_names_list(&book), vec!["Alpha".to_string(), "Beta".to_string(), "Gamma".to_string()]);
        assert_eq!(get_worksheet_names_string(&book), "Alpha,Beta,Gamma");
    }

    #[test]
    fn test_0178_get_worksheet_names_reads_file() {
        let path = "/tmp/rexcell_test_names_ok.xlsx";
        write_xlsx(path, &["One", "Two"]);
        assert_eq!(get_worksheet_names(std::path::Path::new(path)), Ok("One,Two".to_string()));
    }

    #[test]
    fn test_0179_get_worksheet_names_missing_file_errors() {
        let path = std::path::Path::new("/tmp/rexcell_test_names_missing.xlsx");
        let err = get_worksheet_names(path).unwrap_err();
        assert!(err.contains("rexcell_test_names_missing.xlsx"));
    }

    // ------------------------------------------------------------------
    // find_range_in_sheet
    //
    // IterRow (pivot col 1, numeric pivot, REGEX_MULTILINE) yields items:
    //   Multiline - a numeric row + following ' - ' string rows
    //   Merged    - rows covered by a merged cell on the pivot column
    //   Basic     - a single row
    // contains() then requires the same range type, an equal header cell
    // in the cmp column, and every data row of the pattern found inside
    // the item (order does not matter).
    // ------------------------------------------------------------------

    fn put_num(sheet: &mut Worksheet, col: u32, row: u32, value: f64) {
        sheet.get_cell_mut((col, row)).set_value_number(value);
    }

    #[test]
    fn test_0180_find_range_in_sheet_multiline_match() {
        let mut out = ws("FR1O");
        put_num(&mut out, 1, 1, 100.0); put(&mut out, 1, 2, "- item1"); put(&mut out, 1, 3, "- item2");
        put_num(&mut out, 1, 4, 200.0); // lone numeric row -> Basic item
        put_num(&mut out, 1, 5, 300.0); put(&mut out, 1, 6, "- x"); put(&mut out, 1, 7, "- y"); put(&mut out, 1, 8, "- z");

        let mut pat = ws("FR1P");
        put_num(&mut pat, 1, 1, 300.0);
        put(&mut pat, 1, 2, "- y"); // data rows in a different order than in 'out'
        put(&mut pat, 1, 3, "- x");
        let pattern = RangeMultiline::new(range_ops::make_range_from_strings("A1", "C3"), &pat);

        let wsheet = range_ops::effective_bounds(&out, 0, 0);
        let cmp = vec![1];
        let found = find_range_in_sheet(&pattern, &wsheet, &cmp).unwrap();
        assert_eq!(found.get_type(), IterRowNextKind::Multiline);
        assert_eq!(found.get_range().get_range(), "A5:B8");
        assert_ne!(found.get_range().get_range(), "A5:I8"); // old behavior: detected items span pivot..pivot+MAX_COL, so the range is A5:I8
    }

    #[test]
    fn test_0181_find_range_in_sheet_basic_single_row_match() {
        let mut out = ws("FR2O");
        put(&mut out, 1, 1, "hdr");
        put(&mut out, 1, 2, "foo");
        put(&mut out, 1, 3, "bar");
        put(&mut out, 1, 4, "baz");

        let mut pat = ws("FR2P");
        put(&mut pat, 1, 1, "bar");
        let pattern = RangeBasic::new(range_ops::make_range_from_strings("A1", "C1"), &pat);

        let wsheet = range_ops::effective_bounds(&out, 0, 0);
        let cmp = vec![1];
        let found = find_range_in_sheet(&pattern, &wsheet, &cmp).unwrap();
        assert_eq!(found.get_type(), IterRowNextKind::Basic);
        assert_eq!(found.get_range().get_range(), "A3:B3");
        assert_ne!(found.get_range().get_range(), "A3:I3"); // old behavior: detected items span pivot..pivot+MAX_COL, so the range is A3:I3
    }

    #[test]
    fn test_0182_find_range_in_sheet_not_found() {
        let mut out = ws("FR3O");
        put(&mut out, 1, 1, "foo");
        put(&mut out, 1, 2, "bar");
        let mut pat = ws("FR3P");
        put(&mut pat, 1, 1, "nope");
        let pattern = RangeBasic::new(range_ops::make_range_from_strings("A1", "C1"), &pat);
        let wsheet = range_ops::effective_bounds(&out, 0, 0);
        assert!(find_range_in_sheet(&pattern, &wsheet, &vec![1]).is_none());
    }

    #[test]
    fn test_0183_find_range_in_sheet_type_mismatch_returns_none() {
        let mut out = ws("FR4O");
        put_num(&mut out, 1, 1, 300.0); // value 300 exists,
        put(&mut out, 1, 2, "- x");     // but only inside a Multiline item
        let mut pat = ws("FR4P");
        put_num(&mut pat, 1, 1, 300.0);
        let pattern = RangeBasic::new(range_ops::make_range_from_strings("A1", "C1"), &pat);
        // Documented current behavior: contains() requires the same range
        // type, so a Basic pattern can never match a Multiline item.
        let wsheet = range_ops::effective_bounds(&out, 0, 0);
        assert!(find_range_in_sheet(&pattern, &wsheet, &vec![1]).is_none());
    }

    // ------------------------------------------------------------------
    // find_missing_entries
    // ------------------------------------------------------------------

    #[test]
    fn test_0184_find_missing_entries_basic() {
        let mut what = ws("FM1W");
        put(&mut what, 1, 1, "H"); // header row is never compared
        put(&mut what, 1, 2, "a");
        put(&mut what, 1, 3, "b");
        put(&mut what, 1, 4, "c");
        let mut where_ = ws("FM1R");
        put(&mut where_, 1, 1, "H");
        put(&mut where_, 1, 2, "B"); // matches 'b' case-insensitively
        put(&mut where_, 1, 3, "x");

        let what_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C4"), &what);
        let where_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C3"), &where_);

        let res: Vec<String> = find_missing_entries(&where_r, &what_r, &vec![1]).iter().map(|r| r.get_range()).collect();
        // missing rows are reported as 1..ecol_in ranges of 'what'
        assert_eq!(res, vec!["A2:C2".to_string(), "A4:C4".to_string()]);
    }

    #[test]
    fn test_0185_find_missing_entries_type_mismatch_empty() {
        let mut what = ws("FM2W");
        put(&mut what, 1, 1, "H"); put(&mut what, 1, 2, "a");
        let mut where_ = ws("FM2R");
        put(&mut where_, 1, 1, "H"); put(&mut where_, 1, 2, "a");
        where_.add_merge_cells("A1:C2");

        let what_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C2"), &what);
        let where_r = RangeMergedCells::new(range_ops::make_range_from_strings("A1", "C2"), &where_);

        assert!(find_missing_entries(&where_r, &what_r, &vec![1]).is_empty());
    }

    #[test]
    fn test_0186_find_missing_entries_multiple_cmp_cols_duplicate_rows() {
        let mut what = ws("FM3W");
        put(&mut what, 1, 1, "H");
        put(&mut what, 1, 2, "a"); put(&mut what, 2, 2, "1");
        put(&mut what, 1, 3, "b"); put(&mut what, 2, 3, "2");
        put(&mut what, 1, 4, "c"); put(&mut what, 2, 4, "3");
        let mut where_ = ws("FM3R");
        put(&mut where_, 1, 1, "H");
        put(&mut where_, 1, 2, "b"); put(&mut where_, 2, 2, "2");
        put(&mut where_, 1, 3, "x"); put(&mut where_, 2, 3, "y");

        let what_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C4"), &what);
        let where_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C3"), &where_);

        // Documented current behavior: each cmp column re-scans all rows
        // independently, so a row missing from several columns is pushed
        // once per column (duplicates in the result).
        let res: Vec<String> = find_missing_entries(&where_r, &what_r, &vec![1, 2]).iter().map(|r| r.get_range()).collect();
        assert_eq!(res, vec!["A2:C2".to_string(), "A4:C4".to_string(), "A2:C2".to_string(), "A4:C4".to_string()]);
    }

    #[test]
    fn test_0187_find_missing_entries_header_only_range_empty() {
        let mut what = ws("FM4W");
        put(&mut what, 1, 1, "H");
        let mut where_ = ws("FM4R");
        put(&mut where_, 1, 1, "H"); put(&mut where_, 1, 2, "a");

        let what_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C1"), &what);
        let where_r = RangeBasic::new(range_ops::make_range_from_strings("A1", "C2"), &where_);

        // 'what' has no data rows -> nothing can be missing
        assert!(find_missing_entries(&where_r, &what_r, &vec![1]).is_empty());
    }

    // ------------------------------------------------------------------
    // make_largest_range: copies range_in into a TMP_SHEET, then scans
    // sheet_in for same-type, same-header ranges with MORE rows and
    // appends their missing data rows. Numeric values in acc_cols are
    // zeroed on copy (so later accumulation can sum into them).
    // ------------------------------------------------------------------

    #[test]
    fn test_0188_make_largest_range_appends_missing_rows_of_larger_same_header_table() {
        let mut sheet = ws("ML1");
        put_num(&mut sheet, 1, 1, 100.0);
        put(&mut sheet, 1, 2, "- a");
        put(&mut sheet, 1, 3, "- b");
        put_num(&mut sheet, 2, 1, 10.0);
        put_num(&mut sheet, 2, 2, 20.0);
        put_num(&mut sheet, 2, 3, 30.0);
        // table 2: same header, one more row
        put_num(&mut sheet, 1, 5, 100.0);
        put(&mut sheet, 1, 6, "- a");
        put(&mut sheet, 1, 7, "- b");
        put(&mut sheet, 1, 8, "- c");
        put_num(&mut sheet, 2, 8, 40.0);

        let range_in = RangeMultiline::new(range_ops::make_range_from_indexes(1, 1, 9, 3), &sheet);
        let cmp = vec![1];
        let acc = vec![2];

        let tmp = make_largest_range(&range_in,  &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        assert_eq!(tmp.get_name(), "TMP_SHEET");
        assert_eq!(tmp.get_value((1, 1)), "100"); // cmp column copied as-is
        assert_eq!(tmp.get_value((1, 2)), "- a");
        assert_eq!(tmp.get_value((1, 3)), "- b");
        assert_eq!(tmp.get_value((1, 4)), "- c"); // appended from table 2
        // numeric values in acc_cols are replaced with 0 on copy
        assert_eq!(tmp.get_value((2, 1)), "0");
        assert_eq!(tmp.get_value((2, 2)), "0");
        assert_eq!(tmp.get_value((2, 3)), "0");
        assert_eq!(tmp.get_value((2, 4)), "0");
        assert_eq!(tmp.get_highest_row(), 4);
    }

    #[test]
    fn test_0189_make_largest_range_ignores_different_header() {
        let mut sheet = ws("ML2");
        put_num(&mut sheet, 1, 1, 100.0);
        put(&mut sheet, 1, 2, "- a");
        put_num(&mut sheet, 1, 4, 200.0); // different header, larger table
        put(&mut sheet, 1, 5, "- a");
        put(&mut sheet, 1, 6, "- b");

        let range_in = RangeMultiline::new(range_ops::make_range_from_indexes(1, 1, 9, 2), &sheet);
        let cmp = vec![1];
        let acc = Vec::new();
        let tmp = make_largest_range(&range_in, &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        // header differs -> nothing appended
        assert_eq!(tmp.get_highest_row(), 2);
        assert!(tmp.get_value((1, 3)).is_empty());
    }

    #[test]
    fn test_0190_make_largest_range_ignores_smaller_same_header_table() {
        let mut sheet = ws("ML3");
        // the input range is the LARGER table
        put_num(&mut sheet, 1, 1, 100.0);
        put(&mut sheet, 1, 2, "- a");
        put(&mut sheet, 1, 3, "- b");
        put(&mut sheet, 1, 4, "- c");
        // same header, but fewer rows
        put_num(&mut sheet, 1, 6, 100.0);
        put(&mut sheet, 1, 7, "- b");

        let range_in = RangeMultiline::new(range_ops::make_range_from_indexes(1, 1, 9, 4), &sheet);
        let cmp = vec![1];
        let acc = Vec::new();
        let tmp = make_largest_range(&range_in, &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        // rows_it > rows_in is false for the smaller table -> untouched
        assert_eq!(tmp.get_highest_row(), 4);
        assert_eq!(tmp.get_value((1, 4)), "- c");
    }

    #[test]
    fn test_0191_make_largest_range_type_mismatch_ignored() {
        let mut sheet = ws("ML4");
        put_num(&mut sheet, 1, 1, 100.0);
        put(&mut sheet, 1, 2, "- a");
        put(&mut sheet, 1, 3, "- b");
        // a LARGER same-header table, but of Merged type
        put_num(&mut sheet, 1, 5, 100.0);
        put(&mut sheet, 1, 6, "x");
        put(&mut sheet, 1, 7, "y");
        put(&mut sheet, 1, 8, "z");
        sheet.add_merge_cells("A5:A8");

        let range_in = RangeMultiline::new(range_ops::make_range_from_indexes(1, 1, 9, 3), &sheet);
        let cmp = vec![1];
        let acc = Vec::new();
        let tmp = make_largest_range(&range_in, &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        // Documented current behavior: only same-type ranges are
        // considered, so the larger Merged table is ignored.
        assert_eq!(tmp.get_highest_row(), 3);
        assert!(tmp.get_value((1, 4)).is_empty());
    }

    #[test]
    fn test_0192_make_largest_range_merged_recreates_column_a_merge() {
        let mut sheet = ws("ML5");
        put_num(&mut sheet, 1, 1, 100.0);
        put(&mut sheet, 1, 2, "x");
        put(&mut sheet, 1, 3, "y");
        sheet.add_merge_cells("A1:A3"); // input range (Merged)
        put_num(&mut sheet, 1, 5, 100.0);
        put(&mut sheet, 1, 6, "x");
        put(&mut sheet, 1, 7, "y");
        put(&mut sheet, 1, 8, "z");
        sheet.add_merge_cells("A5:A8"); // larger Merged, same header

        let range_in = RangeMergedCells::new(range_ops::make_range_from_indexes(1, 1, 9, 3), &sheet);
        let cmp = vec![1];
        let acc = Vec::new();
        let tmp = make_largest_range(&range_in, &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        assert_eq!(tmp.get_value((1, 1)), "100");
        assert_eq!(tmp.get_value((1, 2)), "x");
        assert_eq!(tmp.get_value((1, 3)), "y");
        assert_eq!(tmp.get_value((1, 4)), "z"); // appended missing row
        // the initial append copies the A1:A3 merge; after appending the
        // Merged branch clears it and re-merges column A over all rows
        let merges: Vec<String> = tmp.get_merge_cells().iter().map(|r| r.get_range()).collect();
        assert_eq!(merges, vec!["A1:A4".to_string()]);
    }

    #[test]
    fn test_0193_make_largest_range_basic_is_never_extended() {
        let mut sheet = ws("ML6");
        put(&mut sheet, 1, 1, "hdr");
        put(&mut sheet, 1, 3, "hdr"); // same header in another row

        let range_in = RangeBasic::new(range_ops::make_range_from_indexes(1, 1, 9, 1), &sheet);
        let cmp = vec![1];
        let acc = Vec::new();
        let tmp = make_largest_range(&range_in, &range_ops::effective_bounds(&sheet, 0, 0), &cmp, &acc);

        // Documented current behavior: Basic items are always single-row,
        // so rows_it > rows_in can never hold and nothing is appended.
        assert_eq!(tmp.get_highest_row(), 1);
        assert_eq!(tmp.get_value((1, 1)), "hdr");
    }

    // ------------------------------------------------------------------
    // convert_to_excel_formula / apply_calculations
    // ------------------------------------------------------------------

    #[test]
    fn test_0194_convert_to_excel_formula_appends_row_to_variables() {
        assert_eq!(convert_to_excel_formula("D=B*C", 2), "D2=B2*C2");
        assert_eq!(convert_to_excel_formula("A+B-C", 3), "A3+B3-C3");
        // a run of letters is ONE variable (multi-letter column)
        assert_eq!(convert_to_excel_formula("AB=C", 5), "AB5=C5");
    }

    #[test]
    fn test_0195_convert_to_excel_formula_digit_suffix_quirk() {
        // Documented current behavior: the row number is inserted BEFORE
        // any existing digit suffix, so "A1" becomes "A{row}1".
        assert_eq!(convert_to_excel_formula("A1=B", 2), "A21=B2");
        assert_eq!(convert_to_excel_formula("A1+B2", 3), "A31+B32");
    }

    #[test]
    fn test_0196_convert_to_excel_formula_no_variables_unchanged() {
        assert_eq!(convert_to_excel_formula("", 3), "");
        assert_eq!(convert_to_excel_formula("1+2", 3), "1+2");
        assert_eq!(convert_to_excel_formula("=B", 4), "=B4");
    }

    #[test]
    fn test_0197_apply_calculations_sets_formula_on_lhs_cell() {
        let mut s = ws("CALC1");
        let calcs = "D=B*C".to_string();
        assert!(apply_calculations(&mut s, 2, &calcs).is_ok());
        let c = s.get_cell((4, 2)).unwrap(); // LHS "D2" is created if absent
        assert!(c.is_formula());
        assert_eq!(c.get_formula(), "B2*C2");
    }

    #[test]
    fn test_0198_apply_calculations_multiple_calcs_multiletter_column() {
        let mut s = ws("CALC2");
        let calcs = "AB=B, D=B*C".to_string(); // spaces after ',' are trimmed
        assert!(apply_calculations(&mut s, 3, &calcs).is_ok());
        assert_eq!(s.get_cell((28, 3)).unwrap().get_formula(), "B3"); // col AB
        assert_eq!(s.get_cell((4, 3)).unwrap().get_formula(), "B3*C3");
    }

    #[test]
    fn test_0199_apply_calculations_missing_equals_errors_keeps_prior() {
        let mut s = ws("CALC3");
        let calcs = "D=B*C, X Y".to_string(); // second calc has no '='
        let err = apply_calculations(&mut s, 2, &calcs).unwrap_err();
        assert_eq!(err, "Missing sign '='!");
        // Documented current behavior: calcs apply in order and the error
        // does not roll back the earlier ones.
        assert_eq!(s.get_cell((4, 2)).unwrap().get_formula(), "B2*C2");
    }

    #[test]
    fn test_0200_apply_calculations_empty_segments_are_skipped() {
        let mut s = ws("CALC4");
        assert!(apply_calculations(&mut s, 1, &"".to_string()).is_ok());
        assert!(apply_calculations(&mut s, 1, &",,".to_string()).is_ok());
        assert_eq!(s.get_highest_row(), 0); // nothing was written

        let calcs = "D=B*C,,".to_string(); // trailing empty segments ok
        assert!(apply_calculations(&mut s, 5, &calcs).is_ok());
        assert_eq!(s.get_cell((4, 5)).unwrap().get_formula(), "B5*C5");
    }

    // ------------------------------------------------------------------
    // filter_sheet_by_col_and_accum
    //
    // For each numeric-leading table in sheet_in: if a table with the
    // same header/pivot rows already exists in sheet_out, the values in
    // the accum columns are summed into it; otherwise the (largest
    // merged) table is appended with zeroed accum columns and the values
    // are accumulated into it on the next inner-loop pass.
    // ------------------------------------------------------------------

    #[test]
    fn test_0201_filter_sheet_by_col_and_accum_accumulates_identical_tables() {
        let mut sheet_in = ws("FA1IN");
        put_num(&mut sheet_in, 1, 1, 100.0); put(&mut sheet_in, 1, 2, "- a"); put(&mut sheet_in, 1, 3, "- b");
        put_num(&mut sheet_in, 2, 1, 1.0); put_num(&mut sheet_in, 2, 2, 10.0); put_num(&mut sheet_in, 2, 3, 20.0);
        put_num(&mut sheet_in, 1, 5, 100.0); put(&mut sheet_in, 1, 6, "- a"); put(&mut sheet_in, 1, 7, "- b");
        put_num(&mut sheet_in, 2, 5, 2.0); put_num(&mut sheet_in, 2, 6, 30.0); put_num(&mut sheet_in, 2, 7, 40.0);
        let mut sheet_out = ws("FA1OUT");

        let res = filter_sheet_by_col_and_accum(
            &range_ops::effective_bounds(&sheet_in,0,0), 
            &mut range_ops::effective_bounds_mut(&mut sheet_out,0,0), 
            &"A".to_string(), &"B".to_string());

        assert!(res);
        // one merged table, accum column B summed over both source tables
        assert_eq!(sheet_out.get_value((1, 1)), "100");
        assert_eq!(sheet_out.get_value((1, 2)), "- a");
        assert_eq!(sheet_out.get_value((1, 3)), "- b");
        assert_eq!(sheet_out.get_cell((2, 1)).unwrap().get_value_number(), Some(3.0));
        assert_eq!(sheet_out.get_cell((2, 2)).unwrap().get_value_number(), Some(40.0));
        assert_eq!(sheet_out.get_cell((2, 3)).unwrap().get_value_number(), Some(60.0));
        assert_eq!(sheet_out.get_highest_row(), 3);
    }

    #[test]
    fn test_0202_filter_sheet_by_col_and_accum_appends_different_tables() {
        let mut sheet_in = ws("FA2IN");
        put_num(&mut sheet_in, 1, 1, 100.0); put(&mut sheet_in, 1, 2, "- a"); put(&mut sheet_in, 1, 3, "- b");
        put_num(&mut sheet_in, 2, 1, 1.0); put_num(&mut sheet_in, 2, 2, 10.0); put_num(&mut sheet_in, 2, 3, 20.0);
        put_num(&mut sheet_in, 1, 5, 200.0); put(&mut sheet_in, 1, 6, "- x"); put(&mut sheet_in, 1, 7, "- y");
        put_num(&mut sheet_in, 2, 5, 2.0); put_num(&mut sheet_in, 2, 6, 30.0); put_num(&mut sheet_in, 2, 7, 40.0);
        let mut sheet_out = ws("FA2OUT");

        let res = filter_sheet_by_col_and_accum(
            &range_ops::effective_bounds(&sheet_in,0,0), 
            &mut range_ops::effective_bounds_mut(&mut sheet_out,0,0), 
            &"A".to_string(), &"B".to_string());

        assert!(res);
        // two separate tables in the output, values accumulated into each
        assert_eq!(sheet_out.get_value((1, 1)), "100");
        assert_eq!(sheet_out.get_value((1, 4)), "200");
        assert_eq!(sheet_out.get_value((1, 6)), "- y");
        assert_eq!(sheet_out.get_cell((2, 1)).unwrap().get_value_number(), Some(1.0));
        assert_eq!(sheet_out.get_cell((2, 2)).unwrap().get_value_number(), Some(10.0));
        assert_eq!(sheet_out.get_cell((2, 3)).unwrap().get_value_number(), Some(20.0));
        assert_eq!(sheet_out.get_cell((2, 4)).unwrap().get_value_number(), Some(2.0));
        assert_eq!(sheet_out.get_cell((2, 5)).unwrap().get_value_number(), Some(30.0));
        assert_eq!(sheet_out.get_cell((2, 6)).unwrap().get_value_number(), Some(40.0));
        assert_eq!(sheet_out.get_highest_row(), 6);
    }

    #[test]
    fn test_0203_filter_sheet_by_col_and_accum_skips_non_numeric_leading() {
        let mut sheet_in = ws("FA3IN");
        put(&mut sheet_in, 1, 1, "text");
        put(&mut sheet_in, 1, 2, "more");
        let mut sheet_out = ws("FA3OUT");

        let res = filter_sheet_by_col_and_accum(
            &range_ops::effective_bounds(&sheet_in,0,0), 
            &mut range_ops::effective_bounds_mut(&mut sheet_out,0,0), 
            &"A".to_string(), &"B".to_string());

        // rows not starting with a numeric pivot cell are skipped entirely
        assert!(!res);
        assert_eq!(sheet_out.get_highest_row(), 0);
    }

    #[test]
    fn test_0204_filter_sheet_by_col_and_accum_merges_larger_table_on_append() {
        let mut sheet_in = ws("FA4IN");
        // small table first ...
        put_num(&mut sheet_in, 1, 1, 100.0); put(&mut sheet_in, 1, 2, "- a");
        put_num(&mut sheet_in, 2, 1, 1.0); put_num(&mut sheet_in, 2, 2, 10.0);
        // ... then a larger one with the same header
        put_num(&mut sheet_in, 1, 4, 100.0);
        put(&mut sheet_in, 1, 5, "- a"); put(&mut sheet_in, 1, 6, "- b"); put(&mut sheet_in, 1, 7, "- c");
        put_num(&mut sheet_in, 2, 4, 2.0); put_num(&mut sheet_in, 2, 5, 30.0);
        put_num(&mut sheet_in, 2, 6, 40.0); put_num(&mut sheet_in, 2, 7, 50.0);
        let mut sheet_out = ws("FA4OUT");

        let res = filter_sheet_by_col_and_accum(
            &range_ops::effective_bounds(&sheet_in,0,0), 
            &mut range_ops::effective_bounds_mut(&mut sheet_out,0,0), 
            &"A".to_string(), &"B".to_string());

        assert!(res);
        // the first (small) table is appended via make_largest_range, which
        // already pulls in the missing rows of the larger same-header table
        assert_eq!(sheet_out.get_value((1, 1)), "100");
        assert_eq!(sheet_out.get_value((1, 2)), "- a");
        assert_eq!(sheet_out.get_value((1, 3)), "- b");
        assert_eq!(sheet_out.get_value((1, 4)), "- c");
        // both source tables then accumulate into the merged rows
        assert_eq!(sheet_out.get_cell((2, 1)).unwrap().get_value_number(), Some(3.0));
        assert_eq!(sheet_out.get_cell((2, 2)).unwrap().get_value_number(), Some(40.0));
        assert_eq!(sheet_out.get_cell((2, 3)).unwrap().get_value_number(), Some(40.0));
        assert_eq!(sheet_out.get_cell((2, 4)).unwrap().get_value_number(), Some(50.0));
        assert_eq!(sheet_out.get_highest_row(), 4);
    }

    #[test]
    fn test_0205_filter_sheet_by_col_and_accum_accumulate_only_returns_false() {
        let mut sheet_in = ws("FA5IN");
        put_num(&mut sheet_in, 1, 1, 100.0); put(&mut sheet_in, 1, 2, "- a"); put(&mut sheet_in, 1, 3, "- b");
        put_num(&mut sheet_in, 2, 1, 1.0); put_num(&mut sheet_in, 2, 2, 10.0); put_num(&mut sheet_in, 2, 3, 20.0);
        let mut sheet_out = ws("FA5OUT");
        // sheet_out already holds the same table with its own values
        put_num(&mut sheet_out, 1, 1, 100.0); put(&mut sheet_out, 1, 2, "- a"); put(&mut sheet_out, 1, 3, "- b");
        put_num(&mut sheet_out, 2, 1, 5.0); put_num(&mut sheet_out, 2, 2, 6.0); put_num(&mut sheet_out, 2, 3, 7.0);

        let res = filter_sheet_by_col_and_accum(
            &range_ops::effective_bounds(&sheet_in,0,0), 
            &mut range_ops::effective_bounds_mut(&mut sheet_out,0,0), 
            &"A".to_string(), &"B".to_string());

        // Documented current behavior: res reflects only the append path;
        // a pure accumulation (no append) leaves it false.
        assert!(!res);
        assert_eq!(sheet_out.get_cell((2, 1)).unwrap().get_value_number(), Some(6.0));
        assert_eq!(sheet_out.get_cell((2, 2)).unwrap().get_value_number(), Some(16.0));
        assert_eq!(sheet_out.get_cell((2, 3)).unwrap().get_value_number(), Some(27.0));
    }

    // ------------------------------------------------------------------
    // get_anaysis_data
    //
    // The analysis sheet is iterated with Offsets::new(0,0,-1,-1), which
    // shifts every detected section range UP one row: the section is
    // identified by the column-A cell ABOVE the analysis_srch_pat marker
    // row, the term scan covers [start-1 .. end-1], and the
    // analysis_col_srch copy source is read from the marker row itself.
    // Copied values are cut after their last ':' and trimmed.
    //
    // Non-matching and empty pivot rows each emit a single-row Basic item
    // pointing one row up, so the identification value must sit on a row
    // >= 2 directly above the marker: at row 1 the is_first_line fallback
    // returns that row's item without offset and it preempts the section.
    // ------------------------------------------------------------------

    #[test]
    fn test_0206_get_anaysis_data_copies_section_values() {
        let mut an = ws("GA6AN");
        put(&mut an, 1, 2, "K1");            // identification value (row above marker)
        put(&mut an, 1, 3, "SEC: alpha");    // section marker (matches ^SEC:)
        put(&mut an, 2, 5, "EndMark");       // term row
        put(&mut an, 7, 5, "total: 42");
        let mut f = ws("GA6F");
        put_num(&mut f, 1, 1, 10.0);
        put(&mut f, 2, 1, "K1");             // search value

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C,D".to_string(), &"E=C*D".to_string());

        assert!(res);
        // col A copy = marker row "SEC: alpha" -> after last ':' -> "alpha"
        assert_eq!(f.get_value((3, 1)), "alpha");
        // col G copy from the term row: "total: 42" -> "42"
        assert_eq!(f.get_value((4, 1)), "42");
        assert_eq!(f.get_cell((5, 1)).unwrap().get_formula(), "C1*D1");
    }

    #[test]
    fn test_0207_get_anaysis_data_no_matching_section() {
        let mut an = ws("GA7AN");
        put(&mut an, 1, 2, "K1");
        put(&mut an, 1, 3, "SEC: alpha");
        put(&mut an, 2, 5, "EndMark");
        put(&mut an, 7, 5, "total: 42");
        let mut f = ws("GA7F");
        put_num(&mut f, 1, 1, 10.0);
        put(&mut f, 2, 1, "KX");             // no section has this above its marker

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C,D".to_string(), &"E=C*D".to_string());

        // no section found: nothing is written
        assert!(!res);
        assert_eq!(f.get_value((3, 1)), "");
        assert_eq!(f.get_value((4, 1)), "");
        assert!(f.get_cell((5, 1)).is_none());
    }

    #[test]
    fn test_0208_get_anaysis_data_copy_col_count_mismatch() {
        let an = ws("GA8AN");
        let mut f = ws("GA8F");

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        // 2 source columns vs 1 destination column: rejected up front
        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C".to_string(), &"".to_string());

        assert!(!res);
        assert_eq!(f.get_highest_row(), 0);
    }

    #[test]
    fn test_0209_get_anaysis_data_multiple_sections() {
        let mut an = ws("GA9AN");
        put(&mut an, 1, 2, "K1");
        put(&mut an, 1, 3, "SEC: one");
        put(&mut an, 2, 5, "EndMark");
        put(&mut an, 7, 5, "11: one-val");
        put(&mut an, 1, 6, "K2");
        put(&mut an, 1, 7, "SEC: two");
        put(&mut an, 2, 9, "EndMark");
        put(&mut an, 7, 9, "22: two-val");
        let mut f = ws("GA9F");
        put_num(&mut f, 1, 1, 10.0);
        put(&mut f, 2, 1, "K1");
        put_num(&mut f, 1, 2, 20.0);
        put(&mut f, 2, 2, "K2");

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C,D".to_string(), &"E=C*D".to_string());

        assert!(res);
        assert_eq!(f.get_value((3, 1)), "one");
        assert_eq!(f.get_value((4, 1)), "one-val");
        assert_eq!(f.get_cell((5, 1)).unwrap().get_formula(), "C1*D1");
        assert_eq!(f.get_value((3, 2)), "two");
        assert_eq!(f.get_value((4, 2)), "two-val");
        assert_eq!(f.get_cell((5, 2)).unwrap().get_formula(), "C2*D2");
    }

    #[test]
    fn test_0210_get_anaysis_data_no_term_row() {
        let mut an = ws("GA10AN");
        put(&mut an, 1, 2, "K1");
        put(&mut an, 1, 3, "SEC: alpha");
        // term row removed: column B stays empty
        put(&mut an, 7, 5, "total: 42");
        let mut f = ws("GA10F");
        put_num(&mut f, 1, 1, 10.0);
        put(&mut f, 2, 1, "K1");

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C,D".to_string(), &"E=C*D".to_string());

        // section header matched, but no term row: nothing is copied
        assert!(!res);
        assert_eq!(f.get_value((3, 1)), "");
        assert_eq!(f.get_value((4, 1)), "");
        assert!(f.get_cell((5, 1)).is_none());
    }

    #[test]
    fn test_0211_get_anaysis_data_bad_calcs_leaves_partial_copy() {
        let mut an = ws("GA11AN");
        put(&mut an, 1, 2, "K1");
        put(&mut an, 1, 3, "SEC: alpha");
        put(&mut an, 2, 5, "EndMark");
        put(&mut an, 7, 5, "total: 42");
        let mut f = ws("GA11F");
        put_num(&mut f, 1, 1, 10.0);
        put(&mut f, 2, 1, "K1");

        let     wan = range_ops::effective_bounds(&an, 0, 0);
        let mut wf = range_ops::effective_bounds_mut(&mut f, 0, 0);

        // calcs without '=' -> apply_calculations errors after the copy is done
        let res = get_anaysis_data(&wan, &"A".to_string(), &"B".to_string(), &"A,G".to_string(),
            &"^SEC:".to_string(), &"EndMark".to_string(),
            &mut wf, &"B".to_string(), &"C,D".to_string(), &"ABC".to_string());

        // Documented current behavior: the copied values are already in the
        // sheet when the calculation failure aborts the call.
        assert!(!res);
        assert_eq!(f.get_value((3, 1)), "alpha");
        assert_eq!(f.get_value((4, 1)), "42");
        assert!(f.get_cell((5, 1)).is_none());
    }
}
