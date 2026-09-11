// use std::process;
use umya_spreadsheet::*;
use std::collections::HashMap;
use std::iter::Iterator;
use crate::range_types::*;
use log::{debug, info, warn, error};
use super::common;
use super::range_ops;
use evalexpr::*;


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
                            // info!("Found formula({}) in '{} {}{}'", formula, utbl_name, index_to_column(utbl_col), utbl_row);
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
                // info!("Found formula({}) in '{} {}{}'", formula, utbl_name, index_to_column(utbl_col), utbl_row);
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
) -> Result<(), String> 
{
    // info!("rtbl:{} utbl:{} col_key:{} col_upd:{}", rtbl.get_name(), utbl.get_name(), col_key, col_upd);

    let mut found = false;
    
    let utbl_max_row = common::MAX_ROW; //utbl.get_highest_row();
    let rtbl_max_row = common::MAX_ROW; //rtbl.get_highest_row();

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

pub fn apply_key_value_data_by_strings(
    rtbl: &Worksheet,
    utbl: &mut Worksheet,
    col_key: &String,
    cols_upd: &String,
) -> Result<(), String>
{
    if cols_upd.len() == 0 
    {
        error!("{}", common::ERROR_DEST_COL_NOT_DEFINED.to_string());
        return Err(common::ERROR_DEST_COL_NOT_DEFINED.to_string());
    }

    for col_upd in cols_upd.split(',') 
    {
        if let Err(err) = apply_key_value_data_by_indexes(rtbl, utbl, range_ops::column_to_index(col_key), range_ops::column_to_index(col_upd)) 
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
pub fn find_range_in_sheet<'a>(range: &'a dyn IRange, sheet: &'a Worksheet, cmp_cols: &'a Vec<u32>) -> Option<RangeType<'a>>
{
    match range_ops::IterRow::new(sheet, common::MAX_ROW, common::MAX_COL, 1, true, "-", range_ops::Offsets::default())
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
pub fn make_largest_range<'a>(range_in: &'a dyn IRange, sheet_in: &'a Worksheet, cmp_cols: &'a Vec<u32>, acc_cols: &'a Vec<u32>) -> Worksheet
{
    let (_, _, _, _, rows_in, cols_in) = range_ops::range_bounds(range_in.get_range());
   
    let mut tmp_sheet = Worksheet::default();
    tmp_sheet.set_name("TMP_SHEET");

    //Add the input range to the temporary sheet. Any rows, which belong to this group will be appened
    if range_ops::append_range(sheet_in, range_in.get_range(), &mut tmp_sheet, acc_cols) 
    {
        info!("Appended range {}:[{}] to {}", sheet_in.get_name(), range_ops::range_to_string(range_in.get_range()), tmp_sheet.get_name());

        let mut range_tmp = make_range_inst_mut(range_in.get_type(), range_ops::make_range_from_indexes(1, 1, cols_in, rows_in), &mut tmp_sheet);

        match range_ops::IterRow::new(sheet_in, common::MAX_ROW, common::MAX_COL, 1, true, "-", range_ops::Offsets::default())
        {
            Ok(iter_sheet) =>
            {
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

                                        if range_ops::append_range(it.get_sheet(), &missing_entry, &mut range_tmp.get_sheet_mut(), acc_cols)
                                        {
                                            let (br, er, bc, ec, _, _) = range_ops::range_bounds(range_tmp.get_range());

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
                                            error!("Failed to append missing range {}:[{}] to {}!", it.get_sheet().get_name(), 
                                                range_ops::range_to_string(&missing_entry), range_tmp.get_sheet().get_name());
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

    match range_ops::IterRow::new(sheet_in, common::MAX_ROW, common::MAX_COL, 1, true, "-", range_ops::Offsets::default())
    {
        Ok(iter_sheet) =>
        {
            for it in iter_sheet 
            {
                let (brow_it, _, _, _, _, _) = range_ops::range_bounds(it.get_range());
                if "n" != it.get_sheet().get_cell_value((1, brow_it)).get_data_type().to_string()
                {
                    info!("Range {}:[{}] skipping none numeric leading data type!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));
                    continue;
                }
                else
                {
                    info!("Processing range '{}:[{}]'!", it.get_sheet().get_name(), range_ops::range_to_string(it.get_range()));
                }

                loop 
                {
                    if let Some(found_range) = find_range_in_sheet(&it, sheet_out, &cmp_cols)
                    { //accumulating
                        let found_range_clone = found_range.get_range().clone();
                        drop(found_range);

                        info!("Range {} already exists in sheet {}! Accumulating data!", range_ops::range_to_string(it.get_range()), sheet_out.get_name());

                        if range_ops::accumulate_ranges(sheet_in, it.get_range(), sheet_out, &found_range_clone, &cmp_cols, &acc_cols)
                        {
                            info!("Accumulated in-range '{}' to out-range '{}'!", range_ops::range_to_string(it.get_range()), range_ops::range_to_string(&found_range_clone));
                        }

                        break; //exit the internal loop
                    }
                    else
                    { //appending
                        let sheet_largest_range = make_largest_range(&it, sheet_in, &cmp_cols, &acc_cols);

                        match range_ops::IterRow::new(&sheet_largest_range, max_row, max_col, 1, true, "-", range_ops::Offsets::default()) 
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
                                        res = range_ops::append_range(it_slr.get_sheet(), &it_slr.get_range(), sheet_out, &acc_cols);

                                        info!("Appended range {}:[{}] to {}: {}", it_slr.get_sheet().get_name(), range_ops::range_to_string(it_slr.get_range()), sheet_out.get_name(), res);
                                    }
                                    loop_cnt += 1;
                                }
                                if 1 < loop_cnt
                                {
                                    error!("Only one largest range expected! Found {}!", loop_cnt);
                                }
                            },
                            Err(err) => 
                            {
                                error!("Failed to create iterator: {}", err);
                            }
                        }
                        //NOTE: no 'beak' here, because we've appended a range with zeroed numeric cells. The next loop should find this appended range and should update its values properly!
                    } //appending            
                }

                debug!("========================================================");
                // process::exit(1);
                // return res;
            }
            info!("Finished filtering loop, exiting");
        }
        Err(err) =>
        {
            error!("Failed to create iterator: {}", err);
        }
    }
    return res;
}

pub fn apply_calculations(
    sheet:  &mut Worksheet, 
    row:    u32,
    calcs:  &Vec<String>
) -> Result<(), String> 
{
    for scalc in calcs
    {
        match build_operator_tree::<DefaultNumericTypes>(scalc)
        {
            Ok(expr) => 
            {
                let scalc_cols: Vec<&str> = expr.iter_variable_identifiers().collect();

                let sdst_col = scalc_cols[0];
                let dst_col = range_ops::column_to_index(sdst_col);

                let mut context = HashMapContext::<DefaultNumericTypes>::new();

                for scalc_col in scalc_cols
                {
                    let calc_col = range_ops::column_to_index(scalc_col);

                    //get the columns from the formula
                    let mut cell_val = 0.0;
                    match sheet.get_cell((calc_col, row))
                    {
                        Some(calc_col_val) =>
                        {
                            if calc_col_val.get_data_type() == "n" && let Some(num) = calc_col_val.get_value_number()
                            {
                                cell_val = num;
                            } 
                        }
                        None => 
                        {
                            return Err(format!("Failed to read value from {}:{}", sheet.get_name(), range_ops::coords_to_str(calc_col, row)));
                        }
                    }

                    context.set_value(scalc_col.into(), Value::Float(cell_val)).unwrap();
                }

                match expr.eval_empty_with_context_mut(&mut context)
                {
                    Ok(()) =>
                    {
                        let final_f64: f64 = match context.get_value(sdst_col) 
                        {
                            Some(Value::Float(f)) => *f,
                            Some(Value::Int(i)) => *i as f64,
                            _ => {
                                error!("Evaluating expression did not return a numer for '{}'!", sdst_col);
                                0.0
                            }
                        };

                        let dst_cell_total_price = sheet.get_cell_mut((dst_col, row));

                        dst_cell_total_price.set_value_number(final_f64);
                    }
                    Err(err) => 
                    {
                        return Err(format!("Evaluating expression failed '{}'! {}", scalc, err));
                    }
                }
            },
            Err(err) => 
            {
                return Err(format!("Failed to create calculation expression '{}'! {}", scalc, err));
            }
        }
    }
    Ok(())
}

/**
 * Get the necessary data from the analysis sheet and add it to the filtered sheet
 * analysis_sheet         - the sheet with analysis data
 * analysis_col_srch      - the column, from analysis sheet, we search to find the text from 'analysis_srch_pat' (beg range, holding all rows for this section)
 * analysis_col_term      - the column, from analysis sheet, we search to find the text from 'analysis_term_pat' (end range, holding all rows for this section)
 * analysis_cols_copy_src - the columns (comma separated), we want to copy to the filtered sheet
 * analysis_srch_pat      - the text we use to find the start of the section from analysis sheet
 * analysis_term_pat      - the text we use to find the end of the section from analysis sheet
 * filtered_sheet         - the sheet with filtered data
 * filtered_col_srch      - the column, from filtered sheet, we search to find the text from 'analysis_srch_pat'
 * filtered_cols_copy_dst - the columns (comma separated), we want to copy from the analysis sheet
 * filtered_calculations  - calculations (comma separated), we want to apply to the filtered sheet
 */
pub fn get_anaysis_data(
    analysis_sheet:          &Worksheet, 
    analysis_col_srch:       &String,
    analysis_col_term:       &String,
    analysis_cols_copy_src:  &String,
    analysis_srch_pat:       &String,
    analysis_term_pat:       &String,
    filtered_sheet:          &mut Worksheet,
    filtered_col_srch:       &String,
    filtered_cols_copy_dst:  &String,
    filtered_calculations:   &String,
) -> bool 
{
    let max_row = common::MAX_ROW; //sheet_in.get_highest_row();
    let max_col = common::MAX_COL; //sheet_in.get_highest_column();

    let mut res: bool = false;

    let aloop_col = range_ops::column_to_index(analysis_col_srch);
    let asrch_col = range_ops::column_to_index(analysis_col_term);

    let acols_copy_src: Vec<u32> = analysis_cols_copy_src.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();
    let fcols_copy_dst: Vec<u32> = filtered_cols_copy_dst.split(',').map(|s| range_ops::column_to_index(s.trim())).collect();
    let fcalculations: Vec<String> = filtered_calculations.split(',').map(|s| s.trim().to_string()).collect();

    if acols_copy_src.len() != fcols_copy_dst.len()
    {
        error!("Columns to copy from (len:{}) must be the same as the columns to copy to (len:{})! ", acols_copy_src.len(), fcols_copy_dst.len());
        return res;
    }

    let fsrch_col = range_ops::column_to_index(filtered_col_srch);

    match range_ops::IterRowMut::new(filtered_sheet, max_row, max_col, 1, true, "-", range_ops::Offsets::default())
    {
        Ok(mut filtered_sheet_it) => 
        {
            while let Some(mut fit) = range_ops::LendingIterator::next(&mut filtered_sheet_it)
            {
                let (ffbr, _, _, _, _, _) = range_ops::range_bounds(fit.get_range()); //(brow, erow, bcol, ecol, rows, cols)

                if "n" != fit.get_sheet().get_cell_value((1, ffbr)).get_data_type().to_string()
                {
                    info!("Range {}:[{}] skipping none numeric leading data type!", fit.get_sheet().get_name(), range_ops::range_to_string(fit.get_range()));
                    continue;
                }

                while let Some(fitr) = Iterator::next(&mut fit) //loop over the rows of the iterator's range
                {
                    // info!("Tmp Range {}:[{}] sub-range:{}", fit.get_sheet().get_name(), range_ops::range_to_string(fit.get_range()), range_ops::range_to_string(&fitr));

                    let (fbr, _fer, _fbc, _fec, _, _) = range_ops::range_bounds(&fitr); //(brow, erow, bcol, ecol, rows, cols)

                    let filtered_cell_value = fit.get_sheet().get_cell_value((fsrch_col, fbr)).get_value();

                    match range_ops::IterRow::new(analysis_sheet, max_row, max_col, 1, false, analysis_srch_pat, range_ops::Offsets::new(0,0,-1,-1))
                    {
                        Ok(analysis_sheet_it) => 
                        {
                            let mut found_analysis_entry = false;
                            let mut found_analysis_value = false;

                            let mut analysis_data_to_copy: Vec<CellValue> = Vec::new();

                            for ait in analysis_sheet_it
                            {
                                let (abr, aer, _, _, _, _) = range_ops::range_bounds(ait.get_range()); //(brow, erow, bcol, ecol, rows, cols)

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
                                    match apply_calculations(fit.get_sheet_mut(), fbr, &fcalculations)
                                    {
                                        Ok(_) =>
                                        {
                                            res = true;
                                        }
                                        Err(err) => 
                                        {
                                            error!("Failed to apply calculation expressions: {}", err);
                                        }
                                    }
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
                debug!("========================================================");

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

                // Create new table with unique values from cfg.tgt_src_col.When repetition is found, accumulate the values in cfg.tgt_dest_col.
                let r = filter_sheet_by_col_and_accum(utbl, &mut fotbl, &cfg.tgt_src_col, &cfg.tgt_dest_col);
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
                
                if let Err(err) = apply_key_value_data_by_strings(rtbl, utbl, &cfg.tgt_src_col, &cfg.tgt_dest_col)
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

                //Filter (accumulate or append values) the data from initial sheet to filter sheet
                let r = filter_sheet_by_col_and_accum(utbl, &mut fotbl, &cfg.tgt_src_col, &cfg.tgt_dest_col);
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

            //The entries are filtered in the filter sheet. Now get the needed values from analysis table and update the filter sheet.
            if false == get_anaysis_data(atbl, 
                &"A".to_string(), 
                &"B".to_string(), 
                &"A,G".to_string(), 
                &"Позиция: *, *Основание:(.*)".to_string(),
                &"Общо".to_string(), 
                &mut fotbl, 
                &"C".to_string(), 
                &"B,F".to_string(), 
                &"G=E*F".to_string()) //WARNING: hardcoded values!
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
                
                if let Err(err) = apply_key_value_data_by_strings(&fotbl, utbl, &"C".to_string(), &"B,F".to_string())
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
        let mut outfile = target_path.to_str().unwrap().to_string();

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
