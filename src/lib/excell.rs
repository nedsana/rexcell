// use clap::error;
use umya_spreadsheet::*;
use core::range;
use std::collections::HashMap;
use super::common;

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

//compare strings, ignoring white spaces (' ',\t, \n, \r)
fn cmp_strs(s1: &str, s2: &str) -> bool {
    let words1 = s1.split_whitespace();
    let words2 = s2.split_whitespace();
    words1.eq(words2)
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

pub fn apply_formulas(
    rtbl: &Worksheet,
    utbl: &mut Worksheet,
    col_key: u32,
)
{
    let utbl_max_row = utbl.get_highest_row();
    let rtbl_max_row = rtbl.get_highest_row();

    for rtbl_row in 1..=rtbl_max_row //loop over the reference table rows
    {
        let rtbl_key_value = rtbl.get_value((col_key, rtbl_row));
        
        if !rtbl_key_value.is_empty()
        {
            for utbl_row in 1..=utbl_max_row //loop over the update table rows
            {
                let utbl_key_value = utbl.get_value((col_key, utbl_row));

                if !utbl_key_value.is_empty() && cmp_strs(&utbl_key_value, &rtbl_key_value) 
                {
                    // let utbl_name = utbl.get_name().to_string();
                    let utbl_max_col = utbl.get_highest_column();
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
    let utbl_max_row = utbl.get_highest_row();
    let utbl_max_col = utbl.get_highest_column();
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
    
    let utbl_max_row = utbl.get_highest_row();
    let rtbl_max_row = rtbl.get_highest_row();

    for utbl_row in 1..=utbl_max_row //loop over the update table rows
    {
        let utbl_key_value = utbl.get_value((col_key, utbl_row));

        if !utbl_key_value.is_empty() 
        {
            let mut found = false;
            
            for rtbl_row in 1..=rtbl_max_row //loop over the reference table rows
            {
                let rtbl_key_value = rtbl.get_value((col_key, rtbl_row));

                if cmp_strs(&utbl_key_value, &rtbl_key_value) 
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
                                        utbl.get_name(), index_to_column(col_upd), utbl_row, rtbl_upd_value,
                                        rtbl.get_name(), index_to_column(col_upd), rtbl_row));

                    found = true;
                    
                    break;
                }
            }

            if !found
            {
                res.1.push(format!("Can't find '{} {}{}' '{}' in '{}'!", utbl.get_name(), index_to_column(col_upd), utbl_row, utbl_key_value, rtbl.get_name()));
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
        let result = apply_key_value_data_by_indexes(rtbl, utbl, column_to_index(col_key), column_to_index(col_upd));

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

pub fn make_range_from_strings(begin: &str, end: &str) -> Range
{
    let range_str = format!("{}:{}", begin, end);
    let mut range = Range::default();
    range.set_range(range_str);
    range
}

pub fn make_range_from_indexes(bcol: u32, brow: u32, ecol: u32, erow: u32) -> Range
{
    let bcoord = umya_spreadsheet::helper::coordinate::coordinate_from_index(&bcol, &brow); // Returns "A1"
    let ecoord = umya_spreadsheet::helper::coordinate::coordinate_from_index(&ecol, &erow); // Returns "C10"
    make_range_from_strings(&bcoord, &ecoord)
}

fn is_row_in_range(row: u32, range: &Range) -> bool 
{
    if let (Some(start), Some(end)) = (range.get_coordinate_start_row(), range.get_coordinate_end_row()) {
        let start_row = *start.get_num();
        let end_row = *end.get_num();
        return row >= start_row && row <= end_row;
    }
    false
}

fn is_col_in_range(col: u32, range: &Range) -> bool 
{
    if let (Some(start), Some(end)) = (range.get_coordinate_start_col(), range.get_coordinate_end_col()) {
        let start_col = *start.get_num();
        let end_col = *end.get_num();
        return col >= start_col && col <= end_col;
    }
    false
}

fn is_range_in_range(sub_range: &Range, main_range: &Range) -> bool 
{
    let m_start_row = *main_range.get_coordinate_start_row().unwrap().get_num();
    let m_end_row = *main_range.get_coordinate_end_row().unwrap().get_num();
    let m_start_col = *main_range.get_coordinate_start_col().unwrap().get_num();
    let m_end_col = *main_range.get_coordinate_end_col().unwrap().get_num();

    let s_start_row = *sub_range.get_coordinate_start_row().unwrap().get_num();
    let s_end_row = *sub_range.get_coordinate_end_row().unwrap().get_num();
    let s_start_col = *sub_range.get_coordinate_start_col().unwrap().get_num();
    let s_end_col = *sub_range.get_coordinate_end_col().unwrap().get_num();

    s_start_row >= m_start_row && s_end_row <= m_end_row &&
    s_start_col >= m_start_col && s_end_col <= m_end_col
}

fn do_ranges_overlap(range_a: &Range, range_b: &Range) -> bool 
{
    let a_start_row = *range_a.get_coordinate_start_row().unwrap().get_num();
    let a_end_row = *range_a.get_coordinate_end_row().unwrap().get_num();
    let a_start_col = *range_a.get_coordinate_start_col().unwrap().get_num();
    let a_end_col = *range_a.get_coordinate_end_col().unwrap().get_num();

    let b_start_row = *range_b.get_coordinate_start_row().unwrap().get_num();
    let b_end_row = *range_b.get_coordinate_end_row().unwrap().get_num();
    let b_start_col = *range_b.get_coordinate_start_col().unwrap().get_num();
    let b_end_col = *range_b.get_coordinate_end_col().unwrap().get_num();

    let row_overlap = a_start_row <= b_end_row && a_end_row >= b_start_row;
    let col_overlap = a_start_col <= b_end_col && a_end_col >= b_start_col;

    row_overlap && col_overlap
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
    filter_col:  Option<FCol>,
    filter_cell: Option<FCell>,
) -> bool 
where FRow:  Fn(&Worksheet, &Range, &mut Worksheet) -> bool,
      FCol:  Fn(&Worksheet, &Range, &mut Worksheet) -> bool, //is this needed
      FCell: Fn(&Worksheet, &Range, &mut Worksheet) -> bool
{
    let sheet_in_merged_cells = sheet_in.get_merge_cells(); 

    let max_row = sheet_in.get_highest_row();
    let max_col = sheet_in.get_highest_column();
    let mut current_new_row = sheet_out.get_highest_row()+1;

    let mut cells_range = Range::default();

    for row in 1..=max_row 
    {
        if is_row_in_range(row, &cells_range)
        {
            continue; // Skip rows that are part of the provided range
        }

        //make range for the whole row, e.g. A1:Z1
        cells_range = make_range_from_indexes(1, row, 1 + max_col, row);

        // Are there any merged cells that include this row?
        let mut is_merged = false;

        if let Some(merged_cells) = sheet_in_merged_cells.iter().find(|range| { is_row_in_range(row, range) }) 
        {
            let merged_cells_value_inst = sheet_in.get_cell_value((
                merged_cells.get_coordinate_start_col().unwrap().get_num(), 
                merged_cells.get_coordinate_start_row().unwrap().get_num()));

            let _merged_cells_value = merged_cells_value_inst.get_value().clone();
            let merged_cells_value_type = merged_cells_value_inst.get_data_type();

            let is_single_col = *merged_cells.get_coordinate_start_col().unwrap().get_num() == *merged_cells.get_coordinate_end_col().unwrap().get_num();
            let is_two_rows   = *merged_cells.get_coordinate_end_row().unwrap().get_num() - *merged_cells.get_coordinate_start_row().unwrap().get_num() == 1;

            if merged_cells_value_type == "n" && is_single_col && is_two_rows
            {
                is_merged = false;

                cells_range = merged_cells.clone();

                println!("[{}] Merged cell range: {}{}:{}{} found for row {}. Value:{} Type:{}!", 
                    sheet_in.get_name(), 
                    index_to_column(*merged_cells.get_coordinate_start_col().unwrap().get_num()), 
                    merged_cells.get_coordinate_start_row().unwrap().get_num(), 
                    index_to_column(*merged_cells.get_coordinate_end_col().unwrap().get_num()), 
                    merged_cells.get_coordinate_end_row().unwrap().get_num(), row, _merged_cells_value, merged_cells_value_type);
            }
            else
            {
                is_merged = true;
            }
        } 

        if !is_merged 
        {
            let mut added_col = false;
            // Execute per row filter logic, if provided. 
            let mut passes_filter = match &filter_row 
            {
                Some(f) => f(sheet_in, &cells_range, sheet_out),
                None => true,
            };

            if passes_filter
            {
                // Copy the data and formatting cell by cell
                for col in 1..=max_col 
                {
                    //shrink the cell range to this particular column, e.g. A1:A10
                    // cells_range = make_range_from_indexes(col, row, col, row); //is this needed
                    // Execute per col filter logic, if provided.
                    passes_filter = match &filter_col {
                        Some(f) => f(sheet_in, &cells_range, sheet_out),
                        None => true,
                    };

                    if passes_filter 
                    {
                        //shrink the cell range to this particular cell, e.g. A1:A1
                        cells_range = make_range_from_indexes(col, row, col, row);
                        // Execute per row and col filter logic, if provided.
                        passes_filter = match &filter_cell {
                            Some(f) => f(sheet_in, &cells_range, sheet_out),
                            None => true,
                        };

                        if passes_filter
                        {
                            let o_src_cell = sheet_in.get_cell((col, row));
                            if let Some(src_cell) = o_src_cell 
                            {
                                let cell_value = src_cell.get_value().clone();
                                let cell_style = src_cell.get_style().clone();
                                let cell_data_type = src_cell.get_data_type().to_string();

                                let dst_cell = sheet_out.get_cell_mut((col, current_new_row));
                                
                                // Preserve data types when copying cells
                                if cell_data_type == "n" && let Some(num) = src_cell.get_value_number() 
                                {
                                    // println!("dst_cell({}{}).set_value_number({})", index_to_column(col), current_new_row, num);
                                    dst_cell.set_value_number(num);
                                } 
                                else 
                                {
                                    // println!("dst_cell({}{}).set_value({})", index_to_column(col), current_new_row, cell_value.as_str());
                                    // For other data types (text, boolean, date, etc.), use set_value
                                    dst_cell.set_value(cell_value);
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
            // println!("[{}] Row {} is part of a merged cell, skipping!", sheet_in.get_name(), row);
        }
    }
    true
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
    let tgt_col = column_to_index(col_filter);

    create_unique_entries_sheet(sheet_in, sheet_out, Some(|sheet_in: &Worksheet, range: &Range, sheet_out: &mut Worksheet| 
        {
            let rrbeg = *range.get_coordinate_start_row().unwrap().get_num();
            let rrend = *range.get_coordinate_end_row().unwrap().get_num();
            let rcbeg = *range.get_coordinate_start_col().unwrap().get_num();
            let rcend = *range.get_coordinate_end_col().unwrap().get_num();

            let bcoord = umya_spreadsheet::helper::coordinate::coordinate_from_index(&rcbeg, &rrbeg); // Returns "A1"
            let ecoord = umya_spreadsheet::helper::coordinate::coordinate_from_index(&rcend, &rrend); // Returns "C10"

            println!("{}: Range [{}:{}]!", sheet_in.get_name(), bcoord, ecoord);

            for row in rrbeg..=rrend 
            {
                let o_src_cell = sheet_in.get_cell((tgt_col, row));
                if let Some(src_cell) = o_src_cell 
                {
                    let src_cell_value = src_cell.get_value();
                    println!("================== {} ====================", src_cell_value);
                    // Check if the value already exists in the output sheet
                    let max_row_out = sheet_out.get_highest_row();
                    for row_out in 1..=max_row_out 
                    {
                        let o_dst_cell = sheet_out.get_cell((tgt_col, row_out));

                        if let Some(dst_cell) = o_dst_cell 
                        {
                            let dst_cell_value = dst_cell.get_value();

                            if cmp_strs(&dst_cell_value, &src_cell_value)
                            // if dst_cell_value == src_cell_value
                            {
                                println!("  <FOUND> DST({}) [row:{} col:{}] '{}' <-> SRC({}) [row:{} col:{}] '{}'", 
                                    sheet_out.get_name(), row_out, tgt_col, dst_cell_value, sheet_in.get_name(), row, tgt_col, src_cell_value);
                                if cols_accum.len() > 0
                                {
                                    for col_accum in cols_accum.split(',') 
                                    {
                                        let quantity_col = column_to_index(col_accum);
                                        if 0 < quantity_col
                                        {
                                            //the entry is found, but we have to update the cell with quantity
                                            let mut q_cell_value_src = 0.0;
                                            let o_q_cell_src = sheet_in.get_cell((quantity_col, row));
                                            if let Some(q_cell_src) = o_q_cell_src
                                            {
                                                if q_cell_src.get_data_type() == "n"
                                                {
                                                    q_cell_value_src = q_cell_src.get_value().parse::<f32>().unwrap_or(0.0);
                                                }
                                            }

                                            let q_cell_dst = sheet_out.get_cell_mut((quantity_col, row_out));
                                            if q_cell_dst.get_data_type() == "n"
                                            {
                                                let q_cell_value_dst = q_cell_dst.get_value().parse::<f32>().unwrap_or(0.0) + q_cell_value_src;
                                                q_cell_dst.set_value_number(q_cell_value_dst);
                                            }
                                        }
                                    }
                                }

                                return false; // already exists, don't copy
                            }
                            else
                            {
                                println!("<MISSING> DST({}) [row:{} col:{}] '{}' <-> SRC({}) [row:{} col:{}] '{}'. Trying next row ...", 
                                    sheet_out.get_name(), row_out, tgt_col, dst_cell_value, sheet_in.get_name(), row, tgt_col, src_cell_value);
                            }
                        }
                    }
                    println!("< APPEND> DST({}) [row:{} col:{}] '{}' <-> SRC({}) [row:{} col:{}] '{}'", 
                        sheet_out.get_name(), max_row_out, tgt_col, src_cell_value, sheet_in.get_name(), row, tgt_col, src_cell_value);
                    return true;
                }
            }
            false
        }),
        None::<fn(&Worksheet, &Range, &mut Worksheet) -> bool>,
        None::<fn(&Worksheet, &Range, &mut Worksheet) -> bool>,
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
