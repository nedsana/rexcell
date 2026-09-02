use std::vec;
use umya_spreadsheet::{Worksheet, Range, Cell};
use super::range_types;
// use super::common;
use log::{debug, info, warn, error};

//compare strings, ignoring white spaces (' ',\t, \n, \r)
pub fn cmp_strs(s1: &str, s2: &str) -> bool 
{
    let words1 = s1.split_whitespace();
    let words2 = s2.split_whitespace();
    words1.eq(words2)
}

pub fn column_to_index(col: &str) -> u32 
{
    let mut index = 0;
    for c in col.chars() {
        index = index * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    index
}

pub fn index_to_column(mut index: u32) -> String 
{
    let mut col = String::new();
    while index > 0 {
        index -= 1;
        let remainder = (index % 26) as u8;
        col.push((b'A' + remainder) as char);
        index /= 26;
    }
    col.chars().rev().collect()
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

pub fn coords_to_str(c: u32, r: u32) -> String 
{
    umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r) // Returns "A1"
}

pub fn is_row_in_range(row: u32, range: &Range) -> bool 
{
    if let (Some(start), Some(end)) = (range.get_coordinate_start_row(), range.get_coordinate_end_row()) {
        let start_row = *start.get_num();
        let end_row = *end.get_num();
        return row >= start_row && row <= end_row;
    }
    false
}

pub fn is_col_in_range(col: u32, range: &Range) -> bool 
{
    if let (Some(start), Some(end)) = (range.get_coordinate_start_col(), range.get_coordinate_end_col()) {
        let start_col = *start.get_num();
        let end_col = *end.get_num();
        return col >= start_col && col <= end_col;
    }
    false
}

pub fn range_bounds(range: &Range) -> (u32, u32, u32, u32, u32, u32) {
    let brow = *range.get_coordinate_start_row().unwrap().get_num();
    let erow = *range.get_coordinate_end_row().unwrap().get_num();
    let bcol = *range.get_coordinate_start_col().unwrap().get_num();
    let ecol = *range.get_coordinate_end_col().unwrap().get_num();
    let rows = 1 + erow - brow;
    let cols = 1 + ecol - bcol;
    (brow, erow, bcol, ecol, rows, cols)
}

pub fn is_range_in_range(sub_range: &Range, main_range: &Range) -> bool 
{
    // info!("sub-range:{} range:{}", range_to_string(sub_range), range_to_string(main_range));

    let (m_start_row, m_end_row, m_start_col, m_end_col, _, _) = range_bounds(main_range);
    let (s_start_row, s_end_row, s_start_col, s_end_col, _, _) = range_bounds(sub_range);

    s_start_row >= m_start_row && s_end_row <= m_end_row &&
    s_start_col >= m_start_col && s_end_col <= m_end_col
}

pub fn do_ranges_overlap(range_a: &Range, range_b: &Range) -> bool 
{
    let (a_start_row, a_end_row, a_start_col, a_end_col, _, _) = range_bounds(range_a);
    let (b_start_row, b_end_row, b_start_col, b_end_col, _, _) = range_bounds(range_b);

    let row_overlap = a_start_row <= b_end_row && a_end_row >= b_start_row;
    let col_overlap = a_start_col <= b_end_col && a_end_col >= b_start_col;

    row_overlap && col_overlap
}

pub fn truncate_str(s: &str, max_chars: usize) -> &str 
{
    match s.char_indices().nth(max_chars) 
    {
        Some((idx, _)) => &s[..idx], // Ако низът е по-дълъг, режем до байтовия индекс на N-тия символ
        None => s,                   // Ако е по-кратък, го връщаме цял
    }
}

pub fn truncate_str_with_dots(s: &str, max_chars: usize) -> String 
{
    if s.chars().count() > max_chars 
    {
        let truncated = match s.char_indices().nth(max_chars) 
        {
            Some((idx, _)) => &s[..idx],
            None => s,
        };
        format!("{}...", truncated)
    } 
    else 
    {
        s.to_string()
    }
}

pub fn limit_str(s: &str, max_chars: usize) -> String 
{
    let count = s.chars().count();
    if count == max_chars 
    {
        s.to_string()
    } 
    else if count > max_chars 
    {
        // match s.char_indices().nth(max_chars) 
        // {
        //     Some((idx, _)) => s[..idx].to_string(),
        //     None => s.to_string(),
        // }
        truncate_str_with_dots(s, max_chars)
    } 
    else 
    {
        let mut res = String::with_capacity(max_chars);
        res.push_str(s);
        for _ in 0..(max_chars - count) 
        {
            res.push(' ');
        }
        res
    }
}

pub fn range_to_string(range: &Range) -> String
{
    let (rbeg, rend, cbeg, cend, rows, cols) = range_bounds(range);

    format!("{}:{} (rows:{} cols:{})", umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
                     umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend),  // Returns "C10"
                     rows, cols
    )
}

pub fn print_range_cells_0(sheet: &Worksheet, range: &Range) 
{
    let (rbeg, rend, cbeg, cend, _, _) = range_bounds(range);

    info!("Sheet {} range ({}:{}) ---", sheet.get_name(), 
        umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
        umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend)  // Returns "C10"
    );

    for r in rbeg..=rend 
    {
        for c in cbeg..=cend 
        {
            let coord_str = umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r);
            let cell_value = sheet.get_cell_value((c, r)).get_value();
            info!("Клетка {}: {}", coord_str, cell_value);
        }
    }
}

pub fn print_range_cells_1(sheet: &Worksheet, range: &Range, truncate_len: Option<u32>) 
{
    let (rbeg, rend, cbeg, cend, _, _) = range_bounds(range);

    // choose separator: when truncate_len is provided, use that many spaces,
    // otherwise keep the original tab-based separator
    let sep = match truncate_len {
        Some(tlen) => " ".repeat(tlen as usize),
        None => "\t".to_string(),
    };

    let mut coord_names = Vec::new();
    let mut cell_values = Vec::new();

    for r in rbeg..=rend 
    {
        for c in cbeg..=cend 
        {
            let coord_str = umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r);        
            let _truncated = match truncate_len 
            {
                Some(tlen) => coord_names.push(limit_str(&coord_str, tlen as usize)),
                None => coord_names.push(coord_str),
            };

            let cell_value = sheet.get_cell_value((c, r)).get_value().to_string();
            let _truncated = match truncate_len 
            {
                Some(tlen) => cell_values.push(limit_str(&cell_value, tlen as usize)),
                None => cell_values.push(cell_value) ,
            };
        }

        // info!("{}", coord_names.join(&sep));
        info!("{}", cell_values.join(&sep));
        coord_names.clear();
        cell_values.clear();
    }
}

pub fn comapre_cell(
    sheet_a: &Worksheet, col_a: u32, row_a: u32,
    sheet_b: &Worksheet, col_b: u32, row_b: u32,
    strict: bool
) -> bool 
{
    let mut r = false;
    //Calculate the actual coordinates for sheet A and sheet B and get the text values of the two cells
    let cell_a_coord = (col_a, row_a);
    let cell_b_coord = (col_b, row_b);
    let val_a = sheet_a.get_cell_value(cell_a_coord).get_value();
    let val_b = sheet_b.get_cell_value(cell_b_coord).get_value();


    //check if the cells have rich text and compare them if they do
    let cell_a_obj = sheet_a.get_cell(cell_a_coord);
    let cell_b_obj = sheet_b.get_cell(cell_b_coord);
    let rich_a = cell_a_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());
    let rich_b = cell_b_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());

    if rich_a != rich_b && strict
    {
        info!("[comapre_cell] Rich text mismatch: {}:{} and {}:{}", coords_to_str(col_a, row_a), val_a, coords_to_str(col_b, row_b), val_b);
        return r;
    }

    // If there is any mismatch, immediately stop and return false
    if cmp_strs(&val_a, &val_b) 
    {
        // info!("[comapre_cell] {}:{} equals {}:{}", coords_to_str(col_a, row_a), val_a, coords_to_str(col_b, row_b), val_b);
        r = true;
    }
    else 
    {
        // info!("[comapre_cell] {}:{} differs {}:{}", coords_to_str(col_a, row_a), val_a, coords_to_str(col_b, row_b), val_b); 
        r = false;
    }
    r
}

pub fn comapre_ranges(
    sheet_a: &Worksheet, range_a: &Range,
    sheet_b: &Worksheet, range_b: &Range,
    strict: bool,
    o_allowed_rows: Option<vec::Vec<u32>>, 
    o_allowed_cols: Option<vec::Vec<u32>>,
) -> bool 
{
    //Get the range numeric boundaries for range_a
    let (brow_a, erow_a, bcol_a, _, rows_a, cols_a) = range_bounds(range_a);

    //Get the range numeric boundaries for range_b
    let (brow_b, erow_b, bcol_b, _, rows_b, cols_b) = range_bounds(range_b);

    //If the legths are different, the ranges cannot be the same
    if strict && (rows_a != rows_b || cols_a != cols_b) 
    {
        info!("[comapre_ranges] Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", 
                range_to_string(range_a), rows_a, range_to_string(range_b), rows_b);
        return false;
    }

    let cols_a_offsets: Vec<u32> = (0..=cols_a).collect();
    let cols_b_offsets: Vec<u32> = (0..=cols_b).collect();    
    let allowed_rows: Vec<u32> = o_allowed_rows.unwrap_or_default();
    let allowed_cols: Vec<u32> = o_allowed_cols.unwrap_or_default();

    let _str_allowed_rows = allowed_rows.iter().map(|r| r.to_string()).collect::<Vec<String>>().join(",");
    let _str_allowed_cols = allowed_cols.iter().map(|c| c.to_string()).collect::<Vec<String>>().join(",");

    let mut row_match = 0;
    let mut col_match = 0;

    //Loop using relative offsets to compare the cells in the two ranges
    // for (row_a_offset, row_b_offset) in rows_a_offsets.iter().zip(rows_b_offsets.iter()) 
    for row_num_a in brow_a..=erow_a 
    {
        if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_a)
        {
            // info!("[comapre_ranges], Row A:{} is not in the allowed list {}!", row_num_a, _str_allowed_rows);
            continue; // skip this row if it's not in the allowed_rows list
        }

        for row_num_b in brow_b..=erow_b 
        {
            if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b)
            {
                // info!("[comapre_ranges], Row B:{} is not in the allowed list {}!", row_num_b, _str_allowed_rows);
                continue; // skip this row if it's not in the allowed_rows list
            }

            col_match = 0;
            for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) 
            {
                let col_num_a = bcol_a + col_a_offset;
                let col_num_b = bcol_b + col_b_offset;

                if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) 
                {
                    // info!("[comapre_ranges], Column A:{} or Column B:{} is not in the allowed list {}!", col_num_a, col_num_b, _str_allowed_cols);
                    col_match += 1;
                    continue; // skip this column if it's not in the allowed_cols list
                }

                if comapre_cell(sheet_a, col_num_a, row_num_a, sheet_b, col_num_b, row_num_b, strict)
                {
                    col_match += 1;
                }
            }
        }

        if col_match == cols_a
        {
            row_match += 1;
        }
    }

    if !strict && row_match != rows_a 
    {
        info!("[comapre_ranges] Range {}:[{}, len:{}] DIFFERS FROM Range {}:[{}, len:{}]", 
                sheet_a.get_name(), range_to_string(range_a), rows_a, sheet_b.get_name(), range_to_string(range_b), rows_b);

        return false;
    }

    info!("[comapre_ranges] Range {}:[{}, len:{}] EQUALS TO Range {}:[{}, len:{}]", 
            sheet_a.get_name(), range_to_string(range_a), rows_a, sheet_b.get_name(), range_to_string(range_b), rows_b);

    true
}

pub fn append_range(
    sheet_in:   &Worksheet, 
    range_in:   &Range,
    sheet_out:  & mut Worksheet,
    clear_numeric_fields: bool
) -> bool
{
    let mut res = false;

    let merged_cells = sheet_in.get_merge_cells();

    let mut current_new_row = sheet_out.get_highest_row()+1;

    let current_old_row = current_new_row;

    let (rbeg, rend, cbeg, cend, _, _) = range_bounds(range_in);

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
                    // info!("[append_range] dst_cell({}{}).set_value_number({})", range_ops::index_to_column(col), current_new_row, num);
                    if clear_numeric_fields
                    {
                        dst_cell.set_value_number(0);
                    }
                    else
                    {
                        dst_cell.set_value_number(num);
                    }
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
                    // info!("[append_range] dst_cell({}{}).set_value({})", range_ops::index_to_column(col), current_new_row, dst_cell.get_value());
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
    if let Some(mrgcells) = merged_cells.iter().find(|range| { is_range_in_range(range, range_in) })
    {
        let (_, _, mcbeg, mcend, mrlen, _) = range_bounds(mrgcells);

        let mrange = make_range_from_indexes(mcbeg, current_old_row, mcend, current_old_row + mrlen - 1);

        info!("[append_range] {}:[{}] contains merged cells! Merge cells in {}:[{}]", sheet_in.get_name(), 
            range_to_string(range_in), sheet_out.get_name(), range_to_string(&mrange));
            
        sheet_out.add_merge_cells(mrange.get_range());
    }

    res
}

pub fn accumulate_ranges(
    sheet_a: &     Worksheet, range_a: &Range,
    sheet_b: & mut Worksheet, range_b: &Range,
    pivot_cols: &vec::Vec<u32>, 
    accum_cols: &vec::Vec<u32>,
) -> bool
{
    let mut accumulated: bool = false;

    //Get the range numeric boundaries for range_a
    let (brow_a, _, bcol_a, _, rows_a, cols_a) = range_bounds(range_a);

    //Get the range numeric boundaries for range_b
    let (brow_b, _, bcol_b, _, rows_b, cols_b) = range_bounds(range_b);

    // //If the legths are different, the ranges cannot proceed with accumulation
    // if rows_a != rows_b || cols_a != cols_b 
    // {
    //     info!("[accumulate_ranges] Range size missmatch! {}:[{}] to {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), sheet_b.get_name(), range_to_string(range_b));
    //     return accumulated; 
    // }

    let rows_offsets_a: Vec<u32> = (0..=(rows_a-1)).collect();
    let cols_offsets_a: Vec<u32> = (0..=(cols_a-1)).collect();

    let rows_offsets_b: Vec<u32> = (0..=(rows_b-1)).collect();
    let cols_offsets_b: Vec<u32> = (0..=(cols_b-1)).collect();

    //Loop using relative offsets to compare the cells in the two ranges
    for row_offset_a in &rows_offsets_a 
    {
        let mut cell_a: Option<&Cell> = None;
        let mut cell_b: Option<&Cell> = None;

        let row_num_a = brow_a + row_offset_a;

        for col_offset_a in &cols_offsets_a
        {
            let col_num_a = bcol_a + col_offset_a;
            if pivot_cols.len() > 0 && !pivot_cols.contains(&col_num_a) && !pivot_cols.contains(&col_num_a) //why not just loop over the pivot_cols
            {
                continue; // skip this column if it's not in the pivot_cols list
            }
            cell_a = sheet_a.get_cell((col_num_a, row_num_a));
            break; //for col_offset_a in &cols_offsets_a
        } //for col_offset_a in &cols_offsets_a

        let mut l_found_pivot = false;
        let mut row_num_b = 0;

        for row_offset_b in &rows_offsets_b
        {
            row_num_b = brow_b + row_offset_b;

            for col_offset_b in &cols_offsets_b
            {
                let col_num_b = bcol_b + col_offset_b;
                if pivot_cols.len() > 0 && !pivot_cols.contains(&col_num_b) && !pivot_cols.contains(&col_num_b) //why not just loop over the pivot_cols
                {
                    continue; // skip this column if it's not in the pivot_cols list
                }
                cell_b = sheet_b.get_cell((col_num_b, row_num_b));
                break; //for row_offset_b in &rows_offsets_b
            }

            if cell_b.is_some() && cell_a.is_some()
            {
                // Compare the pivot cells to determine if they are the sames
                let val_a = cell_a.as_ref().unwrap().get_value();
                let val_b = cell_b.as_ref().unwrap().get_value();
                if cmp_strs(&val_a, &val_b) 
                {
                    info!("[accumulate_ranges] Pivot {}:[{}:'{}'] EQUALS TO {}:[{}:'{}']!", sheet_a.get_name(), range_to_string(range_a), val_a, sheet_b.get_name(), range_to_string(range_b), val_b);
                    l_found_pivot = true;
                    break; //for row_offset_b in &rows_offsets_b
                }
            }
        } //for row_offset_b in &rows_offsets_b
        
        if !l_found_pivot
        {
            if cell_a.is_none()
            {
                info!("[accumulate_ranges] Pivot from {}:[{}] is NONE! {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), 
                    sheet_b.get_name(), range_to_string(range_b));
            }
            else 
            {
                info!("[accumulate_ranges] Pivot {}:[{}:'{}'] NOT FOUND IN {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), 
                    cell_a.as_ref().unwrap().get_value(), sheet_b.get_name(), range_to_string(range_b));
            }

            break; //for row_offset_a in &rows_offsets_a 
        }

        for col_offset_a in &cols_offsets_a
        {
            let col_num_a = bcol_a + col_offset_a;

            if accum_cols.len() > 0 && !accum_cols.contains(&col_num_a) 
            {
                continue; // skip this column if it's not in the accum_cols list
            }

            cell_a = sheet_a.get_cell((col_num_a, row_num_a));
            if cell_a.is_none()
            {
                info!("[accumulate_ranges] Cell {}:{} is None!", sheet_a.get_name(), coords_to_str(col_num_a, row_num_a));
            }
            break; //for col_offset_a in &cols_offsets_a
        } //for col_offset_a in &cols_offsets_a

        for col_offset_b in &cols_offsets_b
        {
            let col_num_b = bcol_b + col_offset_b;

            if accum_cols.len() > 0 && !accum_cols.contains(&col_num_b) 
            {
                continue; // skip this column if it's not in the accum_cols list
            }

            cell_b = sheet_b.get_cell((col_num_b, row_num_b));
            if cell_b.is_none()
            {
                info!("[accumulate_ranges] Cell {}:{} is None!", sheet_b.get_name(), coords_to_str(col_num_b, row_num_b));
            }
            break; //for col_offset_b in &cols_offsets_b
        } //for col_offset_b in &cols_offsets_b

        if !cell_a.is_none() && !cell_b.is_none() 
        {
            let cell_a_coord = cell_a.as_ref().unwrap().get_coordinate();
            let coord_a = (cell_a_coord.get_col_num().clone(), cell_a_coord.get_row_num().clone());

            let cell_b_coord = cell_b.as_ref().unwrap().get_coordinate();
            let coord_b = (cell_b_coord.get_col_num().clone(), cell_b_coord.get_row_num().clone());

            if cell_a.as_ref().unwrap().get_data_type() == "n" && cell_b.as_ref().unwrap().get_data_type() == "n" 
            {
                info!("[accumulate_ranges] Accumulating {}:{} to {}:{}", sheet_a.get_name(), coords_to_str(coord_a.0, coord_a.1), 
                        sheet_b.get_name(), coords_to_str(coord_b.0, coord_b.1));

                let val_a = cell_a.as_ref().unwrap().get_value().parse::<f64>().unwrap_or(0.0);
                let val_b = cell_b.as_ref().unwrap().get_value().parse::<f64>().unwrap_or(0.0);
                let sum: f64 = val_a + val_b;

                let q_cell_dst = sheet_b.get_cell_mut(coord_b);
                q_cell_dst.set_value_number(sum);

                accumulated = true;
            }
            else
            {
                info!("[accumulate_ranges] Can't accumulate none-numeric values {}:{} to {}:{}", sheet_a.get_name(), coords_to_str(coord_a.0, coord_a.1), 
                        sheet_b.get_name(), coords_to_str(coord_b.0, coord_b.1));
            }
        }
    } //for row_offset_a in &rows_offsets_a 

    accumulated
}

fn iter_row_next_impl_shared<'a>(
    sheet: &Worksheet,
    current_row: &mut u32,
    max_row: u32,
    max_col: u32,
    _label: &str,
) -> Option<(Range, range_types::IterRowNextKind)> 
{
    let mut ret: Option<(Range, range_types::IterRowNextKind)> = None;

    if max_row > *current_row 
    {
        let sheet_merged_cells = sheet.get_merge_cells();
        let mut cells_range: Range;

        if let Some(merged_cells) = sheet_merged_cells.iter().find(|range| is_row_in_range(*current_row, range)) 
        {
            // info!("[iter_row_next_impl_shared] Found merged cells range '{}'", range_to_string(&merged_cells));

            let (_, merged_end_row, _, _, _, _) = range_bounds(merged_cells);
            cells_range = make_range_from_indexes(1, *current_row, 1 + max_col, merged_end_row);

            let (_, _, _, _, range_rows, _) = range_bounds(&cells_range);

            *current_row += range_rows;

            ret = Some((cells_range, range_types::IterRowNextKind::Merged));
        } 
        else if let Some(src_cell) = sheet.get_cell((1, *current_row)) //will return None if the cell is empty!
        {
            let _first_cell_value = src_cell.get_value().clone();
            let first_cell_data_type = src_cell.get_data_type().to_string();
            
            cells_range = make_range_from_indexes(1, *current_row, 1 + max_col, *current_row);

            if first_cell_data_type == "n" 
            {
                let mut range_rows = {
                    let (_, _, _, _, rows, _) = range_bounds(&cells_range);
                    rows
                };
                let mut multiline = false;

                let next_row = *current_row + 1;
                for nrow in next_row..=max_row 
                {
                    if let Some(next_cell) = sheet.get_cell((1, nrow)) 
                    {
                        let _next_cell_value = next_cell.get_value().clone();
                        let next_cell_data_type = next_cell.get_data_type().to_string();

                        if next_cell_data_type == "n" 
                        {
                            break;
                        } 
                        else if next_cell_data_type == "s" && _next_cell_value == "-" 
                        {
                            cells_range = make_range_from_indexes(1, *current_row, 1 + max_col, nrow);
                            let (_, _, _, _, new_range_rows, _) = range_bounds(&cells_range);
                            range_rows = new_range_rows;
                            multiline = true;
                        }
                    }
                }

                *current_row += range_rows;

                if multiline 
                {
                    // info!("[iter_row_next_impl_shared] Range [{}]: from multiline cells! current_row={}", range_to_string(&cells_range), current_row);
                    ret = Some((cells_range, range_types::IterRowNextKind::Multiline));
                } 
                else 
                {
                    // info!("[iter_row_next_impl_shared] Range [{}]: from regular cells! current_row={}", range_to_string(&cells_range), current_row);
                    ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                }
            } 
            else 
            {
                // info!("[iter_row_next_impl_shared] Current row {} starts with unexpected type:'{}'!", *current_row, first_cell_data_type);
                ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                *current_row += 1;
            }
        } 
        else 
        {
            //either row with empty cells was found or we've reached the end of the document. To destingwish:
            let mut is_last_row = true;
            let bridx = *current_row;
            let eridx = bridx+20;
            for ridx in bridx..eridx
            {
                if let Some(_cell) = sheet.get_cell((1, ridx))
                {
                    is_last_row = false;
                    break;
                }
            }

            if is_last_row
            {
                info!("[iter_row_next_impl_shared] Processing unexpected row:{}!", *current_row);
            }
            else
            {
                cells_range = make_range_from_indexes(1, *current_row, 1 + max_col, *current_row);
                // info!("[iter_row_next_impl_shared] Range [{}]: from regular empty cells! current_row={}", range_to_string(&cells_range), current_row);
                ret = Some((cells_range, range_types::IterRowNextKind::Basic));
            }
            *current_row += 1;
        }
    }
    else
    {
        info!("[iter_row_next_impl_shared] Reached maximum lines to process:{}!", max_row);
    }
    ret
}

// Helper function that contains the common logic for IterRow and IterRowMut next() method
fn iter_row_next_impl<'a>(
    sheet: &'a Worksheet,
    current_row: &mut u32,
    max_row: u32,
    max_col: u32,
) -> Option<range_types::RangeType<'a>> 
{
    let mut ret: Option<range_types::RangeType<'a>> = None;
    match iter_row_next_impl_shared(sheet, current_row, max_row, max_col, "iter_row_next_impl") 
    {
        Some((range, range_types::IterRowNextKind::Basic)) => ret = Some(range_types::RangeType::Basic(range_types::RangeBasic { range, sheet })),
        Some((range, range_types::IterRowNextKind::Merged)) => ret = Some(range_types::RangeType::Merged(range_types::RangeMergedCells { range, sheet })),
        Some((range, range_types::IterRowNextKind::Multiline)) => ret = Some(range_types::RangeType::Multiline(range_types::RangeMultiline { range, sheet })),
        None => (),
    }
    ret
}

// Helper function that contains the common logic for IterRow and IterRowMut next() method
fn iter_row_next_impl_mut<'a>(
    sheet: &'a mut Worksheet,
    current_row: &mut u32,
    max_row: u32,
    max_col: u32,
) -> Option<range_types::RangeTypeMut<'a>> 
{
    let mut ret: Option<range_types::RangeTypeMut<'a>> = None;
    match iter_row_next_impl_shared(sheet, current_row, max_row, max_col, "iter_row_next_impl_mut") 
    {
        Some((range, range_types::IterRowNextKind::Basic)) => ret = Some(range_types::RangeTypeMut::Basic(range_types::RangeBasicMut { range, sheet })),
        Some((range, range_types::IterRowNextKind::Merged)) => ret = Some(range_types::RangeTypeMut::Merged(range_types::RangeMergedCellsMut { range, sheet })),
        Some((range, range_types::IterRowNextKind::Multiline)) => ret = Some(range_types::RangeTypeMut::Multiline(range_types::RangeMultilineMut { range, sheet })),
        None => (),
    }
    ret
}

// =========================================================
// ITERATOR, NONE-MUTABLE, FOR LOOPING OVER WORKSHEET ROWS
// =========================================================

pub struct IterRow<'a> 
{
    pub sheet: &'a Worksheet,
    pub current_row: u32,
    pub max_row: u32,
    pub max_col: u32,
}

impl<'a> IterRow<'a> 
{
    pub fn new(sheet: &'a Worksheet, mrow: u32, mcol: u32) -> Self {
        Self {
            sheet,
            current_row: 1,
            max_row: mrow,
            max_col: mcol,
        }
    }
}

impl<'a> Iterator for IterRow<'a> 
{
    type Item = range_types::RangeType<'a>;

    fn next(&mut self) -> Option<Self::Item> 
    {
        let mut ret: Option<Self::Item> = None;
        match iter_row_next_impl(self.sheet, &mut self.current_row, self.max_row, self.max_col) {
            Some(range) => ret = Some(range),
            None => (),
        }
        ret
    }
}

// =========================================================
// ITERATOR, MUTABLE, FOR LOOPING OVER WORKSHEET ROWS
// =========================================================

pub struct IterRowMut<'a>
{
    pub sheet: &'a mut Worksheet,
    pub current_row: u32,
    pub max_row: u32,
    pub max_col: u32,
}

impl<'a> IterRowMut<'a>
{
    pub fn new(sheet: &'a mut Worksheet, mrow: u32, mcol: u32) -> Self {
        Self {
            sheet,
            current_row: 1,
            max_row: mrow,
            max_col: mcol,
        }
    }
}

//The statandart iterators, can no be mutable. So we need this specific iterator. NOTE: can't be used in for loops, but works with while!
pub trait LendingIterator 
{
    type Item<'this> where Self: 'this;

    fn next(&mut self) -> Option<Self::Item<'_>>;
}

impl LendingIterator for IterRowMut<'_> 
{
    type Item<'this> = range_types::RangeTypeMut<'this> where Self: 'this;
    
    fn next(&mut self) -> Option<Self::Item<'_>> 
    {
        let mut ret: Option<Self::Item<'_>> = None;
        match iter_row_next_impl_mut(self.sheet, &mut self.current_row, self.max_row, self.max_col) {
            Some(range) => ret = Some(range),
            None => (),
        }
        ret
    }
}

/* ---------------------------------------------------------------------------------------- */

pub fn same_types(a: &dyn range_types::IRange, b: &dyn range_types::IRange) -> bool {
    a.get_type() == b.get_type()
}