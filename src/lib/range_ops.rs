use std::vec;
use umya_spreadsheet::{Worksheet, Range, Cell};
use std::ops::{Deref, DerefMut};
use super::range_types;
// use super::common;
use log::{info, warn, error};
use regex::Regex;

//compare strings, ignoring white spaces (' ',\t, \n, \r)
pub fn cmp_strs(s1: &str, s2: &str) -> bool 
{
    let clean1: String = s1.to_lowercase()
                           .chars()
                           .filter(|c| !c.is_whitespace())
                           .collect();

    let clean2: String = s2.to_lowercase()
                           .chars()
                           .filter(|c| !c.is_whitespace())
                           .collect();

    clean1 == clean2
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

pub fn range_bounds(range: &Range) -> Option<(u32, u32, u32, u32, u32, u32)>
{
    let (Some(obr), 
         Some(oer), 
         Some(obc), 
         Some(oec)) = (range.get_coordinate_start_row(), 
                                         range.get_coordinate_end_row(), 
                                         range.get_coordinate_start_col(), 
                                         range.get_coordinate_end_col()) 
    else 
    {
        error!("Can't get range bounds! Invalid range data!");
        return None;
    };

    let brow = *obr.get_num();
    let erow = *oer.get_num();
    let bcol = *obc.get_num();
    let ecol = *oec.get_num();
    let rows = 1 + erow - brow;
    let cols = 1 + ecol - bcol;
    Some((brow, erow, bcol, ecol, rows, cols))
}

pub fn is_range_in_range(sub_range: &Range, main_range: &Range) -> bool 
{
    if let ( Some((m_start_row, m_end_row, m_start_col, m_end_col, _, _)), 
             Some((s_start_row, s_end_row, s_start_col, s_end_col, _, _)) ) = (range_bounds(main_range), range_bounds(sub_range)) 
    {
        return s_start_row >= m_start_row && s_end_row <= m_end_row && 
               s_start_col >= m_start_col && s_end_col <= m_end_col;
    }
    false
}

pub fn do_ranges_overlap(range_a: &Range, range_b: &Range) -> bool 
{
    if let ( Some((a_start_row, a_end_row, a_start_col, a_end_col, _, _)), 
             Some((b_start_row, b_end_row, b_start_col, b_end_col, _, _)) ) = (range_bounds(range_a), range_bounds(range_b)) 
    {
        let row_overlap = a_start_row <= b_end_row && a_end_row >= b_start_row;
        let col_overlap = a_start_col <= b_end_col && a_end_col >= b_start_col;

        return row_overlap && col_overlap
    }
    false
}

// Truncate to at most max_chars characters. If with_dots is true and the string was
// actually truncated, "..." is appended; short strings are returned as-is, without dots.
pub fn truncate_str(s: &str, max_chars: usize, with_dots: bool) -> String
{
    match s.char_indices().nth(max_chars)
    {
        Some((idx, _)) => { // Longer than max_chars, cut it
            let truncated = &s[..idx];
            if with_dots { format!("{}...", truncated) } else { truncated.to_string() }
        }
        None => s.to_string(), // Shorter than max_chars, return it none-changed
    }
}

pub fn limit_str(s: &str, max_chars: usize) -> String
{
    let count = s.chars().count();
    if count < max_chars
    {
        let mut res = String::from(s);
        res.extend(std::iter::repeat(' ').take(max_chars - count));
        return res;
    }
    truncate_str(s, max_chars, true)
}

pub fn range_to_string(range: &Range) -> String
{
    if let Some((rbeg, rend, cbeg, cend, rows, cols)) = range_bounds(range)
    {
        return format!("{}:{} (rows:{} cols:{})", 
                umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
                umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend), // Returns "C10"
                rows, cols)
    }
    format!("Invalid range!")
}

/// Render the cells of `range` from `sheet` into a string (one line per row:
/// a header line, then the values joined by the separator). With
/// `truncate_len = Some(t)` values are limited to `t` chars (dots if cut) and
/// the separator is `t` spaces; with `None` values are raw and tab-separated.
/// An invalid range yields a single "Invalid range from <sheet>!" line.
pub fn print_range_cells(sheet: &Worksheet, range: &Range, truncate_len: Option<u32>) -> String
{
    let mut out = String::new();
    if let Some((rbeg, rend, cbeg, cend, _, _)) = range_bounds(range)
    {
        // choose separator: when truncate_len is provided, use that many spaces,
        // otherwise keep the original tab-based separator
        let sep = match truncate_len {
            Some(tlen) => " ".repeat(tlen as usize),
            None => "\t".to_string(),
        };

        out.push_str(&format!("Sheet {} range ({}:{}) ---\n", sheet.get_name(),
            umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
            umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend)  // Returns "C10"
        ));

        // let mut coord_names = Vec::new();
        let mut cell_values = Vec::new();

        for r in rbeg..=rend 
        {
            for c in cbeg..=cend 
            {
                // let coord_str = umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r);        
                // let _truncated = match truncate_len 
                // {
                //     Some(tlen) => coord_names.push(limit_str(&coord_str, tlen as usize)),
                //     None => coord_names.push(coord_str),
                // };

                let cell_value = sheet.get_cell_value((c, r)).get_value().to_string();
                match truncate_len 
                {
                    Some(tlen) => cell_values.push(limit_str(&cell_value, tlen as usize)),
                    None => cell_values.push(cell_value) ,
                }
            }

            // info!("{}", coord_names.join(&sep));
            // coord_names.clear();

            out.push_str(&cell_values.join(&sep));
            out.push('\n');
            cell_values.clear();
        }
    }
    else
    {
        out.push_str(&format!("Invalid range from {}!\n", sheet.get_name()))
    }
    out
}

pub fn comapre_cell(
    sheet_a: &Worksheet, col_a: u32, row_a: u32,
    sheet_b: &Worksheet, col_b: u32, row_b: u32,
    strict: bool
) -> bool 
{
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
        error!("Rich text mismatch: {}:{} and {}:{}", coords_to_str(col_a, row_a), val_a, coords_to_str(col_b, row_b), val_b);
        return false;
    }
    cmp_strs(&val_a, &val_b)
}

pub fn comapre_ranges(
    sheet_a: &Worksheet, range_a: &Range,
    sheet_b: &Worksheet, range_b: &Range,
    strict: bool,
    o_allowed_rows: Option<vec::Vec<u32>>, 
    o_allowed_cols: Option<vec::Vec<u32>>,
) -> bool 
{
    if let ( Some((brow_a, erow_a, bcol_a, _, rows_a, cols_a)), 
             Some((brow_b, erow_b, bcol_b, _, rows_b, cols_b)) ) = (range_bounds(range_a), range_bounds(range_b)) 
    {
        //If the legths are different, the ranges cannot be the same
        if strict && (rows_a != rows_b || cols_a != cols_b) 
        {
            error!("Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", 
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
                // info!("Row A:{} is not in the allowed list {}!", row_num_a, _str_allowed_rows);
                continue; // skip this row if it's not in the allowed_rows list
            }

            for row_num_b in brow_b..=erow_b 
            {
                if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b)
                {
                    // info!("Row B:{} is not in the allowed list {}!", row_num_b, _str_allowed_rows);
                    continue; // skip this row if it's not in the allowed_rows list
                }

                col_match = 0;
                for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) 
                {
                    let col_num_a = bcol_a + col_a_offset;
                    let col_num_b = bcol_b + col_b_offset;

                    if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) 
                    {
                        // info!("Column A:{} or Column B:{} is not in the allowed list {}!", col_num_a, col_num_b, _str_allowed_cols);
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
            info!("Range {}:[{}, len:{}] DIFFERS FROM Range {}:[{}, len:{}]", 
                    sheet_a.get_name(), range_to_string(range_a), rows_a, sheet_b.get_name(), range_to_string(range_b), rows_b);

            return false;
        }

        info!("Range {}:[{}, len:{}] EQUALS TO Range {}:[{}, len:{}]", 
                sheet_a.get_name(), range_to_string(range_a), rows_a, sheet_b.get_name(), range_to_string(range_b), rows_b);

        return true;
    }
    false
}

pub fn append_range(
    sheet_in:   &Worksheet, 
    range_in:   &Range,
    sheet_out:  & mut Worksheet,
    clear_columns_content: &Vec<u32>
) -> bool
{
    if let Some((rbeg, rend, cbeg, cend, _, _)) = range_bounds(range_in)
    {
        // info!("clear_columns_content({:?})", clear_columns_content);

        let mut res = false;

        let merged_cells = sheet_in.get_merge_cells();

        let mut current_new_row = sheet_out.get_highest_row()+1;

        let current_old_row = current_new_row;

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

                    if clear_columns_content.contains(&col) && cell_data_type == "n"
                    {
                        dst_cell.set_value_number(0);

                        // info!("Clearing value: dst_cell({}{}).set_value({})", index_to_column(col), current_new_row, dst_cell.get_value());
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

                        // info!("Copied value: dst_cell({}{}).set_value({})", index_to_column(col), current_new_row, dst_cell.get_value());
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
            if let Some((_, _, mcbeg, mcend, mrlen, _)) = range_bounds(mrgcells)
            {
                let mrange = make_range_from_indexes(mcbeg, current_old_row, mcend, current_old_row + mrlen - 1);

                info!("{}:[{}] contains merged cells! Merge cells in {}:[{}]", sheet_in.get_name(), 
                    range_to_string(range_in), sheet_out.get_name(), range_to_string(&mrange));
                    
                sheet_out.add_merge_cells(mrange.get_range());
            }
            else 
            {
                error!("{}:[{}] contains merged cells! Invalid merged cells!", sheet_in.get_name(), range_to_string(range_in));
                res = false;
            }
        }

        return res;
    }
    false
}

pub fn accumulate_ranges(
    sheet_a: &     Worksheet, range_a: &Range,
    sheet_b: & mut Worksheet, range_b: &Range,
    pivot_cols: &vec::Vec<u32>, 
    accum_cols: &vec::Vec<u32>,
) -> bool
{
    if let ( Some((brow_a, _, bcol_a, _, rows_a, cols_a)), 
             Some((brow_b, _, bcol_b, _, rows_b, cols_b)) ) = (range_bounds(range_a), range_bounds(range_b)) 
    {
        let mut accumulated: bool = false;

        // //If the legths are different, the ranges cannot proceed with accumulation
        // if rows_a != rows_b || cols_a != cols_b 
        // {
        //     error!("Range size missmatch! {}:[{}] to {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), sheet_b.get_name(), range_to_string(range_b));
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
                    let (Some(c_a), Some(c_b)) = (cell_a.as_ref(), cell_b.as_ref()) else {
                        error!("Failed to read cells - missinga data!");
                        return false;
                    };

                    // Compare the pivot cells to determine if they are the sames
                    let val_a = c_a.get_value();
                    let val_b = c_b.get_value();

                    if cmp_strs(&val_a, &val_b) 
                    {
                        info!("Pivot {}:[{}:'{}'] EQUALS TO {}:[{}:'{}']!", sheet_a.get_name(), range_to_string(range_a), val_a, sheet_b.get_name(), range_to_string(range_b), val_b);
                        l_found_pivot = true;
                        break; //for row_offset_b in &rows_offsets_b
                    }
                }
            } //for row_offset_b in &rows_offsets_b
            
            if !l_found_pivot
            {
                if cell_a.is_none()
                {
                    info!("Pivot from {}:[{}] is NONE! {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), 
                        sheet_b.get_name(), range_to_string(range_b));
                }
                else 
                {
                    match cell_a.as_ref()
                    {
                        Some(c_a) =>
                        {
                            info!("Pivot {}:[{}:'{}'] NOT FOUND IN {}:[{}]!", sheet_a.get_name(), range_to_string(range_a), 
                                c_a.get_value(), sheet_b.get_name(), range_to_string(range_b)); 
                        }
                        None => 
                        {
                            info!("Pivot {}:[{}] not accessible!", sheet_a.get_name(), range_to_string(range_a)); 
                        }
                    }
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
                    warn!("Cell {}:{} is None!", sheet_a.get_name(), coords_to_str(col_num_a, row_num_a));
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
                    warn!("Cell {}:{} is None!", sheet_b.get_name(), coords_to_str(col_num_b, row_num_b));
                }
                break; //for col_offset_b in &cols_offsets_b
            } //for col_offset_b in &cols_offsets_b

            if !cell_a.is_none() && !cell_b.is_none() 
            {
                let (Some(c_a), Some(c_b)) = (cell_a.as_ref(), cell_b.as_ref()) else {
                    error!("Failed to read cells - data unavailable!");
                    return false;
                };

                let cell_a_coord = c_a.get_coordinate();
                let coord_a = (cell_a_coord.get_col_num().clone(), cell_a_coord.get_row_num().clone());

                let cell_b_coord = c_b.get_coordinate();
                let coord_b = (cell_b_coord.get_col_num().clone(), cell_b_coord.get_row_num().clone());

                if c_a.get_data_type() == "n" && c_b.get_data_type() == "n" 
                {
                    let type_a = c_a.get_data_type();
                    let type_b = c_b.get_data_type();

                    let val_a = c_a.get_value().parse::<f64>().unwrap_or(0.0);
                    let val_b = c_b.get_value().parse::<f64>().unwrap_or(0.0);
                    let sum: f64 = val_a + val_b;

                    info!("Accumulating {}:{} [type:'{}' value:'{}'] to {}:{} [type:'{}' value:'{}'] SUM:{}", 
                            sheet_a.get_name(), coords_to_str(coord_a.0, coord_a.1), type_a, val_a,
                            sheet_b.get_name(), coords_to_str(coord_b.0, coord_b.1), type_b, val_b, sum);

                    let q_cell_dst = sheet_b.get_cell_mut(coord_b);
                    q_cell_dst.set_value_number(sum);

                    accumulated = true;
                }
                else
                {
                    let type_a = c_a.get_data_type();
                    let type_b = c_b.get_data_type();

                    let val_a = c_a.get_value();
                    let val_b = c_b.get_value();

                    error!("Can't accumulate none-numeric values {}:{} [type:'{}' value:'{}'] to {}:{} [type:'{}' value:'{}']", 
                            sheet_a.get_name(), coords_to_str(coord_a.0, coord_a.1), type_a, val_a,
                            sheet_b.get_name(), coords_to_str(coord_b.0, coord_b.1), type_b, val_b);
                }
            }
        } //for row_offset_a in &rows_offsets_a 

        return accumulated;
    }
    false
}

pub fn offset_range(range: &Range, offsets: &Offsets) -> Result<Range, String>
{
    if let Some((br, er, bc, ec, _, _)) = range_bounds(range)
    {
        let ibr = (br as i32) + offsets.offset_beg_row;
        let ier = (er as i32) + offsets.offset_end_row;
        let ibc = (bc as i32) + offsets.offset_beg_col;
        let iec = (ec as i32) + offsets.offset_end_col;

        if 0 < ibc && 0 < ibr && 0 < iec && 0 < ier
        {
            return Ok(make_range_from_indexes(ibc as u32, ibr as u32, iec as u32, ier as u32));
        }
        return Err(format!("Invalid range (brow:{} erow:{} bcol:{} ecol:{}) found after offset!", ibr, ier, ibc, iec));
    }
    Err(format!("Invalid range!"))
}

/// Compute the sheet's effective data bounds: the highest row/column
/// containing content — a cell with a non-empty value or a formula.
/// Empty or styling-only cells do not extend the bounds.
/// Caps: max_row / max_col (0 = no limit).
fn effective_bounds_impl(sheet: &Worksheet, max_row: u32, max_col: u32) -> (u32, u32)
{
    let (mut data_max_row, mut data_max_col) = (0u32, 0u32);

    for cell in sheet.get_cell_collection()
    {
        if !cell.get_value().is_empty() || cell.is_formula()
        {
            data_max_row = data_max_row.max(*cell.get_coordinate().get_row_num());
            data_max_col = data_max_col.max(*cell.get_coordinate().get_col_num());
        }
    }

    let eff_max_row = if max_row > 0 { data_max_row.min(max_row) } else { data_max_row };
    let eff_max_col = if max_col > 0 { data_max_col.min(max_col) } else { data_max_col };

    (eff_max_row, eff_max_col)
}

pub fn effective_bounds(sheet: &Worksheet, max_row: u32, max_col: u32) -> range_types::WorkSheet<&Worksheet>
{
    let (max_row, max_col) = effective_bounds_impl(sheet, max_row, max_col);
    range_types::WorkSheet::new(sheet, max_row, max_col)
}

pub fn effective_bounds_mut(sheet: &mut Worksheet, max_row: u32, max_col: u32) -> range_types::WorkSheet<&mut Worksheet>
{
    let (max_row, max_col) = effective_bounds_impl(sheet, max_row, max_col);
    range_types::WorkSheet::new(sheet, max_row, max_col)
}

fn iter_row_next_impl_shared<'a>(
    sheet:          &Worksheet,
    current_row:    &mut u32,
    max_row:        u32,
    max_col:        u32,
    pivot_col:      u32,
    pivot_numeric:  bool,
    pivot_re:       &Regex,
    offsets:        &Offsets,
) -> Option<(Range, range_types::IterRowNextKind)> 
{
    let mut ret: Option<(Range, range_types::IterRowNextKind)> = None;

    let is_first_line = if 1 == *current_row  { true } else { false };

    if max_row >= *current_row 
    {
        let sheet_merged_cells = sheet.get_merge_cells();
        let mut cells_range: Range;

        if let Some(merged_cells) = sheet_merged_cells.iter().find(|range| is_row_in_range(*current_row, range)) 
        {
            // info!("Found merged cells range '{}'", range_to_string(&merged_cells));

            if let Some((_, merged_end_row, _, _, _, _)) = range_bounds(merged_cells)
            {
                cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, merged_end_row);

                if let Some((_, _, _, _, range_rows, _)) = range_bounds(&cells_range)
                {
                    *current_row += range_rows; //should 'current_row' be affected by the range offset

                    match offset_range(&cells_range, offsets)
                    {
                        Ok(cells_offset_range) =>
                        {
                            ret = Some((cells_offset_range, range_types::IterRowNextKind::Merged));
                        }
                        Err(err) =>
                        {
                            if is_first_line
                            {
                                ret = Some((cells_range, range_types::IterRowNextKind::Merged));
                            }
                            else
                            {
                                error!("Offset error! {} Using none-offset merged range!", err);
                            }
                        }
                    }
                }
                else
                {
                    error!("Invalid range created from merged cells!");
                }
            }
            else
            {
                error!("Invalid range bounds of merged cells!");
            }
        } 
        else if let Some(src_cell) = sheet.get_cell((pivot_col, *current_row)) //will return None if the cell is empty!
        {
            let first_cell_value = src_cell.get_value().clone();
            let first_cell_data_type = src_cell.get_data_type().to_string();
            
            cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, *current_row);

            if let Some((_, _, _, _, rows, _)) = range_bounds(&cells_range)
            {
                if pivot_numeric
                {
                    if first_cell_data_type == "n" 
                    {
                        let mut range_rows = rows;
                        let mut multiline = false;

                        let next_row = *current_row + 1;
                        for nrow in next_row..=max_row 
                        {
                            if let Some(next_cell) = sheet.get_cell((pivot_col, nrow)) 
                            {
                                let next_cell_value = next_cell.get_value().clone();
                                let next_cell_data_type = next_cell.get_data_type().to_string();
                                let pattern_found = pivot_re.is_match(&next_cell_value); //search for rows staring with symbol '-'

                                if next_cell_data_type == "s" && pattern_found
                                {
                                    cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, nrow);

                                    if let Some((_, _, _, _, new_range_rows, _)) = range_bounds(&cells_range)
                                    {
                                        range_rows = new_range_rows;
                                        multiline = true;
                                    }
                                    else
                                    {
                                        error!("Created invalid multiline range! Wrong end bounds!");
                                    }
                                }
                                else
                                {
                                    // info!("Next row {}: [numeric:{}; pattern '{}' found:{}]", nrow, next_cell_data_type, pivot_re.as_str(), pattern_found);
                                    break;
                                }
                            }
                        }

                        *current_row += range_rows; //should 'current_row' be affected by the range offset

                        match offset_range(&cells_range, offsets)
                        {
                            Ok(cells_offset_range) =>
                            {
                                if multiline 
                                {
                                    // info!("Range [{}]: from multiline cells! current_row={}", range_to_string(&cells_offset_range), current_row);
                                    ret = Some((cells_offset_range, range_types::IterRowNextKind::Multiline));
                                } 
                                else 
                                {
                                    // info!("Range [{}]: from regular cells! current_row={}", range_to_string(&cells_offset_range), current_row);
                                    ret = Some((cells_offset_range, range_types::IterRowNextKind::Basic));
                                }
                            }
                            Err(err) =>
                            {
                                if is_first_line
                                {
                                    if multiline 
                                    {
                                        // info!("Range [{}]: from multiline cells! current_row={}", range_to_string(&cells_range), current_row);
                                        ret = Some((cells_range, range_types::IterRowNextKind::Multiline));
                                    } 
                                    else 
                                    {
                                        // info!("Range [{}]: from regular cells! current_row={}", range_to_string(&cells_range), current_row);
                                        ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                                    }
                                }
                                else
                                {
                                    error!("Offset error! {} Using none-offset multiline range!", err);
                                }
                            }
                        }
                    } 
                    else 
                    {
                        // info!("Current row {} starts with unexpected type:'{}'! {}", *current_row, first_cell_data_type, range_to_string(&cells_range));

                        *current_row += 1; //should 'current_row' be affected by the range offset

                        match offset_range(&cells_range, offsets)
                        {
                            Ok(cells_offset_range) =>
                            {
                                ret = Some((cells_offset_range, range_types::IterRowNextKind::Basic));
                            }
                            Err(err) =>
                            {
                                if is_first_line
                                {
                                    ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                                }
                                else
                                {
                                    error!("Offset error! {} Using none-offset basic range!", err);
                                }
                            }
                        }
                    }
                }
                else
                {
                    let mut multiline = false;
                    if pivot_re.is_match(&first_cell_value)
                    {
                        multiline  = true;

                        let mut range_rows = rows;

                        let mut pattern_found = false;
                        let next_row = *current_row + 1;
                        for nrow in next_row..=max_row //loop untill the next pattern is found
                        {
                            if let Some(next_cell) = sheet.get_cell((pivot_col, nrow)) 
                            {
                                if pivot_re.is_match(&next_cell.get_value().clone())
                                {
                                    cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, nrow-1); //this includes the row with matched pattern, so remove one line

                                    if let Some((_, _, _, _, new_range_rows, _)) = range_bounds(&cells_range)
                                    {
                                        range_rows    = new_range_rows; //pattern found, overwrite range_rows
                                        pattern_found = true;
                                    }
                                    else
                                    {
                                        error!("Created invalid multiline range! Wrong end bounds!");
                                    }
                                    break;
                                }
                                else
                                {
                                    range_rows += 1; //count the next row we've searched.
                                }
                            }
                            else //sheet.get_cell() may return None if the cell is empty
                            {
                                range_rows += 1; //count the next row we've searched.
                            }
                        }

                        if false == pattern_found //we've not detected next pattern, so get all of the left lines
                        {
                            let last_row = *current_row + range_rows;
                            cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, last_row);
                        }

                        *current_row += range_rows; //should 'current_row' be affected by the range offset
                    }
                    else 
                    {
                        *current_row += 1; //should 'current_row' be affected by the range offset
                    }

                    match offset_range(&cells_range, offsets)
                    {
                        Ok(cells_offset_range) =>
                        {
                            if multiline 
                            {
                                // info!("Range [{}]: from multiline cells! current_row={}", range_to_string(&cells_offset_range), current_row);
                                ret = Some((cells_offset_range, range_types::IterRowNextKind::Multiline));
                            } 
                            else 
                            {
                                // info!("Range [{}]: from regular cells! current_row={}", range_to_string(&cells_offset_range), current_row);
                                ret = Some((cells_offset_range, range_types::IterRowNextKind::Basic));
                            }
                        }
                        Err(err) =>
                        {
                            if is_first_line
                            {
                                if multiline 
                                {
                                    // info!("Range [{}]: from multiline cells! current_row={}", range_to_string(&cells_range), current_row);
                                    ret = Some((cells_range, range_types::IterRowNextKind::Multiline));
                                } 
                                else 
                                {
                                    // info!("Range [{}]: from regular cells! current_row={}", range_to_string(&cells_range), current_row);
                                    ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                                }
                            }
                            else
                            {
                                error!("Offset error! {} Using none-offset multiline range for none-numeric pivot!", err);
                            }
                        }
                    }
                }
            }
            else
            {
                error!("Created invalid multiline range! Wrong start bounds!");
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
                if let Some(_cell) = sheet.get_cell((pivot_col, ridx))
                {
                    is_last_row = false;
                    break;
                }
            }

            if is_last_row
            {
                // warn!("Last row '{}' was processed!", *current_row);
            }
            else
            {
                cells_range = make_range_from_indexes(pivot_col, *current_row, pivot_col + max_col, *current_row);
                // info!("Range [{}]: from regular empty cells! current_row={}", range_to_string(&cells_range), current_row);

                match offset_range(&cells_range, offsets)
                {
                    Ok(cells_offset_range) =>
                    {
                        ret = Some((cells_offset_range, range_types::IterRowNextKind::Basic));
                    }
                    Err(err) =>
                    {
                        if is_first_line
                        {
                            ret = Some((cells_range, range_types::IterRowNextKind::Basic));
                        }
                        else
                        {
                            error!("Offset error! {} Using none-offset basic range or EOF!", err);
                        }
                    }
                }
            }

            *current_row += 1;
        }
    }
    else
    {
        warn!("Reached maximum lines to process:{}!", max_row);
    }
    ret
}

// Helper function that contains the common logic for IterRow and IterRowMut next() method
fn iter_row_next_impl<'a>(
    sheet:          &'a Worksheet,
    current_row:    &mut u32,
    max_row:        u32,
    max_col:        u32,
    pivot_col:      u32,
    pivot_numeric:  bool,
    pivot_re:       &Regex,
    offsets:        &Offsets,
) -> Option<range_types::RangeType<'a>> 
{
    let mut ret: Option<range_types::RangeType<'a>> = None;
    match iter_row_next_impl_shared(sheet, current_row, max_row, max_col, pivot_col, pivot_numeric, pivot_re, offsets) 
    {
        Some((range, range_types::IterRowNextKind::Basic)) => ret = Some(range_types::RangeType::Basic(range_types::RangeBasic::new(range,sheet))),
        Some((range, range_types::IterRowNextKind::Merged)) => ret = Some(range_types::RangeType::Merged(range_types::RangeMergedCells::new(range,sheet))),
        Some((range, range_types::IterRowNextKind::Multiline)) => ret = Some(range_types::RangeType::Multiline(range_types::RangeMultiline::new(range,sheet))),
        None => (),
    }
    ret
}

// Helper function that contains the common logic for IterRow and IterRowMut next() method
fn iter_row_next_impl_mut<'a>(
    sheet:          &'a mut Worksheet,
    current_row:    &mut u32,
    max_row:        u32,
    max_col:        u32,
    pivot_col:      u32,
    pivot_numeric:  bool,
    pivot_re:       &Regex,
    offsets:        &Offsets,
) -> Option<range_types::RangeTypeMut<'a>> 
{
    let mut ret: Option<range_types::RangeTypeMut<'a>> = None;
    match iter_row_next_impl_shared(sheet, current_row, max_row, max_col, pivot_col, pivot_numeric, pivot_re, offsets) 
    {
        Some((range, range_types::IterRowNextKind::Basic)) => ret = Some(range_types::RangeTypeMut::Basic(range_types::RangeBasicMut::new(range,sheet))),
        Some((range, range_types::IterRowNextKind::Merged)) => ret = Some(range_types::RangeTypeMut::Merged(range_types::RangeMergedCellsMut::new(range,sheet))),
        Some((range, range_types::IterRowNextKind::Multiline)) => ret = Some(range_types::RangeTypeMut::Multiline(range_types::RangeMultilineMut::new(range,sheet))),
        None => (),
    }
    ret
}

// =========================================================
// ITERATOR, GENERIC, FOR LOOPING OVER WORKSHEET ROWS
// =========================================================

pub struct Offsets
{
    pub offset_beg_col: i32,
    pub offset_end_col: i32,
    pub offset_beg_row: i32,
    pub offset_end_row: i32,
}

impl Offsets {
    pub fn new(obc: i32, oec: i32, obr: i32, oer: i32) -> Self 
    {
        Self {
            offset_beg_col: obc,
            offset_end_col: oec,
            offset_beg_row: obr,
            offset_end_row: oer,
        }
    }
}

impl Default for Offsets {
    fn default() -> Self {
        Self {
            offset_beg_col: 0,
            offset_end_col: 0,
            offset_beg_row: 0,
            offset_end_row: 0,
        }
    }
}

//Generic IterRow type
pub struct IterRowGeneric<S>
{
    pub sheet:          S,
    pub current_row:    u32,
    pub pivot_col:      u32,
    pub pivot_numeric:  bool,
    pub pivot_re:       Regex,
    pub offsets:        Offsets,
}

impl<S> IterRowGeneric<S> {
    pub fn new(sheet: S, pivot: u32, numeric: bool, re: &str, offs: Offsets) -> Result<Self, regex::Error> {
        let pivot_re = Regex::new(re)?;
        
        Ok(Self {
            sheet,
            current_row:    1,
            pivot_col:      pivot,
            pivot_numeric:  numeric,
            pivot_re,
            offsets:        offs
        })
    }
}

// =========================================================
// ITERATOR, NONE-MUTABLE, FOR LOOPING OVER WORKSHEET ROWS
// =========================================================

pub type IterRow<'a, R> = IterRowGeneric<&'a range_types::WorkSheet<R>>;

impl<'a, R> Iterator for IterRow<'a, R> 
where
    R: Deref<Target = Worksheet>,
{
    type Item = range_types::RangeType<'a>;

    fn next(&mut self) -> Option<Self::Item> 
    {
        let mut ret: Option<Self::Item> = None;
        match iter_row_next_impl(&self.sheet.sheet, &mut self.current_row, self.sheet.max_row, self.sheet.max_col, self.pivot_col, self.pivot_numeric, &self.pivot_re, &self.offsets) {
            Some(range) => ret = Some(range),
            None => (),
        }
        ret
    }
}

// =========================================================
// ITERATOR, MUTABLE, FOR LOOPING OVER WORKSHEET ROWS
// =========================================================

pub type IterRowMut<'a, R> = IterRowGeneric<&'a mut range_types::WorkSheet<R>>;

//The statandart iterators, can no be mutable. So we need this specific iterator. NOTE: can't be used in for loops, but works with while!
pub trait LendingIterator 
{
    type Item<'this> where Self: 'this;

    fn next(&mut self) -> Option<Self::Item<'_>>;
}

impl<'a, R> LendingIterator for IterRowMut<'a, R>
where
    R: DerefMut<Target = Worksheet>,
{
    type Item<'this> = range_types::RangeTypeMut<'this> where Self: 'this;
    
    fn next(&mut self) -> Option<Self::Item<'_>> 
    {
        let mut ret: Option<Self::Item<'_>> = None;
        match iter_row_next_impl_mut(&mut self.sheet.sheet, &mut self.current_row, self.sheet.max_row, self.sheet.max_col, self.pivot_col, self.pivot_numeric, &self.pivot_re, &self.offsets) {
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

/* ---------------------------------------------------------------------------------------- */

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_0001_cmp_strs_identical() {
        assert!(cmp_strs("abc", "abc"));
        assert!(cmp_strs("", ""));
        assert!(cmp_strs("Товар", "Товар"));
    }

    #[test]
    fn test_0002_cmp_strs_case_insensitive() {
        assert!(cmp_strs("ABC", "abc"));
        assert!(cmp_strs("aBc", "AbC"));
        assert!(cmp_strs("АБВ", "абв"));
        assert!(!cmp_strs("abc", "abd"));
    }

    #[test]
    fn test_0003_cmp_strs_ignores_whitespace() {
        assert!(cmp_strs("a b c", "abc"));
        assert!(cmp_strs("a\tb", "ab"));
        assert!(cmp_strs("a\nb", "ab"));
        assert!(cmp_strs("a\rb", "ab"));
        assert!(cmp_strs("  x  ", "x"));
        assert!(cmp_strs("   ", ""));
    }

    #[test]
    fn test_0004_cmp_strs_unicode() {
        assert!(cmp_strs("a\u{00A0}b", "ab"));
        assert!(cmp_strs("Товар 1 ", "товар1"));
    }

    #[test]
    fn test_0005_cmp_strs_non_whitespace_kept() {
        assert!(!cmp_strs("a-b", "ab"));
        assert!(!cmp_strs("a.b", "ab"));
        assert!(!cmp_strs("abc", "ab"));
    }

    #[test]
    fn test_0006_column_to_index_known_values() {
        assert_eq!(column_to_index("A"), 1);
        assert_eq!(column_to_index("Z"), 26);
        assert_eq!(column_to_index("AA"), 27);
        assert_eq!(column_to_index("AB"), 28);
        assert_eq!(column_to_index("AZ"), 52);
        assert_eq!(column_to_index("BA"), 53);
        assert_eq!(column_to_index("ZZ"), 702);
        assert_eq!(column_to_index("AAA"), 703);
        assert_eq!(column_to_index("ab"), 28);
    }

    #[test]
    fn test_0007_index_to_column_known_values() {
        assert_eq!(index_to_column(1), "A");
        assert_eq!(index_to_column(26), "Z");
        assert_eq!(index_to_column(27), "AA");
        assert_eq!(index_to_column(52), "AZ");
        assert_eq!(index_to_column(53), "BA");
        assert_eq!(index_to_column(702), "ZZ");
        assert_eq!(index_to_column(703), "AAA");
        assert_eq!(index_to_column(0), "");
    }

    #[test]
    fn test_0008_column_index_roundtrip() {
        for i in 1..=1000u32 {
            let col = index_to_column(i);
            assert_eq!(column_to_index(&col), i);
        }
    }

    #[test]
    #[should_panic]
    fn test_0009_column_to_index_panic_on_non_letter() {
        column_to_index("0");
    }

    #[test]
    fn test_0010_coords_to_str_known_values() {
        assert_eq!(coords_to_str(1, 1), "A1");
        assert_eq!(coords_to_str(26, 2), "Z2");
        assert_eq!(coords_to_str(27, 1), "AA1");
        assert_eq!(coords_to_str(1, 100), "A100");
        assert_eq!(coords_to_str(703, 5), "AAA5");
    }

    #[test]
    fn test_0011_make_range_from_strings_basic() {
        let r = make_range_from_strings("A1", "C3");
        assert_eq!(r.get_range(), "A1:C3");
        assert_eq!(*r.get_coordinate_start_col().unwrap().get_num(), 1);
        assert_eq!(*r.get_coordinate_start_row().unwrap().get_num(), 1);
        assert_eq!(*r.get_coordinate_end_col().unwrap().get_num(), 3);
        assert_eq!(*r.get_coordinate_end_row().unwrap().get_num(), 3);
    }

    #[test]
    fn test_0012_make_range_from_strings_single_cell() {
        let r = make_range_from_strings("B2", "B2");
        assert_eq!(r.get_range(), "B2:B2");
    }

    #[test]
    fn test_0013_make_range_from_indexes_basic() {
        let r = make_range_from_indexes(1, 1, 3, 3);
        assert_eq!(r.get_range(), "A1:C3");
    }

    #[test]
    fn test_0014_make_range_from_indexes_multi_letter_cols() {
        let r = make_range_from_indexes(27, 1, 28, 2);
        assert_eq!(r.get_range(), "AA1:AB2");
    }

    #[test]
    fn test_0015_is_row_in_range_tests() {
        let r = make_range_from_strings("A1", "C3");
        assert!(is_row_in_range(1, &r));
        assert!(is_row_in_range(2, &r));
        assert!(is_row_in_range(3, &r));
        assert!(!is_row_in_range(0, &r));
        assert!(!is_row_in_range(4, &r));
        assert!(!is_row_in_range(1, &Range::default()));
    }

    #[test]
    fn test_0016_is_col_in_range_tests() {
        let r = make_range_from_strings("A1", "C3");
        assert!(is_col_in_range(1, &r));
        assert!(is_col_in_range(2, &r));
        assert!(is_col_in_range(3, &r));
        assert!(!is_col_in_range(0, &r));
        assert!(!is_col_in_range(4, &r));
        assert!(!is_col_in_range(1, &Range::default()));
    }

    #[test]
    fn test_0017_range_bounds_valid() {
        // (brow, erow, bcol, ecol, rows, cols)
        assert_eq!(range_bounds(&make_range_from_strings("A1", "C3")), Some((1, 3, 1, 3, 3, 3)));
        assert_eq!(range_bounds(&make_range_from_strings("B2", "B2")), Some((2, 2, 2, 2, 1, 1)));
        assert_eq!(range_bounds(&make_range_from_strings("AA10", "AC12")), Some((10, 12, 27, 29, 3, 3)));
    }

    #[test]
    fn test_0018_truncate_str_no_dots() {
        assert_eq!(truncate_str("abcdef", 3, false), "abc");
        assert_eq!(truncate_str("abcd", 4, false), "abcd");
        assert_eq!(truncate_str("ab", 5, false), "ab");
        assert_eq!(truncate_str("абвгд", 2, false), "аб"); // multibyte chars
    }

    #[test]
    fn test_0019_range_to_string_valid() {
        assert_eq!(range_to_string(&make_range_from_strings("A1", "C3")), "A1:C3 (rows:3 cols:3)");
        assert_eq!(range_to_string(&make_range_from_strings("B2", "B2")), "B2:B2 (rows:1 cols:1)");
        assert_eq!(range_to_string(&make_range_from_strings("AA10", "AC12")), "AA10:AC12 (rows:3 cols:3)");
    }

    #[test]
    fn test_0020_range_to_string_invalid() {
        assert_eq!(range_to_string(&Range::default()), "Invalid range!");
        assert_eq!(range_to_string(&make_range_from_strings("A", "C")), "Invalid range!"); // no rows
    }

    #[test]
    fn test_0021_limit_str_tests() {
        assert_eq!(limit_str("abcd", 4), "abcd");     // exact fit
        assert_eq!(limit_str("abcdef", 3), "abc..."); // longer: truncate with dots
        assert_eq!(limit_str("ab", 4), "ab  ");       // shorter: pad with spaces
        assert_eq!(limit_str("", 3), "   ");          // empty: all spaces
        assert_eq!(limit_str("аб", 3), "аб ");        // multibyte chars
    }

    #[test]
    fn test_0022_truncate_str_with_dots() {
        assert_eq!(truncate_str("abcdef", 3, true), "abc...");
        assert_eq!(truncate_str("abcd", 4, true), "abcd"); // exactly max: no truncation, no dots
        assert_eq!(truncate_str("ab", 5, true), "ab");     // short: no dots
    }

    #[test]
    #[should_panic]
    fn test_0023_range_bounds_inverted_cols_underflow() {
        // Inverted range (start col 27 > end col 3): cols = 1 + ecol - bcol underflows u32.
        // Known latent bug in range_bounds, documented until fixed.
        range_bounds(&make_range_from_strings("AA10", "C12"));
    }

    #[test]
    fn test_0024_range_bounds_missing_coords() {
        assert_eq!(range_bounds(&Range::default()), None);
        // whole-column range has no rows
        assert_eq!(range_bounds(&make_range_from_strings("A", "C")), None);
    }

    #[test]
    fn test_0025_is_range_in_range_tests() {
        let main = make_range_from_strings("A1", "C3");
        assert!(is_range_in_range(&make_range_from_strings("B2", "B2"), &main));
        assert!(is_range_in_range(&main, &main));
        assert!(!is_range_in_range(&make_range_from_strings("A1", "A2"), &make_range_from_strings("A2", "C3")));
        assert!(!is_range_in_range(&make_range_from_strings("A1", "C4"), &main));
        assert!(!is_range_in_range(&main, &make_range_from_strings("B2", "B2")));
        assert!(!is_range_in_range(&main, &Range::default()));
    }

    #[test]
    fn test_0026_do_ranges_overlap_tests() {
        let a = make_range_from_strings("A1", "C3");
        assert!(do_ranges_overlap(&a, &make_range_from_strings("C3", "E5"))); // shared corner
        assert!(do_ranges_overlap(&a, &make_range_from_strings("A1", "C3"))); // identical
        assert!(do_ranges_overlap(&a, &make_range_from_strings("B2", "D4"))); // partial
        assert!(!do_ranges_overlap(&a, &make_range_from_strings("D1", "E5"))); // adjacent col
        assert!(!do_ranges_overlap(&a, &make_range_from_strings("A4", "C5"))); // adjacent row
        assert!(!do_ranges_overlap(&a, &Range::default()));
    }

    // ------------------------------------------------------------------
    // print_range_cells: it returns the rendered range as a String, so the
    // tests assert on the returned value directly — no logger involved, and
    // no shared state, so plain `cargo test` (parallel) is safe.
    // ------------------------------------------------------------------
    use umya_spreadsheet::{CellValue, Spreadsheet};

    // Write to .xlsx
    fn sheet_to_slsx(sheet: &Worksheet, path: &str) {
        let mut book = Spreadsheet::default();
        book.add_sheet(sheet.clone()).unwrap();
        umya_spreadsheet::writer::xlsx::write(&book, std::path::Path::new(path)).unwrap();
    }

    // Read back the .xlsx written by test_sheet() and return a clone of sheet `name`.
    fn sheet_from_xlsx(name: &str) -> Worksheet {
        let path = format!("/tmp/rexcell_test_sheet_{}.xlsx", name);
        let book = umya_spreadsheet::reader::xlsx::read(std::path::Path::new(&path)).unwrap();
        book.get_sheet_by_name(name).unwrap().clone()
    }

    // Create test sheet
    fn test_sheet(name: &str, rows: u32, cols: u32, long_text: bool) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        for r in 1..=rows {
            for c in 1..=cols {
                let mut value = CellValue::default();
                if !long_text {
                    value.set_value_string(format!("{}{}", index_to_column(c), r));
                }
                else {
                    value.set_value_string(format!(">>> {}{} Some long content for <<<", index_to_column(c), r));
                }
                let mut cell = Cell::default();
                cell.set_coordinate((c, r));
                cell.set_cell_value(value);
                sheet.set_cell(cell);
            }
        }
        sheet_to_slsx(&sheet, &format!("/tmp/rexcell_test_sheet_{}.xlsx", name));
        sheet
    }

    #[test]
    fn test_0027_print_range_cells_tab_mode() {
        test_sheet("T01", 5, 5, false);
        let out = print_range_cells(&sheet_from_xlsx("T01"), &make_range_from_strings("A1", "C3"), None);
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T01 range (A1:C3) ---",
            "A1\tB1\tC1",
            "A2\tB2\tC2",
            "A3\tB3\tC3",
        ]);
    }

    #[test]
    fn test_0028_print_range_cells_tab_mode() {
        test_sheet("T11", 3, 3, false);
        let out = print_range_cells(&sheet_from_xlsx("T11"), &make_range_from_strings("A1", "C3"), None);
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T11 range (A1:C3) ---",
            "A1\tB1\tC1",
            "A2\tB2\tC2",
            "A3\tB3\tC3",
        ]);
    }

    #[test]
    fn test_0029_print_range_cells_truncated_mode() {
        test_sheet("T02", 5, 5,false);
        let out = print_range_cells(&sheet_from_xlsx("T02"), &make_range_from_strings("A1", "C3"), Some(2));
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T02 range (A1:C3) ---",
            "A1  B1  C1",
            "A2  B2  C2",
            "A3  B3  C3",
        ]);
    }

    #[test]
    fn test_0030_print_range_cells_truncated_mode() {
        test_sheet("T12", 3, 3,false);
        let out = print_range_cells(&sheet_from_xlsx("T12"), &make_range_from_strings("A1", "C3"), Some(2));
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T12 range (A1:C3) ---",
            "A1  B1  C1",
            "A2  B2  C2",
            "A3  B3  C3",
        ]);
    }

    #[test]
    fn test_0031_print_range_cells_truncated_long_values() {
        test_sheet("T03", 5, 5, true);
        let out = print_range_cells(&sheet_from_xlsx("T03"), &make_range_from_strings("A1", "C3"), Some(7));
        // 7-space separator between truncated values
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T03 range (A1:C3) ---",
            ">>> A1 ...       >>> B1 ...       >>> C1 ...",
            ">>> A2 ...       >>> B2 ...       >>> C2 ...",
            ">>> A3 ...       >>> B3 ...       >>> C3 ...",
        ]);
    }

    #[test]
    fn test_0032_print_range_cells_truncated_long_values() {
        test_sheet("T13", 3, 3, true);
        let out = print_range_cells(&sheet_from_xlsx("T13"), &make_range_from_strings("A1", "C3"), Some(7));
        // 7-space separator between truncated values
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T13 range (A1:C3) ---",
            ">>> A1 ...       >>> B1 ...       >>> C1 ...",
            ">>> A2 ...       >>> B2 ...       >>> C2 ...",
            ">>> A3 ...       >>> B3 ...       >>> C3 ...",
        ]);
    }

    #[test]
    fn test_0033_print_range_cells_padded_short_values() {
        test_sheet("T05", 1, 3, false);
        let out = print_range_cells(&sheet_from_xlsx("T05"), &make_range_from_strings("A1", "C1"), Some(4));
        // each 2-char value padded to 4 ("A1  ") + 4-space separator = 6 spaces
        // between values; line ends with the trailing padding of "C1"
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Sheet T05 range (A1:C1) ---",
            "A1      B1      C1  ",
        ]);
    }

    #[test]
    fn test_0034_print_range_cells_invalid_range() {
        test_sheet("T06", 5, 5, false);
        let out = print_range_cells(&sheet_from_xlsx("T06"), &Range::default(), None);
        assert_eq!(out.lines().collect::<Vec<_>>(), vec![
            "Invalid range from T06!",
        ]);
    }

    // ------------------------------------------------------------------
    // comapre_cell: both sheets are written into one .xlsx and read back,
    // the same round-trip pattern as the print_range_cells tests above.
    // ------------------------------------------------------------------
    use umya_spreadsheet::{RichText, TextElement, Font};

    fn cmp_roundtrip(tag: &str, sheet_a: &Worksheet, sheet_b: &Worksheet) -> (Worksheet, Worksheet) {
        let mut book = Spreadsheet::default();
        book.add_sheet(sheet_a.clone()).unwrap();
        book.add_sheet(sheet_b.clone()).unwrap();
        let path = format!("/tmp/rexcell_test_cmp_{}.xlsx", tag);
        umya_spreadsheet::writer::xlsx::write(&book, std::path::Path::new(&path)).unwrap();
        let loaded = umya_spreadsheet::reader::xlsx::read(std::path::Path::new(&path)).unwrap();
        (loaded.get_sheet_by_name(&format!("{}_A", tag)).unwrap().clone(),
         loaded.get_sheet_by_name(&format!("{}_B", tag)).unwrap().clone())
    }

    fn cmp_str_sheet(name: &str, col: u32, row: u32, value: &str) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        let mut cell = Cell::default();
        cell.set_coordinate((col, row));
        let mut cv = CellValue::default();
        cv.set_value_string(value);
        cell.set_cell_value(cv);
        sheet.set_cell(cell);
        sheet
    }

    fn cmp_rich_sheet(name: &str, col: u32, row: u32, text: &str, bold: bool) -> Worksheet {
        let mut te = TextElement::default();
        te.set_text(text);
        if bold {
            let mut font = Font::default();
            font.set_bold(true);
            te.set_font(font);
        }
        let mut rt = RichText::default();
        rt.add_rich_text_elements(te);
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        let mut cell = Cell::default();
        cell.set_coordinate((col, row));
        cell.set_rich_text(rt);
        sheet.set_cell(cell);
        sheet
    }

    #[test]
    fn test_0035_compare_cell_equal_strings() {
        let (a, b) = cmp_roundtrip("CC01", &cmp_str_sheet("CC01_A", 1, 1, "hello"), &cmp_str_sheet("CC01_B", 2, 2, "hello"));
        assert!(comapre_cell(&a, 1, 1, &b, 2, 2, true));
    }

    #[test]
    fn test_0036_compare_cell_case_and_whitespace_insensitive() {
        let (a, b) = cmp_roundtrip("CC02", &cmp_str_sheet("CC02_A", 1, 1, "ABC\tX"), &cmp_str_sheet("CC02_B", 1, 1, "abc x"));
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, true));
    }

    #[test]
    fn test_0037_compare_cell_different_strings() {
        let (a, b) = cmp_roundtrip("CC03", &cmp_str_sheet("CC03_A", 1, 1, "abc"), &cmp_str_sheet("CC03_B", 1, 1, "abd"));
        assert!(!comapre_cell(&a, 1, 1, &b, 1, 1, true));
        assert!(!comapre_cell(&a, 1, 1, &b, 1, 1, false));
    }

    #[test]
    fn test_0038_compare_cell_both_missing() {
        let (a, b) = cmp_roundtrip("CC04", &cmp_str_sheet("CC04_A", 1, 1, "x"), &cmp_str_sheet("CC04_B", 1, 1, "y"));
        assert!(comapre_cell(&a, 5, 5, &b, 5, 5, true)); // neither sheet has a cell at (5,5)
    }

    #[test]
    fn test_0039_compare_cell_one_missing() {
        let (a, b) = cmp_roundtrip("CC05", &cmp_str_sheet("CC05_A", 1, 1, "x"), &cmp_str_sheet("CC05_B", 1, 1, "y"));
        assert!(!comapre_cell(&a, 1, 1, &b, 5, 5, true));
        assert!(!comapre_cell(&a, 1, 1, &b, 5, 5, false));
    }

    fn cmp_num_sheet(name: &str, col: u32, row: u32, value: f64) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        let mut cell = Cell::default();
        cell.set_coordinate((col, row));
        cell.set_value_number(value);
        sheet.set_cell(cell);
        sheet
    }

    #[test]
    fn test_0040_compare_cell_numbers_equal() {
        let (a, b) = cmp_roundtrip("CC06", &cmp_num_sheet("CC06_A", 1, 1, 1.5), &cmp_num_sheet("CC06_B", 1, 1, 1.5));
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, true));
    }

    #[test]
    fn test_0041_compare_cell_number_vs_same_string() {
        let (a, b) = cmp_roundtrip("CC06B", &cmp_num_sheet("CC06B_A", 1, 1, 1.5), &cmp_str_sheet("CC06B_B", 1, 1, "1.5"));
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, true)); // compared as rendered text
    }

    #[test]
    fn test_0042_compare_cell_rich_text_equal_strict() {
        let (a, b) = cmp_roundtrip("CC07", &cmp_rich_sheet("CC07_A", 1, 1, "text", false), &cmp_rich_sheet("CC07_B", 1, 1, "text", false));
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, true));
    }

    #[test]
    fn test_0043_compare_cell_rich_text_mismatch_strict() {
        let (a, b) = cmp_roundtrip("CC08", &cmp_rich_sheet("CC08_A", 1, 1, "text", false), &cmp_rich_sheet("CC08_B", 1, 1, "text", true));
        assert!(!comapre_cell(&a, 1, 1, &b, 1, 1, true)); // same text, different formatting
    }

    #[test]
    fn test_0044_compare_cell_rich_text_mismatch_not_strict() {
        let (a, b) = cmp_roundtrip("CC09", &cmp_rich_sheet("CC09_A", 1, 1, "text", false), &cmp_rich_sheet("CC09_B", 1, 1, "text", true));
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, false));
    }

    #[test]
    fn test_0045_compare_cell_plain_vs_rich() {
        let (a, b) = cmp_roundtrip("CC10", &cmp_str_sheet("CC10_A", 1, 1, "text"), &cmp_rich_sheet("CC10_B", 1, 1, "text", false));
        assert!(!comapre_cell(&a, 1, 1, &b, 1, 1, true)); // rich present on one side only
        assert!(comapre_cell(&a, 1, 1, &b, 1, 1, false));
    }

    // ------------------------------------------------------------------
    // comapre_ranges. NOTE: the current implementation has several latent
    // bugs which the tests below DOCUMENT rather than codify as "intended"
    // (same approach as test_0023):
    //  - strict mode only checks that sizes match; row_match is never
    //    evaluated when strict, so different content still returns true
    //  - column offsets are built as (0..=cols), i.e. one past the range
    //    end, and a row passes only with EXACTLY `cols` matches — so a
    //    perfect match FAILS in non-strict mode and exactly one mismatch
    //    per row is required
    //  - rows are compared in a cartesian product; col_match is reset for
    //    every row of B, so only B's LAST row influences the result
    //  - non-strict + a proper-allowed_rows subset can never pass, since
    //    row_match is compared against the FULL rows_a, not the allowed count
    // ------------------------------------------------------------------

    fn cmp_cells_sheet(name: &str, cells: &[(u32, u32, &str)]) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        for (c, r, v) in cells {
            let mut cell = Cell::default();
            cell.set_coordinate((*c, *r));
            let mut cv = CellValue::default();
            cv.set_value_string(*v);
            cell.set_cell_value(cv);
            sheet.set_cell(cell);
        }
        sheet
    }

    #[test]
    fn test_0046_compare_ranges_invalid_range() {
        let (a, b) = cmp_roundtrip("CR01", &cmp_str_sheet("CR01_A", 1, 1, "x"), &cmp_str_sheet("CR01_B", 1, 1, "x"));
        assert!(!comapre_ranges(&a, &Range::default(), &b, &make_range_from_strings("A1", "A1"), true, None, None));
        assert!(!comapre_ranges(&a, &make_range_from_strings("A1", "A1"), &b, &Range::default(), false, None, None));
    }

    #[test]
    fn test_0047_compare_ranges_strict_size_mismatch() {
        let (a, b) = cmp_roundtrip("CR02", &cmp_str_sheet("CR02_A", 1, 1, "x"), &cmp_str_sheet("CR02_B", 1, 1, "x"));
        assert!(!comapre_ranges(&a, &make_range_from_strings("A1", "C3"), &b, &make_range_from_strings("A1", "C2"), true, None, None));
    }

    #[test]
    fn test_0048_compare_ranges_strict_equal_content() {
        let (a, b) = cmp_roundtrip("CR03", &cmp_str_sheet("CR03_A", 1, 1, "x"), &cmp_str_sheet("CR03_B", 1, 1, "x"));
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "A2"), &b, &make_range_from_strings("A1", "A2"), true, None, None));
    }

    #[test]
    fn test_0049_compare_ranges_strict_ignores_content() {
        let (a, b) = cmp_roundtrip("CR04", &cmp_str_sheet("CR04_A", 1, 1, "aaa"), &cmp_str_sheet("CR04_B", 1, 1, "bbb"));
        // Documented current behavior: strict only checks sizes; content differences
        // do NOT fail the comparison. Expected to flip to `false` if fixed.
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "A2"), &b, &make_range_from_strings("A1", "A2"), true, None, None));
    }

    #[test]
    fn test_0050_compare_ranges_nonstrict_perfect_match_fails() {
        let (a, b) = cmp_roundtrip("CR05", &cmp_str_sheet("CR05_A", 1, 1, "x"), &cmp_str_sheet("CR05_B", 1, 1, "x"));
        // Documented current behavior: offsets run 0..=cols, so a fully equal row
        // scores cols+1 matches while the check requires exactly cols.
        // Expected to flip to `true` if fixed.
        assert!(!comapre_ranges(&a, &make_range_from_strings("A1", "A1"), &b, &make_range_from_strings("A1", "A1"), false, None, None));
    }

    #[test]
    fn test_0051_compare_ranges_nonstrict_single_mismatch_passes() {
        let a = cmp_cells_sheet("CR06_A", &[(1, 1, "x"), (2, 1, "p")]);
        let b = cmp_cells_sheet("CR06_B", &[(1, 1, "x"), (2, 1, "q")]);
        let (a, b) = cmp_roundtrip("CR06", &a, &b);
        // Exactly one mismatch (in the compared-but-out-of-range column B)
        // yields col_match == cols, so the row passes.
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "A1"), &b, &make_range_from_strings("A1", "A1"), false, None, None));
    }

    #[test]
    fn test_0052_compare_ranges_nonstrict_only_last_row_of_b() {
        // A's rows both equal B's LAST row, while B's first row differs from A.
        let a = cmp_cells_sheet("CR07_A", &[(1, 1, "same"), (2, 1, "p1"), (1, 2, "same"), (2, 2, "p2")]);
        let b = cmp_cells_sheet("CR07_B", &[(1, 1, "other"), (2, 1, "q1"), (1, 2, "same"), (2, 2, "q2")]);
        let (a, b) = cmp_roundtrip("CR07", &a, &b);
        // Documented current behavior: cartesian row loop, only B's last row is
        // effectively compared. Expected to flip to `false` if fixed.
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "A2"), &b, &make_range_from_strings("A1", "A2"), false, None, None));

        // Control: B's last row differs -> no row of A matches.
        let mut a2 = a.clone();
        a2.set_name("CR07C_A");
        let b2 = cmp_cells_sheet("CR07C_B", &[(1, 1, "other"), (2, 1, "q1"), (1, 2, "nope"), (2, 2, "q2")]);
        let (_, b2) = cmp_roundtrip("CR07C", &a2, &b2);
        assert!(!comapre_ranges(&a, &make_range_from_strings("A1", "A2"), &b2, &make_range_from_strings("A1", "A2"), false, None, None));
    }

    #[test]
    fn test_0053_compare_ranges_allowed_rows() {
        // Every row of A equals B's last row (see test_0052); col 2 always differs.
        let a = cmp_cells_sheet("CR08_A", &[(1, 1, "c"), (2, 1, "x1"), (1, 2, "c"), (2, 2, "x2"), (1, 3, "c"), (2, 3, "x3")]);
        let b = cmp_cells_sheet("CR08_B", &[(1, 1, "a"), (2, 1, "y1"), (1, 2, "b"), (2, 2, "y2"), (1, 3, "c"), (2, 3, "y3")]);
        let (a, b) = cmp_roundtrip("CR08", &a, &b);

        // Full allowed list: all 3 rows pass -> row_match == rows_a.
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "A3"), &b, &make_range_from_strings("A1", "A3"), false, Some(vec![1, 2, 3]), None));

        // Documented current behavior: a proper subset can never pass, because
        // row_match (<= allowed rows) is compared against the FULL rows_a.
        // Expected to flip to `true` if fixed.
        assert!(!comapre_ranges(&a, &make_range_from_strings("A1", "A3"), &b, &make_range_from_strings("A1", "A3"), false, Some(vec![3]), None));
    }

    #[test]
    fn test_0054_compare_ranges_allowed_cols_skip_counts_as_match() {
        // Range A1:B1 (cols=2) compares 3 columns; allowed_cols=[1] skips cols 2 and 3,
        // and skipped columns COUNT as matches.
        let a = cmp_cells_sheet("CR09_A", &[(1, 1, "same"), (2, 1, "b1"), (3, 1, "c1")]);

        let b = cmp_cells_sheet("CR09_B", &[(1, 1, "diff"), (2, 1, "d2"), (3, 1, "e3")]);
        let (a, b) = cmp_roundtrip("CR09", &a, &b);
        // col 1 mismatches (0) + 2 skipped = 2 == cols -> passes
        assert!(comapre_ranges(&a, &make_range_from_strings("A1", "B1"), &b, &make_range_from_strings("A1", "B1"), false, None, Some(vec![1])));

        let b2 = cmp_cells_sheet("CR10_B", &[(1, 1, "same"), (2, 1, "d2"), (3, 1, "e3")]);
        let (a2, b2) = cmp_roundtrip("CR10", &cmp_cells_sheet("CR10_A", &[(1, 1, "same"), (2, 1, "b1"), (3, 1, "c1")]), &b2);
        // col 1 matches (1) + 2 skipped = 3 != cols -> fails
        assert!(!comapre_ranges(&a2, &make_range_from_strings("A1", "B1"), &b2, &make_range_from_strings("A1", "B1"), false, None, Some(vec![1])));
    }

    // ------------------------------------------------------------------
    // append_range. Works fully in-memory (mutates sheet_out), so the input
    // is built with the string-cell helper above and the output sheet is
    // asserted on directly — no xlsx roundtrip needed.
    // ------------------------------------------------------------------

    #[test]
    fn test_0055_append_range_invalid_range() {
        let sheet_in = cmp_cells_sheet("AR01", &[(1, 1, "x")]);
        let mut out = Worksheet::default();
        assert!(!append_range(&sheet_in, &Range::default(), &mut out, &vec![]));
        assert!(out.get_cell((1, 1)).is_none());
    }

    #[test]
    fn test_0056_append_range_no_cells_in_range() {
        let sheet_in = cmp_cells_sheet("AR02", &[(1, 1, "x")]);
        let mut out = Worksheet::default();
        assert!(!append_range(&sheet_in, &make_range_from_strings("C5", "D6"), &mut out, &vec![]));
        assert!(out.get_cell((3, 5)).is_none());
    }

    #[test]
    fn test_0057_append_range_basic_block() {
        let sheet_in = cmp_cells_sheet("AR03", &[(1, 1, "a"), (2, 1, "b"), (1, 2, "c"), (2, 2, "d")]);
        let mut out = Worksheet::default();
        assert!(append_range(&sheet_in, &make_range_from_strings("A1", "B2"), &mut out, &vec![]));
        assert_eq!(out.get_cell_value((1, 1)).get_value(), "a");
        assert_eq!(out.get_cell_value((2, 1)).get_value(), "b");
        assert_eq!(out.get_cell_value((1, 2)).get_value(), "c");
        assert_eq!(out.get_cell_value((2, 2)).get_value(), "d");
    }

    #[test]
    fn test_0058_append_range_after_existing_skips_empty_rows() {
        let sheet_in = cmp_cells_sheet("AR04", &[(1, 1, "x"), (2, 1, "y"), (1, 3, "z")]); // input row 2 is empty
        let mut out = cmp_cells_sheet("AR04_OUT", &[(1, 1, "old")]);
        assert!(append_range(&sheet_in, &make_range_from_strings("A1", "B3"), &mut out, &vec![]));
        // append starts at out row 2 (after "old"); empty input row 2 is compacted away
        assert_eq!(out.get_cell_value((1, 2)).get_value(), "x");
        assert_eq!(out.get_cell_value((2, 2)).get_value(), "y");
        assert_eq!(out.get_cell_value((1, 3)).get_value(), "z");
        assert!(out.get_cell((1, 4)).is_none());
    }

    #[test]
    fn test_0059_append_range_clear_columns_content() {
        let mut sheet_in = cmp_cells_sheet("AR05", &[(1, 1, "keep"), (3, 1, "text")]);
        let mut ncell = Cell::default();
        ncell.set_coordinate((2, 1));
        ncell.set_value_number(2.5);
        sheet_in.set_cell(ncell);
        let mut out = Worksheet::default();
        assert!(append_range(&sheet_in, &make_range_from_strings("A1", "C1"), &mut out, &vec![2, 3]));
        // col 2: numeric cleared to 0; col 3: string copied as-is
        assert_eq!(out.get_cell_value((1, 1)).get_value(), "keep");
        assert_eq!(out.get_cell_value((2, 1)).get_value(), "0");
        assert_eq!(out.get_cell((2, 1)).unwrap().get_data_type(), "n");
        assert_eq!(out.get_cell_value((3, 1)).get_value(), "text");
    }

    #[test]
    fn test_0060_append_range_preserves_rich_text() {
        let sheet_in = cmp_rich_sheet("AR06", 1, 1, "hello", true);
        let mut out = Worksheet::default();
        assert!(append_range(&sheet_in, &make_range_from_strings("A1", "A1"), &mut out, &vec![]));
        let dst = out.get_cell((1, 1)).unwrap();
        assert_eq!(dst.get_value(), "hello");
        // rich text (not plain value) was copied
        let rich = dst.get_cell_value().get_raw_value().get_rich_text().unwrap();
        assert_eq!(rich.get_text(), "hello");
    }

    #[test]
    fn test_0061_append_range_merges_cells() {
        let mut sheet_in = cmp_cells_sheet("AR07", &[(1, 1, "m"), (1, 2, "n")]);
        sheet_in.add_merge_cells("A1:B2");
        let mut out = cmp_cells_sheet("AR07_OUT", &[(1, 1, "old")]);
        assert!(append_range(&sheet_in, &make_range_from_strings("A1", "B2"), &mut out, &vec![]));
        assert_eq!(out.get_cell_value((1, 2)).get_value(), "m");
        assert_eq!(out.get_cell_value((1, 3)).get_value(), "n");
        // merge is re-applied at the append offset (out rows 2..3)
        let merged: Vec<String> = out.get_merge_cells().iter().map(|r| r.get_range().to_string()).collect();
        assert!(merged.contains(&"A2:B3".to_string()));
    }

    #[test]
    fn test_0062_append_range_false_if_last_row_empty() {
        let sheet_in = cmp_cells_sheet("AR08", &[(1, 1, "x")]);
        let mut out = Worksheet::default();
        // Documented current behavior: res is taken from the LAST row processed,
        // so a range whose final row has no cells returns false even though the
        // earlier rows were copied. Expected to flip to `true` if fixed.
        assert!(!append_range(&sheet_in, &make_range_from_strings("A1", "A2"), &mut out, &vec![]));
        assert_eq!(out.get_cell_value((1, 1)).get_value(), "x");
    }

    // ------------------------------------------------------------------
    // accumulate_ranges. Mutates sheet_b in-memory (no roundtrip needed).
    // For each row of A: the pivot cell (first column in pivot_cols, or the
    // range's first column when pivot_cols is empty) is looked up in B from
    // its FIRST row downwards; on a match the first accum column of the A
    // row is added to the matched B row. Quirks documented below:
    //  - every A row matches B's FIRST row with an equal pivot, so several
    //    A rows accumulate into the same B row
    //  - the first A row whose pivot is not found in B BREAKS the whole
    //    loop; later A rows are not processed
    // ------------------------------------------------------------------

    fn acc_sheet(name: &str, strings: &[(u32, u32, &str)], numbers: &[(u32, u32, f64)]) -> Worksheet {
        let mut sheet = cmp_cells_sheet(name, strings);
        for (c, r, v) in numbers {
            let mut cell = Cell::default();
            cell.set_coordinate((*c, *r));
            cell.set_value_number(*v);
            sheet.set_cell(cell);
        }
        sheet
    }

    #[test]
    fn test_0063_accumulate_ranges_invalid_range() {
        let sheet_a = acc_sheet("AC01", &[(1, 1, "k")], &[]);
        let mut out_b = acc_sheet("AC01B", &[(1, 1, "k")], &[]);
        assert!(!accumulate_ranges(&sheet_a, &Range::default(), &mut out_b, &make_range_from_strings("A1", "A1"), &vec![1], &vec![1]));
        assert!(!accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "A1"), &mut out_b, &Range::default(), &vec![1], &vec![1]));
    }

    #[test]
    fn test_0064_accumulate_ranges_basic_same_row() {
        let sheet_a = acc_sheet("AC02", &[(1, 1, "k")], &[(2, 1, 1.5)]);
        let mut sheet_b = acc_sheet("AC02B", &[(1, 1, "k")], &[(2, 1, 2.0)]);
        assert!(accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B1"), &mut sheet_b, &make_range_from_strings("A1", "B1"), &vec![1], &vec![2]));
        assert_eq!(sheet_b.get_cell_value((2, 1)).get_value(), "3.5");
        assert_eq!(sheet_b.get_cell((2, 1)).unwrap().get_data_type(), "n");
        // pivot cell untouched
        assert_eq!(sheet_b.get_cell_value((1, 1)).get_value(), "k");
    }

    #[test]
    fn test_0065_accumulate_ranges_pivot_found_on_later_row_of_b() {
        let sheet_a = acc_sheet("AC03", &[(1, 1, "k")], &[(2, 1, 1.0)]);
        let mut sheet_b = acc_sheet("AC03B", &[(1, 1, "other"), (1, 2, "k")], &[(2, 2, 5.0)]);
        assert!(accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B1"), &mut sheet_b, &make_range_from_strings("A1", "B2"), &vec![1], &vec![2]));
        assert_eq!(sheet_b.get_cell_value((2, 2)).get_value(), "6");
        assert!(sheet_b.get_cell((2, 1)).is_none());
    }

    #[test]
    fn test_0066_accumulate_ranges_all_a_rows_hit_first_b_match() {
        // Both A rows have pivot "k"; B has "k" on both rows.
        let sheet_a = acc_sheet("AC04", &[(1, 1, "k"), (1, 2, "k")], &[(2, 1, 1.0), (2, 2, 1.0)]);
        let mut sheet_b = acc_sheet("AC04B", &[(1, 1, "k"), (1, 2, "k")], &[(2, 1, 10.0), (2, 2, 20.0)]);
        assert!(accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B2"), &mut sheet_b, &make_range_from_strings("A1", "B2"), &vec![1], &vec![2]));
        // Documented current behavior: each A row is matched against B's FIRST
        // equal pivot row, so B row 1 gets 10+1+1 and B row 2 is untouched.
        // Expected to change if matching is made positional.
        assert_eq!(sheet_b.get_cell_value((2, 1)).get_value(), "12");
        assert_eq!(sheet_b.get_cell_value((2, 2)).get_value(), "20");
    }

    #[test]
    fn test_0067_accumulate_ranges_missing_pivot_breaks_loop() {
        // A row 1 pivot "missing" is not in B -> the whole loop breaks, so
        // A row 2 (which WOULD match B) is never processed.
        let sheet_a = acc_sheet("AC05", &[(1, 1, "missing"), (1, 2, "k")], &[(2, 1, 1.0), (2, 2, 1.0)]);
        let mut sheet_b = acc_sheet("AC05B", &[(1, 1, "k")], &[(2, 1, 5.0)]);
        assert!(!accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B2"), &mut sheet_b, &make_range_from_strings("A1", "B1"), &vec![1], &vec![2]));
        assert_eq!(sheet_b.get_cell_value((2, 1)).get_value(), "5");
    }

    #[test]
    fn test_0068_accumulate_ranges_non_numeric_not_accumulated() {
        let sheet_a = acc_sheet("AC06", &[(1, 1, "k"), (2, 1, "txt")], &[]);
        let mut sheet_b = acc_sheet("AC06B", &[(1, 1, "k")], &[(2, 1, 5.0)]);
        assert!(!accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B1"), &mut sheet_b, &make_range_from_strings("A1", "B1"), &vec![1], &vec![2]));
        assert_eq!(sheet_b.get_cell_value((2, 1)).get_value(), "5");
    }

    #[test]
    fn test_0069_accumulate_ranges_empty_pivot_cols_uses_first_column() {
        let sheet_a = acc_sheet("AC07", &[(1, 1, "k")], &[(2, 1, 1.0)]);
        let mut sheet_b = acc_sheet("AC07B", &[(1, 1, "k")], &[(2, 1, 2.0)]);
        assert!(accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B1"), &mut sheet_b, &make_range_from_strings("A1", "B1"), &vec![], &vec![2]));
        assert_eq!(sheet_b.get_cell_value((2, 1)).get_value(), "3");
    }

    #[test]
    fn test_0070_accumulate_ranges_missing_accum_cell_in_b() {
        let sheet_a = acc_sheet("AC08", &[(1, 1, "k")], &[(2, 1, 1.0)]);
        let mut sheet_b = acc_sheet("AC08B", &[(1, 1, "k")], &[]); // no B1
        assert!(!accumulate_ranges(&sheet_a, &make_range_from_strings("A1", "B1"), &mut sheet_b, &make_range_from_strings("A1", "B1"), &vec![1], &vec![2]));
        assert!(sheet_b.get_cell((2, 1)).is_none());
    }

    // ------------------------------------------------------------------
    // offset_range: shifts a range by Offsets::new(obc, oec, obr, oer)
    // and fails if any shifted bound lands on or beyond 0.
    // ------------------------------------------------------------------

    #[test]
    fn test_0071_offset_range_zero_offsets_identity() {
        let off = Offsets::default();
        assert_eq!(offset_range(&make_range_from_strings("A1", "C3"), &off).unwrap().get_range(), "A1:C3");
        assert_eq!(offset_range(&make_range_from_strings("B2", "B2"), &off).unwrap().get_range(), "B2:B2");
    }

    #[test]
    fn test_0072_offset_range_shift_down_right() {
        assert_eq!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(1, 1, 1, 1)).unwrap().get_range(), "B2:D4");
    }

    #[test]
    fn test_0073_offset_range_shift_up_left_to_boundary() {
        // lands exactly on row 1 / col 1, which is allowed (bounds must be > 0)
        assert_eq!(offset_range(&make_range_from_strings("B2", "C5"), &Offsets::new(-1, -1, -1, -1)).unwrap().get_range(), "A1:B4");
    }

    #[test]
    fn test_0074_offset_range_asymmetric() {
        assert_eq!(offset_range(&make_range_from_strings("A2", "C5"), &Offsets::new(0, 1, 1, 2)).unwrap().get_range(), "A3:D7");
        // collapsing the range to a single column
        assert_eq!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(2, 0, 0, 0)).unwrap().get_range(), "C1:C3");
    }

    #[test]
    fn test_0075_offset_range_multi_letter_cols() {
        assert_eq!(offset_range(&make_range_from_strings("AA10", "AC12"), &Offsets::new(1, 1, -1, -1)).unwrap().get_range(), "AB9:AD11");
    }

    #[test]
    fn test_0076_offset_range_bound_hits_zero() {
        assert!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(-1, 0, 0, 0)).is_err()); // bcol -> 0
        assert!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(0, 0, -1, 0)).is_err()); // brow -> 0
        assert!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(0, 0, 0, -3)).is_err()); // erow -> 0
        assert!(offset_range(&make_range_from_strings("A1", "C3"), &Offsets::new(0, -3, 0, 0)).is_err()); // ecol -> 0
    }

    #[test]
    fn test_0077_offset_range_negative_bounds() {
        assert!(offset_range(&make_range_from_strings("A1", "B2"), &Offsets::new(-1, -1, -1, -1)).is_err());
        assert!(offset_range(&make_range_from_strings("A1", "B2"), &Offsets::new(-10, -10, -10, -10)).is_err());
        let e = offset_range(&make_range_from_strings("A1", "B2"), &Offsets::new(-1, -1, -1, -1)).unwrap_err();
        assert!(e.contains("found after offset"), "unexpected error: {}", e);
    }

    #[test]
    fn test_0078_offset_range_invalid_input() {
        let off = Offsets::default();
        assert_eq!(offset_range(&Range::default(), &off).unwrap_err(), "Invalid range!");
        assert_eq!(offset_range(&make_range_from_strings("A", "C"), &off).unwrap_err(), "Invalid range!"); // whole-column range, no rows
    }

    // ------------------------------------------------------------------
    // iter_row_next_impl_shared / iter_row_next_impl / iter_row_next_impl_mut
    // Sheets are built directly in memory (no xlsx round-trip).
    // Common params unless noted: max_row=10, max_col=2 (cols A..C),
    // pivot_col=1 (col A), zero offsets, pivot_re = "^-".
    // ------------------------------------------------------------------
    use range_types::{IRange, IRangeMut};

    fn ir_sheet(name: &str, str_cells: &[(u32, u32, &str)], num_cells: &[(u32, u32, f64)]) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        for &(c, r, v) in str_cells {
            let mut value = CellValue::default();
            value.set_value_string(v.to_string());
            let mut cell = Cell::default();
            cell.set_coordinate((c, r));
            cell.set_cell_value(value);
            sheet.set_cell(cell);
        }
        for &(c, r, v) in num_cells {
            let mut cell = Cell::default();
            cell.set_coordinate((c, r));
            cell.set_value_number(v);
            sheet.set_cell(cell);
        }
        sheet
    }

    // calls iter_row_next_impl_shared, returns (result, updated current_row)
    fn shared_next(sheet: &Worksheet, row: u32, max_row: u32, max_col: u32, pivot_col: u32, numeric: bool, re: &Regex, offsets: &Offsets)
        -> (Option<(Range, range_types::IterRowNextKind)>, u32)
    {
        let mut cr = row;
        let res = iter_row_next_impl_shared(sheet, &mut cr, max_row, max_col, pivot_col, numeric, re, offsets);
        (res, cr)
    }

    fn assert_shared(res: &Option<(Range, range_types::IterRowNextKind)>, range: &str, kind: &range_types::IterRowNextKind) {
        let (r, k) = res.as_ref().expect("expected Some((range, kind))");
        assert_eq!(r.get_range(), range);
        assert_eq!(k, kind);
    }

    #[test]
    fn test_0079_iter_shared_basic_single_row() {
        let sheet = ir_sheet("IR01", &[(1, 1, "hello"), (1, 2, "world"), (1, 3, "again")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::default());
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0080_iter_shared_last_row_and_eof() {
        let sheet = ir_sheet("IR02", &[(1, 1, "a"), (1, 2, "b"), (1, 3, "c")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let def = Offsets::default();
        let (res, cr) = shared_next(&sheet, 3, 10, 2, 1, false, &re, &def);
        assert_shared(&res, "A3:C3", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 4);
        // no more cells after row 3 -> end of sheet, row counter still advances
        let (res, cr) = shared_next(&sheet, 4, 10, 2, 1, false, &re, &def);
        assert!(res.is_none());
        assert_eq!(cr, 5);
    }

    #[test]
    fn test_0081_iter_shared_empty_row_with_content_after() {
        let sheet = ir_sheet("IR03", &[(1, 3, "x")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::default());
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0082_iter_shared_multiline_until_next_pattern() {
        let sheet = ir_sheet("IR04", &[(1, 1, "-start"), (1, 2, "detail"), (1, 3, "-next")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 4, 2, 1, false, &re, &Offsets::default());
        // range stops before the next pattern row (row 3)
        assert_shared(&res, "A1:C2", &range_types::IterRowNextKind::Multiline);
        assert_eq!(cr, 3);
    }

    #[test]
    fn test_0083_iter_shared_multiline_no_next_pattern_goes_to_max_row() {
        let sheet = ir_sheet("IR05", &[(1, 1, "-only"), (1, 2, "a"), (1, 3, "b")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 4, 2, 1, false, &re, &Offsets::default());
        // no next pattern: range extends one row past max_row
        assert_shared(&res, "A1:C5", &range_types::IterRowNextKind::Multiline);
        assert_eq!(cr, 5);
    }

    #[test]
    fn test_0084_iter_shared_numeric_pivot_extends_to_matching_row() {
        let sheet = ir_sheet("IR06", &[(1, 2, "-cont"), (1, 3, "plain")], &[(1, 1, 1.0)]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, true, &re, &Offsets::default());
        assert_shared(&res, "A1:C2", &range_types::IterRowNextKind::Multiline);
        assert_eq!(cr, 3);
    }

    #[test]
    fn test_0085_iter_shared_numeric_pivot_no_match_stays_basic() {
        let sheet = ir_sheet("IR07", &[(1, 2, "plain")], &[(1, 1, 1.0)]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, true, &re, &Offsets::default());
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0086_iter_shared_numeric_mode_string_cell_is_basic() {
        let sheet = ir_sheet("IR07B", &[(1, 1, "txt")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, true, &re, &Offsets::default());
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0087_iter_shared_merged_cells() {
        let mut sheet = ir_sheet("IR08", &[(1, 1, "a")], &[]);
        sheet.add_merge_cells("A1:C3");
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::default());
        assert_shared(&res, "A1:C3", &range_types::IterRowNextKind::Merged);
        assert_eq!(cr, 4);
    }

    #[test]
    fn test_0088_iter_shared_merged_range_starting_later() {
        let mut sheet = ir_sheet("IR09", &[(1, 1, "a")], &[]);
        sheet.add_merge_cells("A2:C3");
        let re = Regex::new(r"^-").unwrap();
        // row 1 is outside the merged range -> regular basic row
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::default());
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
        // row 2 is inside the merged range: starts at current row, ends at merged end row
        let (res, cr) = shared_next(&sheet, 2, 10, 2, 1, false, &re, &Offsets::default());
        assert_shared(&res, "A2:C3", &range_types::IterRowNextKind::Merged);
        assert_eq!(cr, 4);
    }

    #[test]
    fn test_0089_iter_shared_max_row_reached() {
        let sheet = ir_sheet("IR10", &[(1, 1, "a")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 5, 4, 2, 1, false, &re, &Offsets::default());
        assert!(res.is_none());
        assert_eq!(cr, 5);
    }

    #[test]
    fn test_0090_iter_shared_offset_shifts_range() {
        let sheet = ir_sheet("IR11", &[(1, 1, "a")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::new(1, 1, 0, 0));
        assert_shared(&res, "B1:D1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0091_iter_shared_offset_error_first_line_falls_back() {
        let sheet = ir_sheet("IR12", &[(1, 1, "a")], &[]);
        let re = Regex::new(r"^-").unwrap();
        // bcol 1 + (-1) = 0 -> invalid, but on the first line the no-offset range is returned
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::new(-1, -1, 0, 0));
        assert_shared(&res, "A1:C1", &range_types::IterRowNextKind::Basic);
        assert_eq!(cr, 2);
    }

    #[test]
    fn test_0092_iter_shared_offset_error_later_row_returns_none() {
        let sheet = ir_sheet("IR13", &[(1, 1, "a"), (1, 2, "b")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 2, 10, 2, 1, false, &re, &Offsets::new(-1, -1, 0, 0));
        assert!(res.is_none());
        assert_eq!(cr, 3);
    }

    #[test]
    fn test_0093_iter_shared_multiline_with_row_offset() {
        let sheet = ir_sheet("IR14", &[(1, 1, "-s"), (1, 2, "x"), (1, 3, "-n")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let (res, cr) = shared_next(&sheet, 1, 10, 2, 1, false, &re, &Offsets::new(0, 0, 1, 1));
        // base range A1:C2 shifted down by one row
        assert_shared(&res, "A2:C3", &range_types::IterRowNextKind::Multiline);
        assert_eq!(cr, 3);
    }

    // ------------------------------------------------------------------
    // iter_row_next_impl: wraps the shared impl into Option<RangeType>
    // ------------------------------------------------------------------

    #[test]
    fn test_0094_iter_row_next_impl_kinds() {
        let re = Regex::new(r"^-").unwrap();
        let def = Offsets::default();

        let sheet = ir_sheet("IR15", &[(1, 1, "hello")], &[]);
        let mut cr = 1;
        let rt = iter_row_next_impl(&sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Basic);
        assert_eq!(rt.get_range().get_range(), "A1:C1");
        assert_eq!(cr, 2);

        let sheet = ir_sheet("IR16", &[(1, 1, "-s"), (1, 2, "x"), (1, 3, "-n")], &[]);
        let mut cr = 1;
        let rt = iter_row_next_impl(&sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Multiline);
        assert_eq!(rt.get_range().get_range(), "A1:C2");
        assert_eq!(cr, 3);

        let mut sheet = ir_sheet("IR17", &[(1, 1, "a")], &[]);
        sheet.add_merge_cells("A1:C3");
        let mut cr = 1;
        let rt = iter_row_next_impl(&sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Merged);
        assert_eq!(rt.get_range().get_range(), "A1:C3");
        assert_eq!(cr, 4);
    }

    #[test]
    fn test_0095_iter_row_next_impl_eof_returns_none() {
        let sheet = ir_sheet("IR18", &[], &[]);
        let re = Regex::new(r"^-").unwrap();
        let mut cr = 1;
        assert!(iter_row_next_impl(&sheet, &mut cr, 10, 2, 1, false, &re, &Offsets::default()).is_none());
        assert_eq!(cr, 2);
    }

    // ------------------------------------------------------------------
    // iter_row_next_impl_mut: wraps the shared impl into Option<RangeTypeMut>
    // ------------------------------------------------------------------

    #[test]
    fn test_0096_iter_row_next_impl_mut_kinds() {
        let re = Regex::new(r"^-").unwrap();
        let def = Offsets::default();

        let mut sheet = ir_sheet("IR19", &[(1, 1, "hello")], &[]);
        let mut cr = 1;
        let rt = iter_row_next_impl_mut(&mut sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Basic);
        assert_eq!(rt.get_range().get_range(), "A1:C1");
        assert_eq!(cr, 2);

        let mut sheet = ir_sheet("IR20", &[(1, 1, "-s"), (1, 2, "x"), (1, 3, "-n")], &[]);
        let mut cr = 1;
        let rt = iter_row_next_impl_mut(&mut sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Multiline);
        assert_eq!(rt.get_range().get_range(), "A1:C2");

        let mut sheet = ir_sheet("IR21", &[(1, 1, "a")], &[]);
        sheet.add_merge_cells("A1:C3");
        let mut cr = 1;
        let rt = iter_row_next_impl_mut(&mut sheet, &mut cr, 10, 2, 1, false, &re, &def).unwrap();
        assert_eq!(rt.get_type(), range_types::IterRowNextKind::Merged);
        assert_eq!(rt.get_range().get_range(), "A1:C3");
    }

    #[test]
    fn test_0097_iter_row_next_impl_mut_sheet_is_mutable() {
        let mut sheet = ir_sheet("IR22", &[(1, 1, "a")], &[]);
        let re = Regex::new(r"^-").unwrap();
        let mut cr = 1;
        let mut rt = iter_row_next_impl_mut(&mut sheet, &mut cr, 10, 2, 1, false, &re, &Offsets::default()).unwrap();
        rt.get_sheet_mut().set_name("renamed");
        drop(rt);
        assert_eq!(sheet.get_name(), "renamed");
    }

    // ------------------------------------------------------------------
    // IterRow / IterRowMut: iterators built on the impls above.
    // IterRow is a standard Iterator (works in for-loops);
    // IterRowMut is a LendingIterator (while-let only).
    // ------------------------------------------------------------------

    #[test]
    fn test_0098_iter_row_new_invalid_regex() {
        let sheet = ir_sheet("IR23", &[], &[]);
        let wsheet = effective_bounds(&sheet, 0, 0);
        assert!(IterRow::new(&wsheet, 1, false, "[", Offsets::default()).is_err());
    }

    #[test]
    fn test_0099_iter_row_collects_basic_rows() {
        let sheet = ir_sheet("IR24", &[(1, 1, "a"), (1, 2, "b"), (1, 3, "c")], &[]);
        let wsheet = effective_bounds(&sheet, 0, 0);
        let mut iter = IterRow::new(&wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        for rt in iter.by_ref() {
            ranges.push(rt.get_range().get_range());
        }
        assert_eq!(ranges, vec!["A1:B1".to_string(), "A2:B2".to_string(), "A3:B3".to_string()]);
        assert_eq!(iter.current_row, 4);
    }

    #[test]
    fn test_0100_iter_row_multiline_kinds() {
        let sheet = ir_sheet("IR25", &[(1, 1, "-s"), (1, 2, "x"), (1, 3, "-n"), (1, 4, "t")], &[]);
        let wsheet = effective_bounds(&sheet, 0, 0);
        let iter = IterRow::new(&wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        let mut kinds = vec![];
        for rt in iter {
            kinds.push(rt.get_type());
            ranges.push(rt.get_range().get_range());
        }
        assert_eq!(ranges, vec!["A1:B2".to_string(), "A3:B5".to_string()]);
        assert_eq!(kinds, vec![range_types::IterRowNextKind::Multiline, range_types::IterRowNextKind::Multiline]);
    }

    #[test]
    fn test_0101_iter_row_merged_in_middle() {
        let mut sheet = ir_sheet("IR26", &[(1, 1, "a"), (1, 5, "e")], &[]);
        sheet.add_merge_cells("A2:C4");
        let wsheet = effective_bounds(&sheet, 0, 0);
        let iter = IterRow::new(&wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        let mut kinds = vec![];
        for rt in iter {
            kinds.push(rt.get_type());
            ranges.push(rt.get_range().get_range());
        }
        assert_eq!(ranges, vec!["A1:B1".to_string(), "A2:B4".to_string(), "A5:B5".to_string()]);
        assert_eq!(kinds, vec![range_types::IterRowNextKind::Basic, range_types::IterRowNextKind::Merged, range_types::IterRowNextKind::Basic]);
    }

    #[test]
    fn test_0102_iter_row_empty_sheet() {
        let sheet = ir_sheet("IR27", &[], &[]);
        let wsheet = effective_bounds(&sheet, 0, 0);
        let mut iter = IterRow::new(&wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        for rt in iter.by_ref() {
            ranges.push(rt.get_range().get_range());
        }
        assert!(ranges.is_empty());
        assert_eq!(iter.current_row, 1);
    }

    #[test]
    fn test_0103_iter_row_mut_iteration_and_mutation() {
        let mut sheet = ir_sheet("IR28", &[(1, 1, "a"), (1, 2, "b"), (1, 3, "c")], &[]);
        let mut wsheet = effective_bounds_mut(&mut sheet, 0, 0);
        let mut iter = IterRowMut::new(&mut wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        while let Some(mut rt) = iter.next() {
            ranges.push(rt.get_range().get_range());
            rt.get_sheet_mut().set_name("renamed");
        }
        assert_eq!(ranges, vec!["A1:B1".to_string(), "A2:B2".to_string(), "A3:B3".to_string()]);
        drop(iter);
        assert_eq!(sheet.get_name(), "renamed");
    }

    #[test]
    fn test_0104_iter_row_mut_multiline_kinds() {
        let mut sheet = ir_sheet("IR29", &[(1, 1, "-s"), (1, 2, "x"), (1, 3, "-n"), (1, 4, "t")], &[]);
        let mut wsheet = effective_bounds_mut(&mut sheet, 0, 0);
        let mut iter = IterRowMut::new(&mut wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut ranges = vec![];
        let mut kinds = vec![];
        while let Some(rt) = iter.next() {
            kinds.push(rt.get_type());
            ranges.push(rt.get_range().get_range());
        }
        assert_eq!(ranges, vec!["A1:B2".to_string(), "A3:B5".to_string()]);
        assert_eq!(kinds, vec![range_types::IterRowNextKind::Multiline, range_types::IterRowNextKind::Multiline]);
        assert_eq!(iter.current_row, 5);
    }

    #[test]
    fn test_0105_iter_row_mut_empty_sheet() {
        let mut sheet = ir_sheet("IR30", &[], &[]);
        let mut wsheet = effective_bounds_mut(&mut sheet, 0, 0);
        let mut iter = IterRowMut::new(&mut wsheet, 1, false, "^-", Offsets::default()).unwrap();
        let mut count = 0;
        while let Some(_rt) = iter.next() {
            count += 1;
        }
        assert_eq!(count, 0);
        assert_eq!(iter.current_row, 1);
    }
}