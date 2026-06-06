use std::vec;

use umya_spreadsheet::{Worksheet, Range};

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

pub fn is_range_in_range(sub_range: &Range, main_range: &Range) -> bool 
{
    // println!("sub-range:{} range:{}", range_to_string(sub_range), range_to_string(main_range));

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

pub fn do_ranges_overlap(range_a: &Range, range_b: &Range) -> bool 
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

pub fn comapre_ranges(
    sheet_a: &Worksheet, range_a: &Range,
    sheet_b: &Worksheet, range_b: &Range
) -> bool 
{
    //Get the range numeric boundaries for range_a
    let a_start_row = *range_a.get_coordinate_start_row().unwrap().get_num();
    let a_end_row   = *range_a.get_coordinate_end_row().unwrap().get_num();
    let a_start_col = *range_a.get_coordinate_start_col().unwrap().get_num();
    let a_end_col   = *range_a.get_coordinate_end_col().unwrap().get_num();

    //Get the range numeric boundaries for range_b
    let b_start_row = *range_b.get_coordinate_start_row().unwrap().get_num();
    let b_end_row   = *range_b.get_coordinate_end_row().unwrap().get_num();
    let b_start_col = *range_b.get_coordinate_start_col().unwrap().get_num();
    let b_end_col   = *range_b.get_coordinate_end_col().unwrap().get_num();

    //Get the legths of the ranges (number of rows and columns)
    let a_rows = a_end_row - a_start_row;
    let a_cols = a_end_col - a_start_col;
    let b_rows = b_end_row - b_start_row;
    let b_cols = b_end_col - b_start_col;

    //If the legths are different, the ranges cannot be the same
    if a_rows != b_rows || a_cols != b_cols 
    {
        return false; 
    }

    //Loop using relative offsets to compare the cells in the two ranges
    for row_offset in 0..=a_rows 
    {
        for col_offset in 0..=a_cols 
        {
            //Calculate the actual coordinates for sheet A and sheet B
            let cell_a_coord = (a_start_col + col_offset, a_start_row + row_offset);
            let cell_b_coord = (b_start_col + col_offset, b_start_row + row_offset);

            let cell_a_obj = sheet_a.get_cell(cell_a_coord);
            let cell_b_obj = sheet_b.get_cell(cell_b_coord);

            let rich_a = cell_a_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());
            let rich_b = cell_b_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());

            if rich_a != rich_b 
            {
                return false; // Different rich text content means the cells are not the same
            }

            // Get the text values of the two cells
            let val_a = sheet_a.get_cell_value(cell_a_coord).get_value();
            let val_b = sheet_b.get_cell_value(cell_b_coord).get_value();

            // If there is any mismatch, immediately stop and return false
            if cmp_strs(&val_a, &val_b) 
            {
                return false;
            }
        }
    }

    // If all cells have matched in order, the content is the same
    true
}

//return the cells in range_a and range_b, which are in the same position and have numeric values, accumulated (summed up)
pub fn accumulate_ranges(
    sheet_a: &Worksheet, range_a: &Range,
    sheet_b: & mut Worksheet, range_b: &Range,
    o_use_rows: Option<vec::Vec<u32>>, 
    o_use_cols: Option<vec::Vec<u32>>,
) -> bool
{
    let mut accumulated: bool = false;

    //Get the range numeric boundaries for range_a
    let brow_a = *range_a.get_coordinate_start_row().unwrap().get_num();
    let erow_a = *range_a.get_coordinate_end_row().unwrap().get_num();
    let bcol_a = *range_a.get_coordinate_start_col().unwrap().get_num();
    let ecol_a = *range_a.get_coordinate_end_col().unwrap().get_num();

    //Get the range numeric boundaries for range_b
    let brow_b = *range_b.get_coordinate_start_row().unwrap().get_num();
    let erow_b = *range_b.get_coordinate_end_row().unwrap().get_num();
    let bcol_b = *range_b.get_coordinate_start_col().unwrap().get_num();
    let ecol_b = *range_b.get_coordinate_end_col().unwrap().get_num();

    //Get the legths of the ranges (number of rows and columns)
    let rows_a = erow_a - brow_a;
    let cols_a = ecol_a - bcol_a;
    let rows_b = erow_b - brow_b;
    let cols_b = ecol_b - bcol_b;

    //If the legths are different, the ranges cannot proceed with accumulation
    if rows_a != rows_b || cols_a != cols_b 
    {
        return accumulated; 
    }

    let rows_offsets: Vec<u32> = (0..=rows_a).collect();
    let cols_offsets: Vec<u32> = (0..=cols_a).collect();
    let use_rows: Vec<u32> = o_use_rows.unwrap_or_default();
    let use_cols: Vec<u32> = o_use_cols.unwrap_or_default();

    //Loop using relative offsets to compare the cells in the two ranges
    for row_offset in &rows_offsets 
    {
        if use_rows.len() > 0 && !use_rows.contains(row_offset) 
        {
            continue; // skip this row if it's not in the use_rows list
        }

        for col_offset in &cols_offsets 
        {
            if use_cols.len() > 0 && !use_cols.contains(col_offset) 
            {
                continue; // skip this column if it's not in the use_cols list
            }

            //Calculate the actual coordinates for sheet A and sheet B
            let coord_a = (bcol_a + col_offset, brow_a + row_offset);
            let coord_b = (bcol_b + col_offset, brow_b + row_offset);

            let cell_a = sheet_a.get_cell(coord_a);
            let cell_b = sheet_b.get_cell(coord_b);

            if !cell_a.is_none() && !cell_b.is_none() 
            {
                if cell_a.as_ref().unwrap().get_data_type() == "n" && cell_b.as_ref().unwrap().get_data_type() == "n" 
                {
                    let val_a = cell_a.as_ref().unwrap().get_value().parse::<f64>().unwrap_or(0.0);
                    let val_b = cell_b.as_ref().unwrap().get_value().parse::<f64>().unwrap_or(0.0);
                    let sum: f64 = val_a + val_b;

                    let q_cell_dst = sheet_b.get_cell_mut(coord_b);
                    q_cell_dst.set_value_number(sum);

                    accumulated = true;
                }
            }
        }
    }

    accumulated
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
    let rbeg = *range.get_coordinate_start_row().unwrap().get_num();
    let rend = *range.get_coordinate_end_row().unwrap().get_num();
    let cbeg = *range.get_coordinate_start_col().unwrap().get_num();
    let cend = *range.get_coordinate_end_col().unwrap().get_num();

    format!("{}:{}", umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
                     umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend)  // Returns "C10"
    )
}

pub fn print_range_cells_0(sheet: &Worksheet, range: &Range) 
{
    let rbeg = *range.get_coordinate_start_row().unwrap().get_num();
    let rend = *range.get_coordinate_end_row().unwrap().get_num();
    let cbeg = *range.get_coordinate_start_col().unwrap().get_num();
    let cend = *range.get_coordinate_end_col().unwrap().get_num();

    println!("Sheet {} range ({}:{}) ---", sheet.get_name(), 
        umya_spreadsheet::helper::coordinate::coordinate_from_index(&cbeg, &rbeg), // Returns "A1"
        umya_spreadsheet::helper::coordinate::coordinate_from_index(&cend, &rend)  // Returns "C10"
    );

    for r in rbeg..=rend 
    {
        for c in cbeg..=cend 
        {
            let coord_str = umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r);
            let cell_value = sheet.get_cell_value((c, r)).get_value();
            println!("Клетка {}: {}", coord_str, cell_value);
        }
    }
}

pub fn print_range_cells_1(sheet: &Worksheet, range: &Range, truncate_len: Option<u32>) 
{
    let rbeg = *range.get_coordinate_start_row().unwrap().get_num();
    let rend = *range.get_coordinate_end_row().unwrap().get_num();
    let cbeg = *range.get_coordinate_start_col().unwrap().get_num();
    let cend = *range.get_coordinate_end_col().unwrap().get_num();

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

        println!("{}", coord_names.join(&sep));
        println!("{}", cell_values.join(&sep));
        coord_names.clear();
        cell_values.clear();
    }
}

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
    type Item = Range; // return Range object
    fn next(&mut self) -> Option<Self::Item> 
    {
        let mut ret: Option<Self::Item> = None;

        if self.max_row > self.current_row
        {
            let sheet_merged_cells = self.sheet.get_merge_cells();

            if let Some(merged_cells) = sheet_merged_cells.iter().find(|range| { is_row_in_range(self.current_row, range) }) 
            {
                //handle rows with merged cells - return all rows which are part of the merged cell
                let cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, 
                                *merged_cells.get_coordinate_end_row().unwrap().get_num());

                let range_rows = cells_range.get_coordinate_end_row().unwrap().get_num() - cells_range.get_coordinate_start_row().unwrap().get_num();
                
                self.current_row += range_rows + 1;

                ret = Some(cells_range);
            } 
            else if let Some(src_cell) = self.sheet.get_cell((1, self.current_row)) 
            {
                //Handle rows without merged cells. If we have in colA numeric, followed by symbol '-', return all rows starting with '-'
                let _first_cell_value = src_cell.get_value().clone();
                let first_cell_data_type = src_cell.get_data_type().to_string();

                let mut cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, self.current_row);

                if first_cell_data_type == "n"
                {
                    //check if the next row starts with numeric. I yes, process the current row. If not make range of all rows starting with '-'
                    let next_row = self.current_row + 1;
                    for nrow in next_row..=self.max_row 
                    {
                        if let Some(next_cell) = self.sheet.get_cell((1, nrow)) 
                        {
                            let _next_cell_value = next_cell.get_value().clone();
                            let next_cell_data_type = next_cell.get_data_type().to_string();

                            if next_cell_data_type == "n"
                            {
                                break;
                            }
                            else if next_cell_data_type == "s" && _next_cell_value == "-"
                            {
                                cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, nrow);
                            }
                        }
                    }
                }

                let range_rows = cells_range.get_coordinate_end_row().unwrap().get_num() - cells_range.get_coordinate_start_row().unwrap().get_num();
                
                self.current_row += range_rows + 1;

                ret = Some(cells_range);

            }
            else
            {
                println!("[{}] Processing unexpected row:{}!", self.sheet.get_name(), self.current_row);
                self.current_row += 1;
            }
        }

        ret
    }
}

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

impl<'a> Iterator for IterRowMut<'a> 
{
    type Item = Range; // return Range object
    fn next(&mut self) -> Option<Self::Item> 
    {
        let mut ret: Option<Self::Item> = None;

        if self.max_row > self.current_row
        {
            let sheet_merged_cells = self.sheet.get_merge_cells();

            if let Some(merged_cells) = sheet_merged_cells.iter().find(|range| { is_row_in_range(self.current_row, range) }) 
            {
                //handle rows with merged cells - return all rows which are part of the merged cell
                let cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, 
                                *merged_cells.get_coordinate_end_row().unwrap().get_num());

                let range_rows = cells_range.get_coordinate_end_row().unwrap().get_num() - cells_range.get_coordinate_start_row().unwrap().get_num();
                
                self.current_row += range_rows + 1;

                ret = Some(cells_range);
            } 
            else if let Some(src_cell) = self.sheet.get_cell((1, self.current_row)) 
            {
                //Handle rows without merged cells. If we have in colA numeric, followed by symbol '-', return all rows starting with '-'
                let _first_cell_value = src_cell.get_value().clone();
                let first_cell_data_type = src_cell.get_data_type().to_string();

                let mut cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, self.current_row);

                if first_cell_data_type == "n"
                {
                    //check if the next row starts with numeric. I yes, process the current row. If not make range of all rows starting with '-'
                    let next_row = self.current_row + 1;
                    for nrow in next_row..=self.max_row 
                    {
                        if let Some(next_cell) = self.sheet.get_cell((1, nrow)) 
                        {
                            let _next_cell_value = next_cell.get_value().clone();
                            let next_cell_data_type = next_cell.get_data_type().to_string();

                            if next_cell_data_type == "n"
                            {
                                break;
                            }
                            else if next_cell_data_type == "s" && _next_cell_value == "-"
                            {
                                cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, nrow);
                            }
                        }
                    }
                }

                let range_rows = cells_range.get_coordinate_end_row().unwrap().get_num() - cells_range.get_coordinate_start_row().unwrap().get_num();
                
                self.current_row += range_rows + 1;

                ret = Some(cells_range);

            }
            else
            {
                println!("[{}] Processing unexpected row:{}!", self.sheet.get_name(), self.current_row);
                self.current_row += 1;
            }
        }

        ret
    }
}