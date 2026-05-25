use umya_spreadsheet::{Worksheet, Range};

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

    let mut coord_names = Vec::new();
    let mut cell_values = Vec::new();

    for r in rbeg..=rend 
    {
        for c in cbeg..=cend 
        {
            let coord_str = umya_spreadsheet::helper::coordinate::coordinate_from_index(&c, &r);        
            let _truncated = match truncate_len 
            {
                Some(tlen) => coord_names.push(truncate_str_with_dots(&coord_str, tlen as usize)),
                None => coord_names.push(coord_str),
            };

            let cell_value = sheet.get_cell_value((c, r)).get_value().to_string();
            let _truncated = match truncate_len 
            {
                Some(tlen) => cell_values.push(truncate_str_with_dots(&cell_value, tlen as usize)),
                None => cell_values.push(cell_value) ,
            };
        }

        println!("{}", coord_names.join("\t\t\t"));
        println!("{}", cell_values.join("\t\t\t"));
        coord_names.clear();
        cell_values.clear();
    }
}

pub struct IterRow<'a> 
{
    sheet: &'a Worksheet,
    sheet_merged_cells: &'a [Range],
    current_row: u32,
    max_row: u32,
    max_col: u32,
}

impl<'a> IterRow<'a> 
{
    // Конструктор за лесно създаване
    pub fn new(sheet: &'a Worksheet, mrow: u32, mcol: u32) -> Self {
        Self {
            sheet,
            sheet_merged_cells: sheet.get_merge_cells(),
            current_row: 1,
            max_row: mrow,
            max_col: mcol,
        }
    }
}

// Имплементиране на стандартния Iterator трейт на Rust
impl<'a> Iterator for IterRow<'a> 
{
    type Item = Range; // Итераторът връща готов Range обект
    fn next(&mut self) -> Option<Self::Item> 
    {
        let mut ret: Option<Self::Item> = None;

        if self.max_row > self.current_row
        {
            if let Some(merged_cells) = self.sheet_merged_cells.iter().find(|range| { is_row_in_range(self.current_row, range) }) 
            {
                //handle rows with merged cells
                let merged_cells_value_inst = self.sheet.get_cell_value((
                    merged_cells.get_coordinate_start_col().unwrap().get_num(), 
                    merged_cells.get_coordinate_start_row().unwrap().get_num()));

                let _merged_cells_value = merged_cells_value_inst.get_value().clone();
                let merged_cells_value_type = merged_cells_value_inst.get_data_type();

                let is_single_col  = *merged_cells.get_coordinate_start_col().unwrap().get_num() == *merged_cells.get_coordinate_end_col().unwrap().get_num();
                let merged_rows_cnt = *merged_cells.get_coordinate_end_row().unwrap().get_num() - *merged_cells.get_coordinate_start_row().unwrap().get_num();
                let is_two_rows    = merged_rows_cnt == 1;

                self.current_row += merged_rows_cnt + 1; // Move to the next row for the next iteration, skipping the merged rows

                if merged_cells_value_type == "n" && is_single_col && is_two_rows
                {
                    let cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, 
                                *merged_cells.get_coordinate_end_row().unwrap().get_num());

                    // println!("[{}] Merged cell range: {}{}:{}{} found for row {}. Value:{} Type:{}!", 
                    //     self.sheet.get_name(), 
                    //     index_to_column(*merged_cells.get_coordinate_start_col().unwrap().get_num()), 
                    //     merged_cells.get_coordinate_start_row().unwrap().get_num(), 
                    //     index_to_column(*merged_cells.get_coordinate_end_col().unwrap().get_num()), 
                    //     merged_cells.get_coordinate_end_row().unwrap().get_num(), self.current_row, _merged_cells_value, merged_cells_value_type);

                    ret = Some(cells_range);
                }
            } 
            else if let Some(src_cell) = self.sheet.get_cell((1, self.current_row)) 
            {
                //handle rows without merged cells. Ignore rows with values in Col:A different from numeric or symbol '-'
                let _first_cell_value = src_cell.get_value().clone();
                let first_cell_data_type = src_cell.get_data_type().to_string();

                let mut cells_range = make_range_from_indexes(1, self.current_row, 1 + self.max_col, self.current_row);

                self.current_row += 1; // Move to the next row for the next iteration

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
                    
                    ret = Some(cells_range);
                }
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