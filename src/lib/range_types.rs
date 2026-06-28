use umya_spreadsheet::{Range, Worksheet};
use super::range_ops;
use std::vec;

// ==========================================
// TRAIT
// ==========================================

// pub fn comapre_ranges(
//     sheet_a: &Worksheet, range_a: &Range,
//     sheet_b: &Worksheet, range_b: &Range,
//     strict: bool,
//     o_allowed_rows: Option<vec::Vec<u32>>, 
//     o_allowed_cols: Option<vec::Vec<u32>>,
// ) -> bool 

pub trait IRange {
    fn get_range(&self) -> &Range;
    fn get_sheet(&self) -> &Worksheet;
    fn contains(&self, other: &Range) -> bool;
    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool;
    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool;
}

pub trait IRangeMut: IRange {
    fn get_sheet_mut(&mut self) -> &mut Worksheet;
}

// ==========================================
// STRUCTS
// ==========================================

pub struct RangeBasic<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

pub struct RangeBasicMut<'a> {
    pub range: Range,
    pub sheet: &'a mut Worksheet,
}

pub struct RangeMergedCells<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

pub struct RangeMergedCellsMut<'a> {
    pub range: Range,
    pub sheet: &'a mut Worksheet,
}

pub struct RangeMultiline<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

pub struct RangeMultilineMut<'a> {
    pub range: Range,
    pub sheet: &'a mut Worksheet,
}

// ==========================================
// STRUCT IMPLEMENTATION
// ==========================================

// ----------------------  RangeBasic ----------------------

impl<'a> IRange for RangeBasic<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!() 
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        let range_a = self.get_range();
        let range_b = other.get_range();

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

        //If the legths are different, the ranges cannot be the same
        if strict && (rows_a != rows_b || cols_a != cols_b) 
        {
            println!("[RangeBasic::compare] Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", 
                    range_ops::range_to_string(range_a), rows_a, range_ops::range_to_string(range_b), rows_b);
            return false;
        }

        let cols_a_offsets: Vec<u32> = (0..=cols_a).collect();
        let cols_b_offsets: Vec<u32> = (0..=cols_b).collect();    
        let allowed_rows: Vec<u32> = o_use_rows.unwrap_or_default();
        let allowed_cols: Vec<u32> = o_use_cols.unwrap_or_default();

        let _str_allowed_rows = allowed_rows.iter().map(|r| r.to_string()).collect::<Vec<String>>().join(",");
        let _str_allowed_cols = allowed_cols.iter().map(|c| c.to_string()).collect::<Vec<String>>().join(",");

        let mut row_match = 0;
        let mut col_match = 0;

        for row_num_a in brow_a..=erow_a 
        {
            if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_a)
            {
                println!("[RangeBasic::compare], Row A:{} is not in the allowed list {}!", row_num_a, _str_allowed_rows);
                continue; // skip this row if it's not in the allowed_rows list
            }

            for row_num_b in brow_b..=erow_b 
            {
                if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b)
                {
                    println!("[RangeBasic::compare], Row B:{} is not in the allowed list {}!", row_num_b, _str_allowed_rows);
                    continue; // skip this row if it's not in the allowed_rows list
                }

                col_match = 0;
                for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) 
                {
                    let col_num_a = bcol_a + col_a_offset;
                    let col_num_b = bcol_b + col_b_offset;

                    if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) 
                    {
                        println!("[RangeBasic::compare], Column A:{} or Column B:{} is not in the allowed list {}!", col_num_a, col_num_b, _str_allowed_cols);
                        col_match += 1;
                        continue; // skip this column if it's not in the allowed_cols list
                    }

                    if self.compare_cell(other, col_num_a, row_num_a, col_num_b, row_num_b, strict)
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
            return false;
        }
        true
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool 
    {
        let sheet_a = self.get_sheet();
        let sheet_b = other.get_sheet();

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
            println!("[comapre_cell] Rich text mismatch: {}:{} and {}:{}", range_ops::coords_to_str(col_a, row_a), val_a, range_ops::coords_to_str(col_b, row_b), val_b);
            return r;
        }

        // If there is any mismatch, immediately stop and return false
        if range_ops::cmp_strs(&val_a, &val_b) 
        {
            println!("[comapre_cell] {}:{} equals {}:{}", range_ops::coords_to_str(col_a, row_a), val_a, range_ops::coords_to_str(col_b, row_b), val_b);
            r = true;
        }
        else 
        {
            println!("[comapre_cell] {}:{} differs {}:{}", range_ops::coords_to_str(col_a, row_a), val_a, range_ops::coords_to_str(col_b, row_b), val_b); 
            r = false;
        }
        r
    }
}

impl<'a> IRange for RangeBasicMut<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!() 
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        todo!()
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        todo!()
    }    
}

impl<'a> IRangeMut for RangeBasicMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeBasic<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

impl<'a> PartialEq for RangeBasicMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

// ----------------------  RangeMergedCells ----------------------

impl<'a> IRange for RangeMergedCells<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!()
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        todo!()
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        todo!()
    }
}

impl<'a> IRange for RangeMergedCellsMut<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!()
    }
    
    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        todo!()
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        todo!()
    }
}

impl<'a> IRangeMut for RangeMergedCellsMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeMergedCells<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

impl<'a> PartialEq for RangeMergedCellsMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

// ----------------------  RangeMultiline ----------------------

impl<'a> IRange for RangeMultiline<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!()
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        todo!()
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        todo!()
    }
}

impl<'a> IRange for RangeMultilineMut<'a> {
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, _other: &Range) -> bool {
        todo!()
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        todo!()
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        todo!()
    }
}

impl<'a> IRangeMut for RangeMultilineMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeMultiline<'a> {
    fn eq(&self, _other: &Self) -> bool {
	    todo!()
    }
}

impl<'a> PartialEq for RangeMultilineMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
	    todo!()
    }
}

// ==========================================
// ENUM, UNITING THE TYPES
// ==========================================

pub enum RangeType<'a> {
    Basic(RangeBasic<'a>),
    Merged(RangeMergedCells<'a>),
    Multiline(RangeMultiline<'a>),
}

pub enum RangeTypeMut<'a> {
    Basic(RangeBasicMut<'a>),
    Merged(RangeMergedCellsMut<'a>),
    Multiline(RangeMultilineMut<'a>),
}

impl<'a> IRange for RangeType<'a> {
    fn get_range(&self) -> &Range {
        match self {
            RangeType::Basic(r) => r.get_range(),
            RangeType::Merged(r) => r.get_range(),
            RangeType::Multiline(r) => r.get_range(),
        }
    }

    fn get_sheet(&self) -> &Worksheet {
        match self {
            RangeType::Basic(r) => r.get_sheet(),
            RangeType::Merged(r) => r.get_sheet(),
            RangeType::Multiline(r) => r.get_sheet(),
        }
    }

    fn contains(&self, other: &Range) -> bool {
        match self {
            RangeType::Basic(r) => r.contains(other),
            RangeType::Merged(r) => r.contains(other),
            RangeType::Multiline(r) => r.contains(other),
        }
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        match self {
            RangeType::Basic(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
            RangeType::Merged(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
            RangeType::Multiline(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
        }
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        match self {
            RangeType::Basic(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
            RangeType::Merged(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
            RangeType::Multiline(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
        }
    }
}

impl<'a> IRange for RangeTypeMut<'a> {
    fn get_range(&self) -> &Range {
        match self {
            RangeTypeMut::Basic(r) => r.get_range(),
            RangeTypeMut::Merged(r) => r.get_range(),
            RangeTypeMut::Multiline(r) => r.get_range(),
        }
    }

    fn get_sheet(&self) -> &Worksheet {
        match self {
            RangeTypeMut::Basic(r) => r.get_sheet(),
            RangeTypeMut::Merged(r) => r.get_sheet(),
            RangeTypeMut::Multiline(r) => r.get_sheet(),
        }
    }

    fn contains(&self, other: &Range) -> bool {
        match self {
            RangeTypeMut::Basic(r) => r.contains(other),
            RangeTypeMut::Merged(r) => r.contains(other),
            RangeTypeMut::Multiline(r) => r.contains(other),
        }
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        match self {
            RangeTypeMut::Basic(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
            RangeTypeMut::Merged(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
            RangeTypeMut::Multiline(r) => r.compare_range(other, strict, o_use_rows, o_use_cols),
        }
    }

    fn compare_cell(&self, other: &dyn IRange, col_a: u32, row_a: u32, col_b: u32, row_b: u32, strict: bool) -> bool {
        match self {
            RangeTypeMut::Basic(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
            RangeTypeMut::Merged(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
            RangeTypeMut::Multiline(r) => r.compare_cell(other, col_a, row_a, col_b, row_b, strict),
        }
    }
}

impl<'a> IRangeMut for RangeTypeMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        match self {
            RangeTypeMut::Basic(r)     => r.get_sheet_mut(),
            RangeTypeMut::Merged(r)  => r.get_sheet_mut(),
            RangeTypeMut::Multiline(r) => r.get_sheet_mut(),
        }
    }
}
