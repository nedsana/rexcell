use umya_spreadsheet::{Range, Worksheet};
use super::range_ops;
use std::vec;

fn range_bounds(range: &Range) -> (u32, u32, u32, u32, u32, u32) {
    let brow = *range.get_coordinate_start_row().unwrap().get_num();
    let erow = *range.get_coordinate_end_row().unwrap().get_num();
    let bcol = *range.get_coordinate_start_col().unwrap().get_num();
    let ecol = *range.get_coordinate_end_col().unwrap().get_num();
    let rows = erow - brow;
    let cols = ecol - bcol;
    (brow, erow, bcol, ecol, rows, cols)
}

fn compare_cell_impl<T>(
    this: &T,           col_this: u32,  row_this: u32,
    other: &dyn IRange, col_other: u32, row_other: u32,
    strict: bool,   label: &str,
) -> bool
where
    T: IRange,
{
    let sheet_this = this.get_sheet();
    let sheet_other = other.get_sheet();

    let mut r = false;
    let cell_a_coord = (col_this, row_this);
    let cell_b_coord = (col_other, row_other);
    let val_a = sheet_this.get_cell_value(cell_a_coord).get_value();
    let val_b = sheet_other.get_cell_value(cell_b_coord).get_value();

    let cell_a_obj = sheet_this.get_cell(cell_a_coord);
    let cell_b_obj = sheet_other.get_cell(cell_b_coord);
    let rich_a = cell_a_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());
    let rich_b = cell_b_obj.and_then(|c| c.get_cell_value().get_raw_value().get_rich_text());

    if rich_a != rich_b && strict 
    {
        println!("[{}::compare_cell] Rich text mismatch: {}:{} and {}:{}", label, 
            range_ops::coords_to_str(col_this, row_this), val_a,
            range_ops::coords_to_str(col_other, row_other), val_b);
        return r;
    }

    if range_ops::cmp_strs(&val_a, &val_b) 
    {
        println!("[{}::compare_cell] {}:{} equals {}:{}", label,
            range_ops::coords_to_str(col_this, row_this), val_a,
            range_ops::coords_to_str(col_other, row_other), val_b);
        r = true;
    } 
    else 
    {
        println!("[{}::compare_cell] {}:{} differs {}:{}", label,
            range_ops::coords_to_str(col_this, row_this), val_a,
            range_ops::coords_to_str(col_other, row_other), val_b);
        r = false;
    }
    r
}

fn compare_simple_range_impl<T>(
    this: &T,
    other: &dyn IRange,
    strict: bool,
    o_use_rows: Option<vec::Vec<u32>>,
    o_use_cols: Option<vec::Vec<u32>>,
    label: &str,
) -> bool
where
    T: IRange,
{
    let range_a = this.get_range();
    let range_b = other.get_range();

    let (brow_a, erow_a, bcol_a, _ecol_a, rows_a, cols_a) = range_bounds(range_a);
    let (brow_b, erow_b, bcol_b, _ecol_b, rows_b, cols_b) = range_bounds(range_b);

    if strict && (rows_a != rows_b || cols_a != cols_b) {
        println!("[{}::compare_range] Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", 
            label, range_ops::range_to_string(range_a), rows_a, range_ops::range_to_string(range_b), rows_b);
        return false;
    }

    let cols_a_offsets: Vec<u32> = (0..=cols_a).collect();
    let cols_b_offsets: Vec<u32> = (0..=cols_b).collect();
    let allowed_rows: Vec<u32> = o_use_rows.unwrap_or_default();
    let allowed_cols: Vec<u32> = o_use_cols.unwrap_or_default();

    let _str_allowed_rows = allowed_rows.iter().map(|r| r.to_string()).collect::<Vec<String>>().join(",");
    let _str_allowed_cols = allowed_cols.iter().map(|c| c.to_string()).collect::<Vec<String>>().join(",");

    let mut row_match = 0;
    let mut col_match: u32;

    for row_num_a in brow_a..=erow_a {
        if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_a) {
            continue;
        }

        for row_num_b in brow_b..=erow_b {
            if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b) {
                continue;
            }

            col_match = 0;
            for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) {
                let col_num_a = bcol_a + col_a_offset;
                let col_num_b = bcol_b + col_b_offset;

                if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) {
                    col_match += 1;
                    continue;
                }

                if this.compare_cell(col_num_a, row_num_a, other, col_num_b, row_num_b, strict) {
                    col_match += 1;
                }
            }

            if col_match == cols_a {
                row_match += 1;
            }
        }
    }

    if !strict && row_match != rows_a {
        return false;
    }
    true
}

fn compare_merged_range_impl<T>(
    this: &T,
    other: &dyn IRange,
    strict: bool,
    o_use_rows: Option<vec::Vec<u32>>,
    o_use_cols: Option<vec::Vec<u32>>,
    label: &str,
) -> bool
where
    T: IRange,
{
    let range_a = this.get_range();
    let range_b = other.get_range();

    let (brow_a, erow_a, bcol_a, _ecol_a, rows_a, cols_a) = range_bounds(range_a);
    let (brow_b, erow_b, bcol_b, _ecol_b, rows_b, cols_b) = range_bounds(range_b);

    if strict && (rows_a != rows_b || cols_a != cols_b) {
        println!("[{}::compare_range] Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", 
            label, range_ops::range_to_string(range_a), rows_a, range_ops::range_to_string(range_b), rows_b);
        return false;
    }

    let cols_a_offsets: Vec<u32> = (0..=cols_a).collect();
    let cols_b_offsets: Vec<u32> = (0..=cols_b).collect();
    let allowed_rows: Vec<u32> = o_use_rows.unwrap_or_default();
    let allowed_cols: Vec<u32> = o_use_cols.unwrap_or_default();

    let _str_allowed_rows = allowed_rows.iter().map(|r| r.to_string()).collect::<Vec<String>>().join(",");
    let _str_allowed_cols = allowed_cols.iter().map(|c| c.to_string()).collect::<Vec<String>>().join(",");

    let mut row_match = 0;
    let mut col_match: u32;

    for row_num_a in brow_a..=erow_a {
        if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_a) {
            continue;
        }

        for row_num_b in brow_b..=erow_b {
            if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b) {
                continue;
            }

            col_match = 0;
            for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) {
                let col_num_a = bcol_a + col_a_offset;
                let col_num_b = bcol_b + col_b_offset;

                if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) {
                    col_match += 1;
                    continue;
                }

                if this.compare_cell(col_num_a, row_num_a, other, col_num_b, row_num_b, strict) {
                    col_match += 1;
                }
            }

            if col_match == cols_a {
                row_match += 1;
            }
        }
    }

    if !strict && row_match != rows_a {
        return false;
    }
    true
}

fn compare_multiline_range_impl<T>(
    this: &T,
    other: &dyn IRange,
    strict: bool,
    o_use_rows: Option<vec::Vec<u32>>,
    o_use_cols: Option<vec::Vec<u32>>,
    label: &str,
) -> bool
where
    T: IRange,
{
    let range_this = this.get_range();
    let range_other = other.get_range();

    let (brow_this, erow_this, bcol_this, _ecol_this, rows_this, cols_this) = range_bounds(range_this);

    let (brow_other, erow_other, bcol_other, _ecol_other, rows_other, cols_other) = range_bounds(range_other);

    let rows_cnt_this = rows_this + 1;
    let cols_cnt_this = cols_this + 1;
    let rows_cnt_other = rows_other + 1;
    let cols_cnt_other = cols_other + 1;

    println!( "[{}::compare_range] Range A:[{} Rows:{} Cols:{}] vs Range B:[{} Rows:{} Cols:{}]", label,
        range_ops::range_to_string(range_this), rows_cnt_this, cols_cnt_this,
        range_ops::range_to_string(range_other), rows_cnt_other, cols_cnt_other);

    if strict && (rows_cnt_this != rows_cnt_other || cols_cnt_this != cols_cnt_other) 
    {
        println!("[{}::compare_range] Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", label,
            range_ops::range_to_string(range_this), rows_cnt_this,
            range_ops::range_to_string(range_other), rows_cnt_other);
        return false;
    }

    let cols_cnt_this_offsets: Vec<u32> = (0..=(cols_cnt_this - 1)).collect();
    let cols_cnt_other_offsets: Vec<u32> = (0..=(cols_cnt_other - 1)).collect();
    let allowed_rows: Vec<u32> = o_use_rows.unwrap_or_default();
    let allowed_cols: Vec<u32> = o_use_cols.unwrap_or_default();

    let _str_allowed_rows = allowed_rows.iter().map(|r| range_ops::index_to_column(*r)).collect::<Vec<String>>().join(",");
    let _str_allowed_cols = allowed_cols.iter().map(|c| range_ops::index_to_column(*c)).collect::<Vec<String>>().join(",");

    let mut row_match = 0;
    for row_num_a in brow_this..=erow_this 
    {
        if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_a) 
        {
            println!("[{}::compare_range] A{} is not in the allowed row list: {}!", label, row_num_a, _str_allowed_rows);
            continue;
        }

        for row_num_b in brow_other..=erow_other 
        {
            if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b) 
            {
                println!("[{}::compare_range] B{} is not in the allowed row list: {}!", label, row_num_b, _str_allowed_rows);
                continue;
            }

            let mut col_match = 0;
            for (col_a_offset, col_b_offset) in cols_cnt_this_offsets.iter().zip(cols_cnt_other_offsets.iter()) 
            {
                let col_num_a = bcol_this + col_a_offset;
                let col_num_b = bcol_other + col_b_offset;

                if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) 
                {
                    println!("[{}::compare_range] {}{} or {}{} is not in the allowed column list: {}!", label,
                        range_ops::index_to_column(col_num_a), row_num_a,
                        range_ops::index_to_column(col_num_b), row_num_b, _str_allowed_cols);
                    col_match += 1;
                    continue;
                }

                if this.compare_cell(col_num_a, row_num_a, other, col_num_b, row_num_b, strict) 
                {
                    col_match += 1;
                }
            }

            println!("[{}::compare_range] Matching columns: {}/{} ! Returning {}", label, col_match, cols_cnt_this, col_match == cols_cnt_this);

            if col_match == cols_cnt_this 
            {
                row_match += 1;
            }
        }
    }

    println!("[{}::compare_range] Matching rows: {}/{} ! Returning {}", label, row_match, rows_cnt_this, row_match == rows_cnt_this);

    if !strict && row_match != rows_cnt_this 
    {
        return false;
    }
    true
}
// ==========================================
// TRAIT
// ==========================================

pub trait IRange {
    fn get_range(&self) -> &Range;
    fn get_sheet(&self) -> &Worksheet;
    fn contains(&self, other: &Range) -> bool;
    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool;
    fn compare_cell(&self, col_a: u32, row_a: u32, other: &dyn IRange, col_b: u32, row_b: u32, strict: bool) -> bool;
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
        false
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_simple_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeBasic")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeBasic")
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
        false
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_simple_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeBasicMut")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeBasicMut")
    }    
}

impl<'a> IRangeMut for RangeBasicMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeBasic<'a> {
    fn eq(&self, _other: &Self) -> bool {
        false
    }
}

impl<'a> PartialEq for RangeBasicMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
        false
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
        false
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_merged_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeMergedCells")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeMergedCells")
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
        false
    }
    
    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_merged_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeMergedCellsMut")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeMergedCellsMut")
    }
}

impl<'a> IRangeMut for RangeMergedCellsMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeMergedCells<'a> {
    fn eq(&self, _other: &Self) -> bool {
        false
    }
}

impl<'a> PartialEq for RangeMergedCellsMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
        false
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
        false
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_multiline_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeMultiline")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeMultiline")
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
        false
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool 
    {
        compare_multiline_range_impl(self, other, strict, o_use_rows, o_use_cols, "RangeMultilineMut")
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool 
    {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict, "RangeMultilineMut")
    }
}

impl<'a> IRangeMut for RangeMultilineMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<'a> PartialEq for RangeMultiline<'a> {
    fn eq(&self, _other: &Self) -> bool {
	    false
    }
}

impl<'a> PartialEq for RangeMultilineMut<'a> {
    fn eq(&self, _other: &Self) -> bool {
	    false
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

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool {
        match self {
            RangeType::Basic(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
            RangeType::Merged(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
            RangeType::Multiline(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
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

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool {
        match self {
            RangeTypeMut::Basic(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
            RangeTypeMut::Merged(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
            RangeTypeMut::Multiline(r) => r.compare_cell(col_this, row_this, other, col_other, row_other, strict),
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
