use umya_spreadsheet::{Range, Worksheet};
use super::range_ops;
use std::vec;
use std::ops::{Deref};//, DerefMut};
use std::marker::PhantomData;
use log::{info, error};

fn compare_cell_impl<T>(
    this: &T,           
    col_this: u32,  
    row_this: u32,
    other: &dyn IRange, 
    col_other: u32, 
    row_other: u32,
    strict: bool,
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
        error!("Rich text mismatch: {}:[{}:'{}'] and {}:[{}:'{}']", //label, 
            sheet_this.get_name(), range_ops::coords_to_str(col_this, row_this), val_a,
            sheet_other.get_name(), range_ops::coords_to_str(col_other, row_other), val_b);
        return r;
    }

    if range_ops::cmp_strs(&val_a, &val_b) 
    {
        info!("[{}:[{}:'{}'] EQUALS {}:[{}:'{}']", //label,
            sheet_this.get_name(), range_ops::coords_to_str(col_this, row_this), val_a,
            sheet_other.get_name(), range_ops::coords_to_str(col_other, row_other), val_b);
        r = true;
    } 
    else 
    {
        info!("{}:[{}:'{}'] DIFFERS {}:[{}:'{}']", //label,
            sheet_this.get_name(), range_ops::coords_to_str(col_this, row_this), val_a,
            sheet_other.get_name(), range_ops::coords_to_str(col_other, row_other), val_b);
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
) -> bool
where
    T: IRange,
{
    let range_a = this.get_range();
    let range_b = other.get_range();

    if let ( Some((brow_a, erow_a, bcol_a, _ecol_a, rows_a, cols_a)), 
             Some((brow_b, erow_b, bcol_b, _ecol_b, rows_b, cols_b)) ) = (range_ops::range_bounds(range_a), 
                                                                                                        range_ops::range_bounds(range_b)) 
    {
        if strict && (rows_a != rows_b || cols_a != cols_b) {
            error!("Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", //label, 
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
        return true;
    }
    false
}

fn compare_merged_range_impl<T>(
    this: &T,
    other: &dyn IRange,
    strict: bool,
    o_use_rows: Option<vec::Vec<u32>>,
    o_use_cols: Option<vec::Vec<u32>>,
) -> bool
where
    T: IRange,
{
    let range_a = this.get_range();
    let range_b = other.get_range();

    if let ( Some((brow_a, erow_a, bcol_a, _ecol_a, rows_a, cols_a)), 
             Some((brow_b, erow_b, bcol_b, _ecol_b, rows_b, cols_b)) ) = (range_ops::range_bounds(range_a), 
                                                                                                        range_ops::range_bounds(range_b)) 
    {
        if rows_a != rows_b || cols_a != cols_b
        {
            error!("Size missmatch! Range {}:[{}] != Range {}:[{}]", //label, 
                    this.get_sheet().get_name(), range_ops::range_to_string(range_a),
                    other.get_sheet().get_name(), range_ops::range_to_string(range_b));

            return false;
        }

        let cols_a_offsets: Vec<u32> = (0..=(cols_a-1)).collect();
        let cols_b_offsets: Vec<u32> = (0..=(cols_b-1)).collect();
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
                for (col_a_offset, col_b_offset) in cols_a_offsets.iter().zip(cols_b_offsets.iter()) 
                {
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

                // info!("cols_a:{}, col_match:{}]", //label, 
                //     cols_a, col_match);

                if col_match == cols_a {
                    row_match += 1;
                }
            }
        }

        // info!("rows_a:{}, row_match:{}]", //label, 
        //         rows_a, row_match);

        if row_match != rows_a 
        {
            // info!("Range {}:[{}, len:{}] DIFFERS FROM Range {}:[{}, len:{}]", //label, 
            //     this.get_sheet().get_name(), range_ops::range_to_string(range_a), rows_a, other.get_sheet().get_name(), range_ops::range_to_string(range_b), rows_b);

            return false;
        }

        // info!("Range {}:[{}, len:{}] EQUALS TO Range {}:[{}, len:{}]", //label, 
        //     this.get_sheet().get_name(), range_ops::range_to_string(range_a), rows_a, other.get_sheet().get_name(), range_ops::range_to_string(range_b), rows_b);

        return true
    }
    false
}

fn compare_multiline_range_impl<T>(
    this: &T,
    other: &dyn IRange,
    strict: bool,
    o_use_rows: Option<vec::Vec<u32>>,
    o_use_cols: Option<vec::Vec<u32>>,
) -> bool
where
    T: IRange,
{
    let range_this = this.get_range();
    let range_other = other.get_range();

    if let ( Some((brow_this, erow_this, bcol_this, _ecol_this, rows_this, cols_this)), 
             Some((brow_other, erow_other, bcol_other, _ecol_other, rows_other, cols_other)) ) = (range_ops::range_bounds(range_this), 
                                                                                                                                range_ops::range_bounds(range_other)) 
    {
        // let (brow_this, erow_this, bcol_this, _ecol_this, rows_this, cols_this) = range_ops::range_bounds(range_this);
        // let (brow_other, erow_other, bcol_other, _ecol_other, rows_other, cols_other) = range_ops::range_bounds(range_other);

        let rows_cnt_this = rows_this + 1;
        let cols_cnt_this = cols_this + 1;
        let rows_cnt_other = rows_other + 1;
        let cols_cnt_other = cols_other + 1;

        // info!( "Range A:[{} Rows:{} Cols:{}] vs Range B:[{} Rows:{} Cols:{}]", //label,
        //     range_ops::range_to_string(range_this), rows_cnt_this, cols_cnt_this,
        //     range_ops::range_to_string(range_other), rows_cnt_other, cols_cnt_other);

        if strict && (rows_cnt_this != rows_cnt_other || cols_cnt_this != cols_cnt_other) 
        {
            error!("Size missmatch! Range A:[{}, len:{}] != Range B:[{}, len:{}]", //label,
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
                info!("A{} is not in the allowed row list: {}!", //label, 
                    row_num_a, _str_allowed_rows);
                continue;
            }

            for row_num_b in brow_other..=erow_other 
            {
                if allowed_rows.len() > 0 && !allowed_rows.contains(&row_num_b) 
                {
                    info!("B{} is not in the allowed row list: {}!", //label, 
                        row_num_b, _str_allowed_rows);
                    continue;
                }

                let mut col_match = 0;
                for (col_a_offset, col_b_offset) in cols_cnt_this_offsets.iter().zip(cols_cnt_other_offsets.iter()) 
                {
                    let col_num_a = bcol_this + col_a_offset;
                    let col_num_b = bcol_other + col_b_offset;

                    if allowed_cols.len() > 0 && !allowed_cols.contains(&col_num_a) && !allowed_cols.contains(&col_num_b) 
                    {
                        // info!("{}{} or {}{} is not in the allowed column list: {}!", //label,
                        //     range_ops::index_to_column(col_num_a), row_num_a,
                        //     range_ops::index_to_column(col_num_b), row_num_b, _str_allowed_cols);
                        col_match += 1;
                        continue;
                    }

                    if this.compare_cell(col_num_a, row_num_a, other, col_num_b, row_num_b, strict) 
                    {
                        col_match += 1;
                    }
                }

                info!("Matching columns: {}/{} ! Returning {}", //label, 
                    col_match, cols_cnt_this, col_match == cols_cnt_this);

                if col_match == cols_cnt_this 
                {
                    row_match += 1;
                }
            }
        }

        info!("Matching rows: {}/{} ! Returning {}", //label, 
            row_match, rows_cnt_this, row_match == rows_cnt_this);

        if !strict && row_match != rows_cnt_this 
        {
            return false;
        }
        return true;
    }
    false
}

fn contains_impl<T>(
    this: &T,
    other: &dyn IRange,
    o_use_rows: Option<vec::Vec<u32>>, 
    o_use_cols: Option<vec::Vec<u32>>,
) -> bool
where
    T: IRange,
{
    let mut res = false;
    if range_ops::same_types(this, other)
    {
        let _cmp_rows: Vec<u32> = o_use_rows.unwrap_or_default(); //not used for now
        let cmp_cols: Vec<u32> = o_use_cols.unwrap_or_default();

        if let (Some((brow_t, erow_t, _, _, _, _)), 
                Some((brow_o, erow_o, _, _, rows_o, _)) ) = (range_ops::range_bounds(this.get_range()), 
                                                                            range_ops::range_bounds(other.get_range())) 
        {
            for cmp_col in cmp_cols.iter() 
            {
                let bcol_t = *cmp_col;
                let bcol_o = *cmp_col;

                //check if the first line of the range_in matches the first line of the current range in the sheets
                let hdr_t = this.get_sheet().get_cell_value((bcol_t, brow_t)).get_value();
                let hdr_o = other.get_sheet().get_cell_value((bcol_o, brow_o)).get_value();

                if range_ops::cmp_strs(&hdr_t, &hdr_o)
                {
                    let mut found_cnt = 1; //the header is already the same, so start form 1

                    //check if the rest of the items in 'other' are present in 'this'
                    for row_o in (brow_o+1)..=erow_o
                    {
                        let entry_o = other.get_sheet().get_cell_value((bcol_o, row_o)).get_value();

                        for row_t in (brow_t+1)..=erow_t
                        {
                            let entry_t = this.get_sheet().get_cell_value((bcol_t, row_t)).get_value();

                            info!("COMPARE \"{}:[{} {}{}='{}']\" to \"{}:[{} {}{}='{}']\"", //label,
                                other.get_sheet().get_name(), range_ops::range_to_string(other.get_range()), range_ops::index_to_column(bcol_o), row_o, entry_o,
                                this.get_sheet().get_name(),  range_ops::range_to_string(this.get_range()),  range_ops::index_to_column(bcol_t), row_t, entry_t);

                            if range_ops::cmp_strs(&entry_t, &entry_o) 
                            {
                                found_cnt += 1;
                                break;
                            }
                        }
                    }

                    info!("Found {}/{} entries!", //label, 
                        found_cnt, rows_o);
                    res = found_cnt == rows_o
                }
                else 
                {
                    info!("\"{}[{} {}{}='{}']\" DEFFERS FROM \"{}[{} {}{}='{}']\"", //label,
                            other.get_sheet().get_name(), range_ops::range_to_string(other.get_range()), range_ops::index_to_column(bcol_o), brow_o, hdr_o,
                            this.get_sheet().get_name(),  range_ops::range_to_string(this.get_range()),  range_ops::index_to_column(bcol_t), brow_t, hdr_t);
                    break;
                }
            }
        }
    }
    else
    {
        info!("Types mismatch: {}:[{} {}] vs {}:[{} {}]", //label, 
                other.get_sheet().get_name(), range_ops::range_to_string(other.get_range()), other.get_type_name(),
                other.get_sheet().get_name(), range_ops::range_to_string(this.get_range()),  this.get_type_name()); 
    }
    res
}

fn reset_iter_impl<S, TAG>(this: &mut RangeGeneric<S, TAG>) 
where
    RangeGeneric<S, TAG>: IRange,
{
    this.current_row = 1;

    if let Some(rbr) = this.get_range().get_coordinate_start_row() 
    {
        this.current_row = *rbr.get_num();
    }
}

// ==========================================
// TRAIT
// ==========================================
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum IterRowNextKind {
    Basic,
    Merged,
    Multiline,
}

pub trait IRange {
    fn get_range(&self) -> &Range;
    fn get_sheet(&self) -> &Worksheet;
    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool;
    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool;
    fn compare_cell(&self, col_a: u32, row_a: u32, other: &dyn IRange, col_b: u32, row_b: u32, strict: bool) -> bool;
    fn get_type(&self) -> IterRowNextKind;
    fn get_type_name(&self) -> String {
        match self.get_type() {
            IterRowNextKind::Basic => "Basic".to_string(),
            IterRowNextKind::Merged => "Merged".to_string(),
            IterRowNextKind::Multiline => "Multiline".to_string(),
        }
    }
    fn reset_iter(&mut self);
}

pub trait IRangeMut: IRange {
    fn get_sheet_mut(&mut self) -> &mut Worksheet;
}

//The statandart iterators, can no be mutable. So we need this specific iterator. NOTE: can't be used in for loops, but works with while!
pub trait LendingIterator
{
    type Item<'this> where Self: 'this;

    fn next(&mut self) -> Option<Self::Item<'_>>;
}

// ==========================================
// STRUCTS
// ==========================================

pub struct WorkSheet<S>
{
    pub sheet: S,
    pub max_row: u32,
    pub max_col: u32,
}

impl<S> WorkSheet<S> {
    pub fn new(ws: S, mr: u32, mc: u32) -> Self {
        Self {
            sheet:   ws,
            max_row: mr,
            max_col: mc,
        }
    }
}

pub trait IWorkSheet<T> { 
    fn sheet(&self) -> &T;
    fn max_row(&self) -> u32;
    fn max_col(&self) -> u32;
}

pub trait IWorkSheetMut<T>: IWorkSheet<T> { 
    fn sheet_mut(&mut self) -> &mut T;
}

impl<T> IWorkSheet<T> for WorkSheet<T> 
// where
//     T: Deref<Target = Worksheet>,
{ 
    fn sheet(&self)   -> &T  { &self.sheet }
    fn max_row(&self) -> u32 { self.max_row }
    fn max_col(&self) -> u32 { self.max_col }
}

impl<T> IWorkSheetMut<T> for WorkSheet<T> 
// where
//     T: DerefMut<Target = Worksheet>,
{ 
    fn sheet_mut(&mut self) -> &mut T  { &mut self.sheet }
}

// //Allows direct access to Worksheet interface. If no present, we will have to access ".sheet" member everywhere!
// impl<S> Deref for WorkSheet<S>
// where
//     S: Deref<Target = Worksheet>,
// {
//     type Target = Worksheet;

//     fn deref(&self) -> &Worksheet {
//         &*self.sheet
//     }
// }

// //Allows direct access to Worksheet interface. If no present, we will have to access ".sheet" member everywhere!
// impl<S> DerefMut for WorkSheet<S>
// where
//     S: DerefMut<Target = Worksheet>,
// {
//     fn deref_mut(&mut self) -> &mut Worksheet {
//         &mut *self.sheet
//     }
// }


pub struct TagRangeBasic;
pub struct TagRangeMergedCells;
pub struct TagRangeMultiline;

pub struct RangeGeneric<S, TAG> {
    pub range: Range,
    pub sheet: S,
    pub current_row: u32,
    _tag: PhantomData<TAG> //needed to create different types of this struct
}

impl<S, TAG> RangeGeneric<S, TAG> 
{
    pub fn new(p_range: Range, p_sheet: S) -> Self 
    {
        let mut br = 1;
        if let Some(rbr) = p_range.get_coordinate_start_row() 
        {
            br = *rbr.get_num();
        }

        Self { 
            range:          p_range, 
            sheet:          p_sheet, 
            current_row:    br,
            _tag:           PhantomData 
        }
    }
}

impl<S, TAG> Iterator for RangeGeneric<S, TAG>
where
    Self: IRange,
{
    type Item = Range;

    fn next(&mut self) -> Option<Self::Item>
    {
        let mut ret: Option<Self::Item> = None;

        if let Some((_, er, bc, ec, _, _)) = range_ops::range_bounds(self.get_range())
        {
            if self.current_row <= er
            {
                ret = Some(range_ops::make_range_from_indexes(bc, self.current_row, ec, self.current_row));

                self.current_row += 1;
            }
        }

        ret
    }
}

//This implementation is not used and most likely incorrect: the range-type-iterator should return 
//row-by-row the content of the 'range' member. But here we return the LendingIterator itself as mutable.
//The row-by-row access should be done externally which makes no sense.
//Keep the implementation for now, but it is not used!
impl<S, TAG> LendingIterator for RangeGeneric<S, TAG>
where
    Self: IRange,
{
    type Item<'this> = &'this mut Self where Self: 'this;

    fn next(&mut self) -> Option<Self::Item<'_>>
    {
        let mut ret: Option<Self::Item<'_>> = None;

        if let Some((_, er, _, _, _, _)) = range_ops::range_bounds(self.get_range())
        {
            if self.current_row <= er
            {
                // ret = Some(range_ops::make_range_from_indexes(bc, self.current_row, ec, self.current_row));

                self.current_row += 1;

                ret = Some(self);
            }
        }

        ret
    }
}

pub type RangeBasic<'a>             = RangeGeneric<&'a Worksheet, TagRangeBasic>;
pub type RangeBasicMut<'a>          = RangeGeneric<&'a mut Worksheet, TagRangeBasic>;
pub type RangeMergedCells<'a>       = RangeGeneric<&'a Worksheet, TagRangeMergedCells>;
pub type RangeMergedCellsMut<'a>    = RangeGeneric<&'a mut Worksheet, TagRangeMergedCells>;
pub type RangeMultiline<'a>         = RangeGeneric<&'a Worksheet, TagRangeMultiline>;
pub type RangeMultilineMut<'a>      = RangeGeneric<&'a mut Worksheet, TagRangeMultiline>;

// ==========================================
// STRUCT IMPLEMENTATION
// ==========================================

// ----------------------  RangeBasic ----------------------

// Implement IRange for RangeBasicGeneric, as long as 'S' can be converted to Worksheet
impl<S> IRange for RangeGeneric<S, TagRangeBasic> 
where
    S: Deref<Target = Worksheet>,
{
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        contains_impl(self, other, o_use_rows, o_use_cols)
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        compare_simple_range_impl(self, other, strict, o_use_rows, o_use_cols)
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict)
    }

    fn get_type(&self) -> IterRowNextKind { IterRowNextKind::Basic }

    fn reset_iter(&mut self)
    {
        reset_iter_impl(self);
    }
}

// Explicit mutable interface
impl<'a> IRangeMut for RangeBasicMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<S> PartialEq for RangeGeneric<S, TagRangeBasic> {
    fn eq(&self, _other: &Self) -> bool { false }
}

// ----------------------  RangeMergedCells ----------------------

// Implement IRange for RangeMergedCells, as long as 'S' can be converted to Worksheet
impl<S> IRange for RangeGeneric<S, TagRangeMergedCells> 
where
    S: Deref<Target = Worksheet>,
{
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        contains_impl(self, other, o_use_rows, o_use_cols)
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        compare_merged_range_impl(self, other, strict, o_use_rows, o_use_cols)
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict)
    }

    fn get_type(&self) -> IterRowNextKind { IterRowNextKind::Merged }

    fn reset_iter(&mut self)
    {
        reset_iter_impl(self);
    }
}

// Explicit mutable interface
impl<'a> IRangeMut for RangeMergedCellsMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<S> PartialEq for RangeGeneric<S, TagRangeMergedCells> {
    fn eq(&self, _other: &Self) -> bool { false }
}

// ----------------------  RangeMultiline ----------------------

// Implement IRange for RangeMergedCells, as long as 'S' can be converted to Worksheet
impl<S> IRange for RangeGeneric<S, TagRangeMultiline> 
where
    S: Deref<Target = Worksheet>,
{
    fn get_range(&self) -> &Range {
        &self.range
    }

    fn get_sheet(&self) -> &Worksheet {
        &self.sheet
    }

    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        contains_impl(self, other, o_use_rows, o_use_cols)
    }

    fn compare_range(&self, other: &dyn IRange, strict: bool, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        compare_multiline_range_impl(self, other, strict, o_use_rows, o_use_cols)
    }

    fn compare_cell(&self, col_this: u32, row_this: u32, other: &dyn IRange, col_other: u32, row_other: u32, strict: bool) -> bool {
        compare_cell_impl(self, col_this, row_this, other, col_other, row_other, strict)
    }

    fn get_type(&self) -> IterRowNextKind { IterRowNextKind::Multiline }

    fn reset_iter(&mut self)
    {
        reset_iter_impl(self);
    }
}

// Explicit mutable interface
impl<'a> IRangeMut for RangeMultilineMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        self.sheet
    }
}

impl<S> PartialEq for RangeGeneric<S, TagRangeMultiline> {
    fn eq(&self, _other: &Self) -> bool { false }
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

impl<'a> RangeType<'a> {
    pub fn basic(range: Range, sheet: &'a Worksheet) -> Self {
        Self::Basic(RangeBasic::new(range, sheet))
    }

    pub fn merged(range: Range, sheet: &'a Worksheet) -> Self {
        Self::Merged(RangeMergedCells::new(range, sheet))
    }

    pub fn multiline(range: Range, sheet: &'a Worksheet) -> Self {
        Self::Multiline(RangeMultiline::new(range, sheet))
    }
}

impl<'a> RangeTypeMut<'a> {
    pub fn basic(range: Range, sheet: &'a mut Worksheet) -> Self {
        Self::Basic(RangeBasicMut::new(range, sheet))
    }

    pub fn merged(range: Range, sheet: &'a mut Worksheet) -> Self {
        Self::Merged(RangeMergedCellsMut::new(range, sheet))
    }

    pub fn multiline(range: Range, sheet: &'a mut Worksheet) -> Self {
        Self::Multiline(RangeMultilineMut::new(range, sheet))
    }
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

    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        match self {
            RangeType::Basic(r) => r.contains(other, o_use_rows, o_use_cols),
            RangeType::Merged(r) => r.contains(other, o_use_rows, o_use_cols),
            RangeType::Multiline(r) => r.contains(other, o_use_rows, o_use_cols),
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

    fn get_type(&self) -> IterRowNextKind {
        match self {
            RangeType::Basic(r) => r.get_type(),
            RangeType::Merged(r) => r.get_type(),
            RangeType::Multiline(r) => r.get_type(),
        }
    }

    fn reset_iter(&mut self) {
        match self {
            RangeType::Basic(r) => r.reset_iter(),
            RangeType::Merged(r) => r.reset_iter(),
            RangeType::Multiline(r) => r.reset_iter(),
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

    fn contains(&self, other: &dyn IRange, o_use_rows: Option<vec::Vec<u32>>, o_use_cols: Option<vec::Vec<u32>>) -> bool {
        match self {
            RangeTypeMut::Basic(r) => r.contains(other, o_use_rows, o_use_cols),
            RangeTypeMut::Merged(r) => r.contains(other, o_use_rows, o_use_cols),
            RangeTypeMut::Multiline(r) => r.contains(other, o_use_rows, o_use_cols),
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

    fn get_type(&self) -> IterRowNextKind {
        match self {
            RangeTypeMut::Basic(r) => r.get_type(),
            RangeTypeMut::Merged(r) => r.get_type(),
            RangeTypeMut::Multiline(r) => r.get_type(),
        }
    }

    fn reset_iter(&mut self) {
        match self {
            RangeTypeMut::Basic(r) => r.reset_iter(),
            RangeTypeMut::Merged(r) => r.reset_iter(),
            RangeTypeMut::Multiline(r) => r.reset_iter(),
        }
    }
}

impl<'a> IRangeMut for RangeTypeMut<'a> {
    fn get_sheet_mut(&mut self) -> &mut Worksheet {
        match self {
            RangeTypeMut::Basic(r)      => r.get_sheet_mut(),
            RangeTypeMut::Merged(r)     => r.get_sheet_mut(),
            RangeTypeMut::Multiline(r)  => r.get_sheet_mut(),
        }
    }
}

impl<'a> Iterator for RangeType<'a>
{
    type Item = Range;

    fn next(&mut self) -> Option<Self::Item>
    {
        match self
        {
            RangeType::Basic(r)     => Iterator::next(r),
            RangeType::Merged(r)    => Iterator::next(r),
            RangeType::Multiline(r) => Iterator::next(r),
        }
    }
}

impl<'a> Iterator for RangeTypeMut<'a>
{
    type Item = Range;

    fn next(&mut self) -> Option<Self::Item>
    {
        match self
        {
            RangeTypeMut::Basic(r)     => Iterator::next(r),
            RangeTypeMut::Merged(r)    => Iterator::next(r),
            RangeTypeMut::Multiline(r) => Iterator::next(r),
        }
    }
}

//This implementation is not used and most likely incorrect! See above!
impl<'a> LendingIterator for RangeTypeMut<'a>
{
    type Item<'this> = &'this mut RangeTypeMut<'a> where Self: 'this;

    fn next(&mut self) -> Option<Self::Item<'_>>
    {
        let internal_has_next = match self
        {
            RangeTypeMut::Basic(r)     => LendingIterator::next(r).is_some(),
            RangeTypeMut::Merged(r)    => LendingIterator::next(r).is_some(),
            RangeTypeMut::Multiline(r) => LendingIterator::next(r).is_some(),
        };

        if internal_has_next
        {
            Some(self)
        }
        else
        {
            None
        }
    }
}

pub fn make_range_inst(instance_type: IterRowNextKind, range: Range, sheet: &Worksheet) -> RangeType<'_>
{
    match instance_type {
        IterRowNextKind::Basic => RangeType::basic(range, sheet),
        IterRowNextKind::Merged => RangeType::merged(range, sheet),
        IterRowNextKind::Multiline => RangeType::multiline(range, sheet),
    }
}

pub fn make_range_inst_mut(instance_type: IterRowNextKind, range: Range, sheet: &mut Worksheet) -> RangeTypeMut<'_>
{
    match instance_type {
        IterRowNextKind::Basic => RangeTypeMut::basic(range, sheet),
        IterRowNextKind::Merged => RangeTypeMut::merged(range, sheet),
        IterRowNextKind::Multiline => RangeTypeMut::multiline(range, sheet),
    }
}

/* ---------------------------------------------------------------------------------------- */

#[cfg(test)]
mod tests {
    use super::*;
    use umya_spreadsheet::{Cell, CellValue, RichText, TextElement, Font};

    // Sheets are built directly in memory; compare_cell_impl only reads cells.
    fn cc_sheet(name: &str, cells: &[(u32, u32, &str)]) -> Worksheet {
        let mut sheet = Worksheet::default();
        sheet.set_name(name);
        for &(c, r, v) in cells {
            let mut value = CellValue::default();
            value.set_value_string(v.to_string());
            let mut cell = Cell::default();
            cell.set_coordinate((c, r));
            cell.set_cell_value(value);
            sheet.set_cell(cell);
        }
        sheet
    }

    fn cc_rich_sheet(name: &str, col: u32, row: u32, text: &str, bold: bool) -> Worksheet {
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

    fn a1() -> Range { range_ops::make_range_from_strings("A1", "A1") }

    // ------------------------------------------------------------------
    // compare_cell_impl (direct calls, RangeBasic used as the 'this' IRange)
    // ------------------------------------------------------------------

    #[test]
    fn test_0106_compare_cell_impl_strings() {
        let a = cc_sheet("RT01A", &[(1, 1, "ABC\tX")]);
        let b_eq = cc_sheet("RT01B", &[(2, 2, "abc x")]);
        let b_diff = cc_sheet("RT01C", &[(1, 1, "abd")]);
        let b_empty = cc_sheet("RT01D", &[]);
        let ra = RangeBasic::new(a1(), &a);

        assert!(compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_eq), 2, 2, true)); // case/ws-insensitive, other coords may differ
        assert!(!compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_diff), 1, 1, true));
        assert!(!compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_diff), 1, 1, false));
        assert!(compare_cell_impl(&ra, 5, 5, &RangeBasic::new(a1(), &b_empty), 5, 5, true)); // both missing -> "" == ""
        assert!(!compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_empty), 5, 5, true)); // one missing
    }

    #[test]
    fn test_0107_compare_cell_impl_rich_text_equal_strict() {
        let a = cc_rich_sheet("RT02A", 1, 1, "text", false);
        let b = cc_rich_sheet("RT02B", 1, 1, "text", false);
        let ra = RangeBasic::new(a1(), &a);
        assert!(compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b), 1, 1, true));
    }

    #[test]
    fn test_0108_compare_cell_impl_rich_text_mismatch() {
        let a = cc_rich_sheet("RT03A", 1, 1, "text", false);
        let b_bold = cc_rich_sheet("RT03B", 1, 1, "text", true); // same text, different formatting
        let ra = RangeBasic::new(a1(), &a);
        assert!(!compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_bold), 1, 1, true));
        assert!(compare_cell_impl(&ra, 1, 1, &RangeBasic::new(a1(), &b_bold), 1, 1, false)); // non-strict ignores formatting
    }

    // ------------------------------------------------------------------
    // IRange::compare_cell: the same impl reached through each IRange type
    // ------------------------------------------------------------------

    #[test]
    fn test_0109_compare_cell_range_basic() {
        let a = cc_sheet("RT04A", &[(1, 1, "x")]);
        let b_eq = cc_sheet("RT04B", &[(1, 1, "X")]);
        let b_diff = cc_sheet("RT04C", &[(1, 1, "y")]);
        let ra = RangeBasic::new(a1(), &a);
        assert!(ra.compare_cell(1, 1, &RangeBasic::new(a1(), &b_eq), 1, 1, true));
        assert!(!ra.compare_cell(1, 1, &RangeBasic::new(a1(), &b_diff), 1, 1, true));
    }

    #[test]
    fn test_0110_compare_cell_range_merged() {
        let a = cc_sheet("RT05A", &[(1, 1, "x")]);
        let b_eq = cc_sheet("RT05B", &[(1, 1, "X")]);
        let b_diff = cc_sheet("RT05C", &[(1, 1, "y")]);
        let ra = RangeMergedCells::new(a1(), &a);
        assert!(ra.compare_cell(1, 1, &RangeMergedCells::new(a1(), &b_eq), 1, 1, true));
        assert!(!ra.compare_cell(1, 1, &RangeMergedCells::new(a1(), &b_diff), 1, 1, true));
    }

    #[test]
    fn test_0111_compare_cell_range_multiline() {
        let a = cc_sheet("RT06A", &[(1, 1, "x")]);
        let b_eq = cc_sheet("RT06B", &[(1, 1, "X")]);
        let b_diff = cc_sheet("RT06C", &[(1, 1, "y")]);
        let ra = RangeMultiline::new(a1(), &a);
        assert!(ra.compare_cell(1, 1, &RangeMultiline::new(a1(), &b_eq), 1, 1, true));
        assert!(!ra.compare_cell(1, 1, &RangeMultiline::new(a1(), &b_diff), 1, 1, true));
    }

    #[test]
    fn test_0112_compare_cell_cross_types() {
        // 'this' and 'other' may be different IRange implementations
        let a = cc_sheet("RT07A", &[(1, 1, "x")]);
        let b = cc_sheet("RT07B", &[(1, 1, "x")]);
        let ra = RangeBasic::new(a1(), &a);
        let rb = RangeMultiline::new(a1(), &b);
        assert!(ra.compare_cell(1, 1, &rb, 1, 1, true));
    }

    #[test]
    fn test_0113_compare_cell_range_type_enum() {
        let a = cc_sheet("RT08A", &[(1, 1, "x")]);
        let b = cc_sheet("RT08B", &[(1, 1, "X")]);
        let c = cc_sheet("RT08C", &[(1, 1, "y")]);
        for mk in [RangeType::basic, RangeType::merged, RangeType::multiline] {
            let ra = mk(a1(), &a);
            assert!(ra.compare_cell(1, 1, &mk(a1(), &b), 1, 1, true), "variant not compared equal");
            assert!(!ra.compare_cell(1, 1, &mk(a1(), &c), 1, 1, true), "variant not compared different");
        }
    }

    // fresh sheets per call: the &mut borrows must not outlive the function.
    // goes through make_range_inst_mut, which also exercises the variant dispatch
    fn check_mut_variant(kind: IterRowNextKind) {
        let mut a = cc_sheet("TMP_A", &[(1, 1, "x")]);
        let mut b = cc_sheet("TMP_B", &[(1, 1, "X")]);
        let mut c = cc_sheet("TMP_C", &[(1, 1, "y")]);
        let ra = make_range_inst_mut(kind, a1(), &mut a);
        let rb = make_range_inst_mut(kind, a1(), &mut b);
        let rc = make_range_inst_mut(kind, a1(), &mut c);
        assert!(ra.compare_cell(1, 1, &rb, 1, 1, true));
        assert!(!ra.compare_cell(1, 1, &rc, 1, 1, true));
    }

    #[test]
    fn test_0114_compare_cell_range_type_mut_enum() {
        check_mut_variant(IterRowNextKind::Basic);
        check_mut_variant(IterRowNextKind::Merged);
        check_mut_variant(IterRowNextKind::Multiline);
    }

    // ------------------------------------------------------------------
    // IRange::compare_range, per variant. Documented quirks (same family
    // as the comapre_ranges tests in range_ops.rs, test_0046..0054):
    //  - strict mode (Basic/Multiline) only checks sizes, content is ignored
    //  - Basic/Multiline build column offsets 0..=cols (one past the range
    //    end); Basic requires EXACTLY `cols` matches, so a perfect row match
    //    FAILS in non-strict mode
    //  - rows are compared in a cartesian product and row_match is compared
    //    against rows_a (Basic/Merged) or rows_a+1 (Multiline)
    //  - Merged always applies the size check AND the row_match check, so
    //    strict and non-strict behave the same
    // ------------------------------------------------------------------
    fn a1a2() -> Range { range_ops::make_range_from_strings("A1", "A2") }

    #[test]
    fn test_0115_compare_range_basic_strict_size_mismatch() {
        let a = cc_sheet("RR01A", &[(1, 1, "x")]);
        let b = cc_sheet("RR01B", &[(1, 1, "x")]);
        let ra = RangeBasic::new(range_ops::make_range_from_strings("A1", "C3"), &a);
        assert!(!ra.compare_range(&RangeBasic::new(range_ops::make_range_from_strings("A1", "C2"), &b), true, None, None));
        // invalid range on either side
        assert!(!ra.compare_range(&RangeBasic::new(Range::default(), &b), true, None, None));
        assert!(!RangeBasic::new(a1(), &a).compare_range(&RangeBasic::new(Range::default(), &b), false, None, None));
    }

    #[test]
    fn test_0116_compare_range_basic_strict_ignores_content() {
        let a = cc_sheet("RR02A", &[(1, 1, "aaa"), (1, 2, "b")]);
        let b = cc_sheet("RR02B", &[(1, 1, "ccc"), (1, 2, "d")]);
        let ra = RangeBasic::new(a1a2(), &a);
        // Documented current behavior: strict only checks sizes.
        // Expected to flip to `false` if fixed.
        assert!(ra.compare_range(&RangeBasic::new(a1a2(), &b), true, None, None));
    }

    #[test]
    fn test_0117_compare_range_basic_nonstrict_perfect_1x1_fails() {
        let a = cc_sheet("RR03A", &[(1, 1, "x")]);
        let b = cc_sheet("RR03B", &[(1, 1, "x")]);
        let ra = RangeBasic::new(a1(), &a);
        // Documented current behavior: offsets run 0..=cols, so a fully equal
        // row scores cols+1 matches while the check requires exactly cols.
        // Expected to flip to `true` if fixed.
        assert!(!ra.compare_range(&RangeBasic::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0118_compare_range_basic_nonstrict_single_mismatch_passes() {
        let a = cc_sheet("RR04A", &[(1, 1, "x"), (2, 1, "p")]);
        let b = cc_sheet("RR04B", &[(1, 1, "x"), (2, 1, "q")]);
        let ra = RangeBasic::new(a1(), &a);
        // exactly one mismatch (in the compared-but-out-of-range column B)
        // yields col_match == cols, so the row passes
        assert!(ra.compare_range(&RangeBasic::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0119_compare_range_basic_nonstrict_cartesian_rows() {
        // A's rows both equal B's LAST row, while B's first row differs from A
        // (col 2 always differs, so a row pair needs the col-1 match only)
        let a = cc_sheet("RR05A", &[(1, 1, "same"), (2, 1, "p1"), (1, 2, "same"), (2, 2, "p2")]);
        let b = cc_sheet("RR05B", &[(1, 1, "other"), (2, 1, "q1"), (1, 2, "same"), (2, 2, "q2")]);
        let ra = RangeBasic::new(a1a2(), &a);
        // Documented current behavior: cartesian row loop, row_match counts every
        // matching (row_a, row_b) pair; here exactly rows_a pairs match (both A
        // rows vs B's last row). Expected to change if fixed.
        assert!(ra.compare_range(&RangeBasic::new(a1a2(), &b), false, None, None));

        // Control: B's last row differs too -> no pair matches
        let b2 = cc_sheet("RR05C", &[(1, 1, "other"), (2, 1, "q1"), (1, 2, "nope"), (2, 2, "q2")]);
        assert!(!ra.compare_range(&RangeBasic::new(a1a2(), &b2), false, None, None));
    }

    #[test]
    fn test_0120_compare_range_merged_1x1_equal() {
        let a = cc_sheet("RR06A", &[(1, 1, "x")]);
        let b = cc_sheet("RR06B", &[(1, 1, "X")]);
        let ra = RangeMergedCells::new(a1(), &a);
        // single row: exactly one matching pair, row_match == rows_a
        assert!(ra.compare_range(&RangeMergedCells::new(a1(), &b), true, None, None));
        assert!(ra.compare_range(&RangeMergedCells::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0121_compare_range_merged_size_mismatch_always_fails() {
        let a = cc_sheet("RR07A", &[(1, 1, "x")]);
        let b = cc_sheet("RR07B", &[(1, 1, "x")]);
        let ra = RangeMergedCells::new(a1a2(), &a);
        // the size check applies regardless of strict
        assert!(!ra.compare_range(&RangeMergedCells::new(a1(), &b), true, None, None));
        assert!(!ra.compare_range(&RangeMergedCells::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0122_compare_range_merged_content_differs_same_size() {
        let a = cc_sheet("RR08A", &[(1, 1, "x")]);
        let b = cc_sheet("RR08B", &[(1, 1, "y")]);
        let ra = RangeMergedCells::new(a1(), &a);
        // unlike Basic/Multiline, Merged checks content in strict mode too
        assert!(!ra.compare_range(&RangeMergedCells::new(a1(), &b), true, None, None));
        assert!(!ra.compare_range(&RangeMergedCells::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0123_compare_range_merged_multi_row() {
        // perfectly equal 2-row ranges: all 4 cartesian pairs match,
        // row_match=4 != rows_a=2 -> fails
        let a = cc_sheet("RR09A", &[(1, 1, "s"), (1, 2, "s")]);
        let b = cc_sheet("RR09B", &[(1, 1, "s"), (1, 2, "s")]);
        let ra = RangeMergedCells::new(a1a2(), &a);
        // Documented current behavior. Expected to flip to `true` if fixed.
        assert!(!ra.compare_range(&RangeMergedCells::new(a1a2(), &b), false, None, None));

        // but only B's last row matching both A rows gives row_match == rows_a
        let b2 = cc_sheet("RR09C", &[(1, 1, "other"), (1, 2, "s")]);
        assert!(ra.compare_range(&RangeMergedCells::new(a1a2(), &b2), false, None, None));
    }

    #[test]
    fn test_0124_compare_range_multiline_strict_size_mismatch() {
        let a = cc_sheet("RR10A", &[(1, 1, "x")]);
        let b = cc_sheet("RR10B", &[(1, 1, "x")]);
        let ra = RangeMultiline::new(a1a2(), &a);
        assert!(!ra.compare_range(&RangeMultiline::new(a1(), &b), true, None, None));
        // invalid range on either side
        assert!(!ra.compare_range(&RangeMultiline::new(Range::default(), &b), true, None, None));
        assert!(!RangeMultiline::new(a1(), &a).compare_range(&RangeMultiline::new(Range::default(), &b), false, None, None));
    }

    #[test]
    fn test_0125_compare_range_multiline_strict_ignores_content() {
        let a = cc_sheet("RR11A", &[(1, 1, "a"), (1, 2, "b")]);
        let b = cc_sheet("RR11B", &[(1, 1, "c"), (1, 2, "d")]);
        let ra = RangeMultiline::new(a1a2(), &a);
        // Documented current behavior: strict only checks sizes.
        // Expected to flip to `false` if fixed.
        assert!(ra.compare_range(&RangeMultiline::new(a1a2(), &b), true, None, None));
    }

    #[test]
    fn test_0126_compare_range_multiline_nonstrict_1x1_fails() {
        let a = cc_sheet("RR12A", &[(1, 1, "x")]);
        let b = cc_sheet("RR12B", &[(1, 1, "x")]);
        let ra = RangeMultiline::new(a1(), &a);
        // Documented current behavior: row_match (1 matching pair) is compared
        // against rows+1 (2), so even identical ranges fail.
        // Expected to flip to `true` if fixed.
        assert!(!ra.compare_range(&RangeMultiline::new(a1(), &b), false, None, None));
    }

    #[test]
    fn test_0127_compare_range_multiline_nonstrict_match() {
        // A is 1 row; B has 2 rows both equal to A's row (incl. the
        // compared-but-out-of-range column B). 2 matching pairs == rows+1
        let a = cc_sheet("RR13A", &[(1, 1, "x")]);
        let b = cc_sheet("RR13B", &[(1, 1, "x"), (1, 2, "x")]);
        let ra = RangeMultiline::new(a1(), &a);
        assert!(ra.compare_range(&RangeMultiline::new(a1a2(), &b), false, None, None));
    }

    // ------------------------------------------------------------------
    // compare_range dispatch through the RangeType / RangeTypeMut enums
    // ------------------------------------------------------------------
    fn check_range_variant(kind: IterRowNextKind) {
        let a = cc_sheet("TMP_A", &[(1, 1, "x")]);
        let b = cc_sheet("TMP_B", &[(1, 1, "X")]);
        let ra = make_range_inst(kind, a1(), &a);
        let rb = make_range_inst(kind, a1(), &b);
        assert!(ra.compare_range(&rb, true, None, None));
        // strict size mismatch
        let rc = make_range_inst(kind, a1a2(), &a);
        assert!(!rc.compare_range(&rb, true, None, None));
    }

    #[test]
    fn test_0128_compare_range_range_type_enum() {
        check_range_variant(IterRowNextKind::Basic);
        check_range_variant(IterRowNextKind::Merged);
        check_range_variant(IterRowNextKind::Multiline);
    }

    fn check_mut_range_variant(kind: IterRowNextKind) {
        let mut a = cc_sheet("TMP_A", &[(1, 1, "x")]);
        let mut b = cc_sheet("TMP_B", &[(1, 1, "X")]);
        {
            let ra = make_range_inst_mut(kind, a1(), &mut a);
            let rb = make_range_inst_mut(kind, a1(), &mut b);
            assert!(ra.compare_range(&rb, true, None, None));
        } // drops ra/rb, releasing the &mut sheet borrows
        let rc = make_range_inst_mut(kind, a1a2(), &mut a);
        let rb = make_range_inst_mut(kind, a1(), &mut b);
        assert!(!rc.compare_range(&rb, true, None, None));
    }

    #[test]
    fn test_0129_compare_range_range_type_mut_enum() {
        check_mut_range_variant(IterRowNextKind::Basic);
        check_mut_range_variant(IterRowNextKind::Merged);
        check_mut_range_variant(IterRowNextKind::Multiline);
    }

    // ------------------------------------------------------------------
    // IRange::contains (all variants dispatch to contains_impl).
    // Semantics: 'other' must be contained in 'this' per o_use_cols —
    // the header rows (first row of each range) must match and every data
    // row of 'other' must be found somewhere in 'this' data rows (order
    // does not matter). Documented quirks:
    //  - o_use_cols=None -> empty column list -> always false
    //  - o_use_rows is accepted but ignored
    //  - 'other' rows are NOT consumed, duplicates can match one 'this' row
    //  - with several cmp_cols, res is whatever the LAST checked column set;
    //    a header mismatch breaks the loop without resetting res
    // ------------------------------------------------------------------
    fn a1a3() -> Range { range_ops::make_range_from_strings("A1", "A3") }

    #[test]
    fn test_0130_contains_requires_matching_types() {
        let a = cc_sheet("CT01A", &[(1, 1, "h"), (1, 2, "a")]);
        let b = cc_sheet("CT01B", &[(1, 1, "h"), (1, 2, "a")]);
        let tb = RangeBasic::new(a1a2(), &a);
        let tm = RangeMultiline::new(a1a2(), &b);
        assert!(!tb.contains(&tm, None, Some(vec![1])));
        assert!(!tm.contains(&tb, None, Some(vec![1])));
    }

    #[test]
    fn test_0131_contains_requires_use_cols() {
        let a = cc_sheet("CT02A", &[(1, 1, "h"), (1, 2, "a")]);
        let b = cc_sheet("CT02B", &[(1, 1, "h"), (1, 2, "a")]);
        let ra = RangeBasic::new(a1a2(), &a);
        // identical data, but no columns to compare -> false
        assert!(!ra.contains(&RangeBasic::new(a1a2(), &b), None, None));
    }

    #[test]
    fn test_0132_contains_true_order_independent() {
        let a = cc_sheet("CT03A", &[(1, 1, "hdr"), (1, 2, "a"), (1, 3, "b"), (1, 4, "c")]);
        let b = cc_sheet("CT03B", &[(1, 1, "hdr"), (1, 2, "c"), (1, 3, "a")]);
        let ra = RangeBasic::new(range_ops::make_range_from_strings("A1", "A4"), &a);
        assert!(ra.contains(&RangeBasic::new(a1a3(), &b), None, Some(vec![1])));
        // o_use_rows is accepted but ignored
        assert!(ra.contains(&RangeBasic::new(a1a3(), &b), Some(vec![1]), Some(vec![1])));
    }

    #[test]
    fn test_0133_contains_missing_entry_false() {
        let a = cc_sheet("CT04A", &[(1, 1, "hdr"), (1, 2, "a"), (1, 3, "b")]);
        let b = cc_sheet("CT04B", &[(1, 1, "hdr"), (1, 2, "a"), (1, 3, "z")]);
        let ra = RangeBasic::new(a1a3(), &a);
        assert!(!ra.contains(&RangeBasic::new(a1a3(), &b), None, Some(vec![1])));
    }

    #[test]
    fn test_0134_contains_header_mismatch_false() {
        let a = cc_sheet("CT05A", &[(1, 1, "hdr"), (1, 2, "a")]);
        let b = cc_sheet("CT05B", &[(1, 1, "H2"), (1, 2, "a")]);
        let ra = RangeBasic::new(a1a2(), &a);
        assert!(!ra.contains(&RangeBasic::new(a1a2(), &b), None, Some(vec![1])));
    }

    #[test]
    fn test_0135_contains_header_only_other() {
        let a = cc_sheet("CT06A", &[(1, 1, "hdr"), (1, 2, "a")]);
        let b_ok = cc_sheet("CT06B", &[(1, 1, "hdr")]);
        let b_bad = cc_sheet("CT06C", &[(1, 1, "x")]);
        let ra = RangeBasic::new(a1a2(), &a);
        // single-row 'other': only the header has to match
        assert!(ra.contains(&RangeBasic::new(a1(), &b_ok), None, Some(vec![1])));
        assert!(!ra.contains(&RangeBasic::new(a1(), &b_bad), None, Some(vec![1])));
    }

    #[test]
    fn test_0136_contains_duplicate_rows_not_consumed() {
        let a = cc_sheet("CT07A", &[(1, 1, "h"), (1, 2, "a")]);
        let b = cc_sheet("CT07B", &[(1, 1, "h"), (1, 2, "a"), (1, 3, "a")]);
        let ra = RangeBasic::new(a1a2(), &a);
        // Documented current behavior: 'other' rows are not consumed, so both
        // "a" rows match the single "a" row of 'this'.
        assert!(ra.contains(&RangeBasic::new(a1a3(), &b), None, Some(vec![1])));
    }

    #[test]
    fn test_0137_contains_multi_col_last_checked_decides() {
        // this: col A ["h1","b","c"], col B ["x1","p","q"];
        // other: col A ["h1","b"], col B ["x2","q"] (col B header differs)
        let a = cc_sheet("CT08A", &[(1, 1, "h1"), (1, 2, "b"), (1, 3, "c"), (2, 1, "x1"), (2, 2, "p"), (2, 3, "q")]);
        let b = cc_sheet("CT08B", &[(1, 1, "h1"), (1, 2, "b"), (2, 1, "x2"), (2, 2, "q")]);
        let ra = RangeBasic::new(range_ops::make_range_from_strings("A1", "B3"), &a);
        let rb = RangeBasic::new(range_ops::make_range_from_strings("A1", "B2"), &b);
        // Documented current behavior: col 1 passes (res=true), then col 2's
        // header mismatch breaks the loop WITHOUT resetting res.
        assert!(ra.contains(&rb, None, Some(vec![1, 2])));
        // reversed order: col 2 mismatches first -> res stays false
        assert!(!ra.contains(&rb, None, Some(vec![2, 1])));
    }

    #[test]
    fn test_0138_contains_invalid_range() {
        let a = cc_sheet("CT09A", &[(1, 1, "h"), (1, 2, "a")]);
        let b = cc_sheet("CT09B", &[(1, 1, "h"), (1, 2, "a")]);
        let ra = RangeBasic::new(a1a2(), &a);
        assert!(!RangeBasic::new(Range::default(), &a).contains(&RangeBasic::new(a1a2(), &b), None, Some(vec![1])));
        assert!(!ra.contains(&RangeBasic::new(Range::default(), &b), None, Some(vec![1])));
    }

    #[test]
    fn test_0139_contains_missing_cells_equal() {
        // missing cells compare as "" == ""
        let a = cc_sheet("CT10A", &[(1, 1, "h"), (1, 2, "a")]); // A3 missing
        let b = cc_sheet("CT10B", &[(1, 1, "h")]);               // A2 missing
        let ra = RangeBasic::new(range_ops::make_range_from_strings("A1", "A3"), &a);
        assert!(ra.contains(&RangeBasic::new(a1a2(), &b), None, Some(vec![1])));
    }

    // ------------------------------------------------------------------
    // contains dispatch through the RangeType / RangeTypeMut enums
    // ------------------------------------------------------------------
    fn check_contains_variant(kind: IterRowNextKind) {
        let a = cc_sheet("TMP_A", &[(1, 1, "hdr"), (1, 2, "a"), (1, 3, "b")]);
        let b = cc_sheet("TMP_B", &[(1, 1, "HDR"), (1, 2, "b")]);
        let ra = make_range_inst(kind, range_ops::make_range_from_strings("A1", "A3"), &a);
        let rb = make_range_inst(kind, a1a2(), &b);
        assert!(ra.contains(&rb, None, Some(vec![1])));
    }

    #[test]
    fn test_0140_contains_range_type_enum() {
        check_contains_variant(IterRowNextKind::Basic);
        check_contains_variant(IterRowNextKind::Merged);
        check_contains_variant(IterRowNextKind::Multiline);
        // dispatch keeps the same-types rule
        let a = cc_sheet("CT11A", &[(1, 1, "h"), (1, 2, "a")]);
        let b = cc_sheet("CT11B", &[(1, 1, "h"), (1, 2, "a")]);
        let rb = make_range_inst(IterRowNextKind::Basic, a1a2(), &b);
        let rm = make_range_inst(IterRowNextKind::Multiline, a1a2(), &a);
        assert!(!rb.contains(&rm, None, Some(vec![1])));
    }

    fn check_mut_contains_variant(kind: IterRowNextKind) {
        let mut a = cc_sheet("TMP_A", &[(1, 1, "hdr"), (1, 2, "a"), (1, 3, "b")]);
        let mut b = cc_sheet("TMP_B", &[(1, 1, "HDR"), (1, 2, "b")]);
        let ra = make_range_inst_mut(kind, range_ops::make_range_from_strings("A1", "A3"), &mut a);
        let rb = make_range_inst_mut(kind, a1a2(), &mut b);
        assert!(ra.contains(&rb, None, Some(vec![1])));
    }

    #[test]
    fn test_0141_contains_range_type_mut_enum() {
        check_mut_contains_variant(IterRowNextKind::Basic);
        check_mut_contains_variant(IterRowNextKind::Merged);
        check_mut_contains_variant(IterRowNextKind::Multiline);
    }

    // ------------------------------------------------------------------
    // IRange::get_type / get_type_name / get_range / get_sheet / reset_iter
    // ------------------------------------------------------------------

    fn a2c4() -> Range { range_ops::make_range_from_strings("A2", "C4") }

    #[test]
    fn test_0142_get_type_per_variant() {
        let a = cc_sheet("GT01A", &[(1, 1, "x")]);
        assert_eq!(RangeBasic::new(a1(), &a).get_type(), IterRowNextKind::Basic);
        assert_eq!(RangeMergedCells::new(a1(), &a).get_type(), IterRowNextKind::Merged);
        assert_eq!(RangeMultiline::new(a1(), &a).get_type(), IterRowNextKind::Multiline);
    }

    #[test]
    fn test_0143_get_type_name_per_variant() {
        let a = cc_sheet("GT02A", &[(1, 1, "x")]);
        assert_eq!(RangeBasic::new(a1(), &a).get_type_name(), "Basic");
        assert_eq!(RangeMergedCells::new(a1(), &a).get_type_name(), "Merged");
        assert_eq!(RangeMultiline::new(a1(), &a).get_type_name(), "Multiline");
    }

    #[test]
    fn test_0144_get_range_returns_constructed_range() {
        let a = cc_sheet("GT03A", &[(1, 1, "x")]);
        assert_eq!(RangeBasic::new(a1(), &a).get_range().get_range(), "A1:A1");
        assert_eq!(RangeMergedCells::new(a1(), &a).get_range().get_range(), "A1:A1");
        assert_eq!(RangeMultiline::new(a1(), &a).get_range().get_range(), "A1:A1");
        let r = RangeBasic::new(a2c4(), &a);
        assert_eq!(r.get_range().get_range(), "A2:C4");
        assert_eq!(range_ops::range_bounds(r.get_range()), Some((2, 4, 1, 3, 3, 3)));
    }

    #[test]
    fn test_0145_get_sheet_returns_backing_sheet() {
        let a = cc_sheet("GT04A", &[(1, 1, "x")]);
        assert_eq!(RangeBasic::new(a1(), &a).get_sheet().get_name(), "GT04A");
        assert_eq!(RangeMergedCells::new(a1(), &a).get_sheet().get_name(), "GT04A");
        assert_eq!(RangeMultiline::new(a1(), &a).get_sheet().get_name(), "GT04A");
    }

    #[test]
    fn test_0146_reset_iter_restores_range_start_row() {
        let a = cc_sheet("GT05A", &[]);
        // start row 2 (range A2:C4): reset must restore 2, not the row-1 default
        let mut rb = RangeBasic::new(a2c4(), &a);
        rb.current_row = 9;
        rb.reset_iter();
        assert_eq!(rb.current_row, 2);
        let mut rm = RangeMergedCells::new(a2c4(), &a);
        rm.current_row = 9;
        rm.reset_iter();
        assert_eq!(rm.current_row, 2);
        let mut rml = RangeMultiline::new(a2c4(), &a);
        rml.current_row = 9;
        rml.reset_iter();
        assert_eq!(rml.current_row, 2);
    }

    #[test]
    fn test_0147_reset_iter_after_iteration() {
        let a = cc_sheet("GT06A", &[]);
        let mut r = RangeBasic::new(a2c4(), &a);
        assert_eq!(r.current_row, 2); // new() starts at the range start row
        Iterator::next(&mut r);
        Iterator::next(&mut r);
        assert_eq!(r.current_row, 4);
        r.reset_iter();
        assert_eq!(r.current_row, 2);
        // exhaust, then reset must make the iterator usable again (rows 2..=4)
        while Iterator::next(&mut r).is_some() {}
        assert_eq!(r.current_row, 5);
        r.reset_iter();
        assert_eq!(r.by_ref().count(), 3);
    }

    #[test]
    fn test_0148_reset_iter_invalid_range_row1_fallback() {
        let a = cc_sheet("GT07A", &[]);
        let mut r = RangeBasic::new(Range::default(), &a);
        assert_eq!(r.current_row, 1); // new() also falls back to row 1
        r.current_row = 9;
        r.reset_iter();
        assert_eq!(r.current_row, 1);
    }

    // getters + reset_iter dispatch through the RangeType / RangeTypeMut enums
    fn check_getters<T>(kind: IterRowNextKind, mut r: T)
    where
        T: IRange + Iterator<Item = Range>,
    {
        let expected_name = match kind {
            IterRowNextKind::Basic => "Basic",
            IterRowNextKind::Merged => "Merged",
            IterRowNextKind::Multiline => "Multiline",
        };
        assert_eq!(r.get_type(), kind);
        assert_eq!(r.get_type_name(), expected_name);
        assert_eq!(r.get_range().get_range(), "A2:C4");
        assert_eq!(r.get_sheet().get_name(), "GT08A");
        while Iterator::next(&mut r).is_some() {} // exhaust rows 2..=4
        r.reset_iter();
        assert_eq!(r.by_ref().count(), 3);
    }

    #[test]
    fn test_0149_getters_range_type_enum() {
        let a = cc_sheet("GT08A", &[]);
        check_getters(IterRowNextKind::Basic, RangeType::basic(a2c4(), &a));
        check_getters(IterRowNextKind::Merged, RangeType::merged(a2c4(), &a));
        check_getters(IterRowNextKind::Multiline, RangeType::multiline(a2c4(), &a));
    }

    #[test]
    fn test_0150_getters_range_type_mut_enum() {
        let mut a = cc_sheet("GT08A", &[]);
        {
            let r = make_range_inst_mut(IterRowNextKind::Basic, a2c4(), &mut a);
            check_getters(IterRowNextKind::Basic, r);
        }
        {
            let r = make_range_inst_mut(IterRowNextKind::Merged, a2c4(), &mut a);
            check_getters(IterRowNextKind::Merged, r);
        }
        {
            let r = make_range_inst_mut(IterRowNextKind::Multiline, a2c4(), &mut a);
            check_getters(IterRowNextKind::Multiline, r);
        }
    }

    // ------------------------------------------------------------------
    // Row-by-row iterators: Iterator for RangeGeneric / RangeType /
    // RangeTypeMut (yields one Range per row), LendingIterator impls
    // (source comments: "not used and most likely incorrect" — the tests
    // pin the current behavior), and the always-false PartialEq impls.
    // ------------------------------------------------------------------

    #[test]
    fn test_0151_range_generic_iterator_yields_rows() {
        let a = cc_sheet("IT01A", &[]);
        let mut r = RangeBasic::new(a2c4(), &a);
        // qualified: RangeGeneric also implements LendingIterator (same `next` name)
        assert_eq!(Iterator::next(&mut r).unwrap().get_range(), "A2:C2");
        assert_eq!(Iterator::next(&mut r).unwrap().get_range(), "A3:C3");
        assert_eq!(Iterator::next(&mut r).unwrap().get_range(), "A4:C4");
        assert!(Iterator::next(&mut r).is_none());
        // invalid range yields nothing
        let mut r2 = RangeBasic::new(Range::default(), &a);
        assert!(Iterator::next(&mut r2).is_none());
    }

    #[test]
    fn test_0152_range_type_iterator_yields_rows() {
        let a = cc_sheet("IT02A", &[]);
        for mk in [RangeType::basic, RangeType::merged, RangeType::multiline] {
            let mut rt = mk(a2c4(), &a);
            let rows: Vec<String> = (&mut rt).map(|rr| rr.get_range()).collect();
            assert_eq!(rows, vec!["A2:C2".to_string(), "A3:C3".to_string(), "A4:C4".to_string()]);
        }
    }

    fn check_mut_row_iterator(kind: IterRowNextKind) {
        let mut a = cc_sheet("IT03A", &[]);
        let mut rt = make_range_inst_mut(kind, a2c4(), &mut a);
        let rows: Vec<String> = (&mut rt).map(|rr| rr.get_range()).collect();
        assert_eq!(rows, vec!["A2:C2".to_string(), "A3:C3".to_string(), "A4:C4".to_string()]);
        let cur = match &rt {
            RangeTypeMut::Basic(r) => r.current_row,
            RangeTypeMut::Merged(r) => r.current_row,
            RangeTypeMut::Multiline(r) => r.current_row,
        };
        assert_eq!(cur, 5);
    }

    #[test]
    fn test_0153_range_type_mut_iterator_yields_rows() {
        check_mut_row_iterator(IterRowNextKind::Basic);
        check_mut_row_iterator(IterRowNextKind::Merged);
        check_mut_row_iterator(IterRowNextKind::Multiline);
    }

    #[test]
    fn test_0154_lending_iterator_range_generic() {
        let a = cc_sheet("IT04A", &[]);
        let mut r = RangeBasic::new(a2c4(), &a);
        let mut cnt = 0;
        while let Some(item) = LendingIterator::next(&mut r) {
            // Documented quirk: the item is the WHOLE range object, not the
            // current row (the impl is marked "not used" in the source).
            assert_eq!(item.get_range().get_range(), "A2:C4");
            cnt += 1;
        }
        assert_eq!(cnt, 3);
        assert_eq!(r.current_row, 5);
    }

    fn check_mut_lending_variant(kind: IterRowNextKind) {
        let mut a = cc_sheet("IT05A", &[]);
        let mut rt = make_range_inst_mut(kind, a2c4(), &mut a);
        let mut cnt = 0;
        while let Some(_item) = LendingIterator::next(&mut rt) {
            cnt += 1;
        }
        // each call probes the inner iterator, advancing the row counter once
        let cur = match &rt {
            RangeTypeMut::Basic(r) => r.current_row,
            RangeTypeMut::Merged(r) => r.current_row,
            RangeTypeMut::Multiline(r) => r.current_row,
        };
        assert_eq!(cnt, 3);
        assert_eq!(cur, 5);
    }

    #[test]
    fn test_0155_lending_iterator_range_type_mut() {
        check_mut_lending_variant(IterRowNextKind::Basic);
        check_mut_lending_variant(IterRowNextKind::Merged);
        check_mut_lending_variant(IterRowNextKind::Multiline);
    }

    #[test]
    fn test_0156_partial_eq_always_false() {
        let a = cc_sheet("EQ01A", &[(1, 1, "x")]);
        // no Debug on RangeGeneric, so assert via ==  directly
        assert!(!((RangeBasic::new(a2c4(), &a)) == (RangeBasic::new(a2c4(), &a))));
        assert!(!((RangeMergedCells::new(a2c4(), &a)) == (RangeMergedCells::new(a2c4(), &a))));
        assert!(!((RangeMultiline::new(a2c4(), &a)) == (RangeMultiline::new(a2c4(), &a))));
    }
}
