// 1. Всички необходими импорти от външната библиотека
use umya_spreadsheet::{Range, Worksheet};

// ==========================================
// TRAIT
// ==========================================

pub trait IRange {
    fn get_range(&self) -> &Range;
    fn get_sheet(&self) -> &Worksheet;
    fn contains(&self, other: &Range) -> bool;
}

// ==========================================
// STRUCTS
// ==========================================

pub struct RangeBasic<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

pub struct RangeMergedCells<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

pub struct RangeMultiline<'a> {
    pub range: Range,
    pub sheet: &'a Worksheet,
}

// ==========================================
// STRUCT IMPLEMENTATION
// ==========================================

// --- RangeBasic ---
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
}

impl<'a> PartialEq for RangeBasic<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

// --- RangeMergedCells ---
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
}

impl<'a> PartialEq for RangeMergedCells<'a> {
    fn eq(&self, _other: &Self) -> bool {
        todo!()
    }
}

// --- RangeMultiline ---
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
}

impl<'a> PartialEq for RangeMultiline<'a> {
    fn eq(&self, _other: &Self) -> bool {
	    todo!()
    }
}

// ==========================================
// ENUM, UNITING THE TYPES
// ==========================================

#[derive(PartialEq)]
pub enum RangeType<'a> {
    Basic(RangeBasic<'a>),
    Merged(RangeMergedCells<'a>),
    Multiline(RangeMultiline<'a>),
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
}
