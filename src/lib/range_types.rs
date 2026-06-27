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

#[derive(PartialEq)]
pub enum RangeType<'a> {
    Basic(RangeBasic<'a>),
    Merged(RangeMergedCells<'a>),
    Multiline(RangeMultiline<'a>),
}

#[derive(PartialEq)]
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