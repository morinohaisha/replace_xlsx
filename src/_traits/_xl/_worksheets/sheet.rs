use crate::_structs::_xl::_worksheets::sheet::C;
use crate::_structs::cell::Cell;
pub trait Worksheet {
    fn cells(&self, index: &u32) -> Vec<Cell>;
}

pub trait Row {
    fn refs(&self, index: &u32) -> Vec<Cell>;
    fn string_cells(&self) -> Vec<&C>;
}
