use crate::_structs::replace::Replaces;
pub trait Convert {
    fn convert(&self) -> Replaces;
}
