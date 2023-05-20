use crate::_structs::cell::Cell;
use crate::_structs::image::Image;
use crate::_structs::input::Input;

#[derive(Debug)]
pub struct Replace<'a> {
    pub input: &'a Input,
    pub index: Option<u32>,
    pub cells: Vec<Cell>,
    pub image: Option<Image>,
}

pub type Replaces<'a> = Vec<Replace<'a>>;

pub struct ReplaceXml {
    pub file_name: String,
    pub xml: Vec<u8>,
}

pub type ReplaceXmls = Vec<ReplaceXml>;