use crate::_structs::replace::Replaces;
use crate::_structs::zip::XlsxReader;

pub trait Extract {
    fn extract(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String>;
    fn extract_index(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String>;
    fn extract_cells(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String>;
    fn setup_images(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String>;
}

pub trait Replace {
    fn replace(&mut self, replaces: &Replaces) -> anyhow::Result<Vec<u8>>;
}

pub trait GetReplace {
    fn get(&mut self, index: u32) -> Option<u32>;
}
