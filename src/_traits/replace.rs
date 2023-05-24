use crate::_structs::replace::{ReplaceXml, Replaces};

use crate::_traits::xlsx_reader::XlsxArchive;

pub trait Extract {
    fn extract<R>(&mut self, reader: &mut R) -> anyhow::Result<String>
    where
        R: XlsxArchive;
    fn extract_index<R>(&mut self, reader: &mut R) -> anyhow::Result<String>
    where
        R: XlsxArchive;
    fn extract_cells<R>(&mut self, reader: &mut R) -> anyhow::Result<String>
    where
        R: XlsxArchive;
    fn setup_images<R>(&mut self, reader: &mut R) -> anyhow::Result<String>
    where
        R: XlsxArchive;
}

pub trait Replace {
    fn replace(&mut self, replaces: &Replaces) -> anyhow::Result<ReplaceXml>;
}

pub trait GetReplace {
    fn get(&mut self, index: u32) -> Option<u32>;
}

pub trait IsSkip {
    fn is_skip(&self, file_name: &str) -> bool;
}
