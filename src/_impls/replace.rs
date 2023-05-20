use crate::_structs::_xl::_drawings::_rels::drawing_rels;
use crate::_structs::_xl::_worksheets::sheet;
use crate::_structs::_xl::shared_strings;
use crate::_structs::image::Image;
use crate::_structs::input::Input;
use crate::_structs::replace::{Replace, Replaces, ReplaceXmls};
use crate::_structs::xlsx_reader::XlsxReader;
use crate::_traits::_xl::_worksheets::sheet::Worksheet;
use crate::_traits::_xl::shared_strings::Sst;
use crate::_traits::replace::Extract;
use crate::_traits::replace::GetReplace;
use crate::_traits::replace::IsSkip;

impl Replace<'_> {
    pub fn new(input: &Input) -> Replace {
        Replace {
            input,
            index: None,
            cells: Vec::new(),
            image: None,
        }
    }
}

impl Extract for Replaces<'_> {
    fn extract(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String> {
        self.extract_index(reader)?;
        self.extract_cells(reader)?;
        self.setup_images(reader)?;
        Ok("".to_string())
    }

    fn extract_index(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String> {
        let shared_strings: shared_strings::sst = shared_strings::sst::new(reader)?;
        for replace in self.iter_mut() {
            match replace.input {
                Input::Text { from, to: _ } => replace.index = shared_strings.index(from),
                Input::Image { from, to: _ } => replace.index = shared_strings.index(from),
            }
        }
        Ok("".to_string())
    }

    fn extract_cells(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String> {
        let sheet: sheet::worksheet = sheet::worksheet::new(1, reader)?;
        for replace in self.iter_mut() {
            match replace.index {
                None => (),
                Some(index) => replace.cells = sheet.cells(&index),
            }
        }
        Ok("".to_string())
    }

    fn setup_images(&mut self, reader: &mut XlsxReader) -> anyhow::Result<String> {
        let drawing_rels: drawing_rels::Relationships =
            drawing_rels::Relationships::new(1, reader)?;
        let mut index = drawing_rels
            .relationships
            .map_or(0, |relationship_lst| relationship_lst.len());
        for replace in self.iter_mut() {
            match replace.input {
                Input::Text { from: _, to: _ } => (),
                Input::Image { from: _, to } => {
                    index += 1;
                    let mut image = Image::new(to);
                    image.set_index(index as u32);
                    replace.image = Some(image);
                }
            }
        }
        Ok("".to_string())
    }
}

impl GetReplace for Replaces<'_> {
    fn get(&mut self, index: u32) -> Option<u32> {
        for replace in self.iter() {
            if replace.index == Some(index) {
                return replace.index;
            }
        }
        None
    }
}

impl IsSkip for ReplaceXmls {
    fn is_skip(&self, file_name: &str) -> bool {
        for replxml in self.iter() {
            if replxml.file_name.eq(file_name) {
                return true;
            }
        }
        false
    }
}
