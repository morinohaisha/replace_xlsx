use crate::_structs::_xl::_worksheets::sheet;
use crate::_structs::_xl::_worksheets::sheet::{drawing, C};
use crate::_structs::cell::Cell;
use crate::_structs::input::Input;
use crate::_structs::xml::XmlReader;
use crate::_structs::xlsx_reader::XlsxReader;
use crate::_structs::replace::ReplaceXml;
use crate::_traits::_xl::_worksheets::sheet::Row;
use crate::_traits::_xl::_worksheets::sheet::Worksheet;
use crate::_traits::replace::Replace;
use crate::_traits::xlsx_reader::XlsxArchive;
use anyhow::Context;
use quick_xml::de::from_str;
use quick_xml::events::{BytesText, Event};
use quick_xml::se::to_string;
use quick_xml::writer::Writer;
#[allow(unused_imports)]
use std::io::{BufWriter, Cursor, Write};

impl sheet::worksheet {
    pub fn new(num: u32, reader: &mut XlsxReader) -> anyhow::Result<sheet::worksheet> {
        let mut buf: String = String::new();
        let file_name = format!("xl/worksheets/sheet{}.xml", num);
        reader.get_file(&file_name, &mut buf)?;
        let mut worksheet: sheet::worksheet = from_str(buf.as_str())?;
        worksheet.xml = Some(buf);
        worksheet.drawing = Some(vec![drawing {
            r_id: "rId1".to_string(),
        }]);
        worksheet.file_name = file_name;
        Ok(worksheet)
    }
}

impl Worksheet for sheet::worksheet {
    fn cells(&self, index: &u32) -> Vec<Cell> {
        self.sheetData
            .row
            .iter()
            .map(|r| r.refs(index))
            .into_iter()
            .flatten()
            .collect()
    }
}

impl Row for sheet::Row {
    fn string_cells(&self) -> Vec<&C> {
        match &self.c {
            Some(c) => c
                .iter()
                .filter(|c| match &c.t {
                    None => false,
                    Some(t) => t == "s",
                })
                .collect(),
            None => vec![],
        }
    }

    fn refs(&self, index: &u32) -> Vec<Cell> {
        let mut cells = Vec::new();
        for cell in self.string_cells().iter() {
            match &cell.v {
                None => (),
                Some(v) => {
                    if v.v == index.to_string() {
                        cells.push(Cell::new(&cell.r));
                    }
                }
            }
        }
        cells
    }
}

impl Replace for sheet::worksheet {
    fn replace(
        &mut self,
        replaces: &crate::_structs::replace::Replaces,
    ) -> anyhow::Result<ReplaceXml> {
        let mut writer = Writer::new(Cursor::new(Vec::<u8>::new()));
        let xml: String = self.xml.clone().context("xml is empty")?;
        let mut reader = XmlReader::new(&xml); // xml文字からreader生成
        reader.trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"worksheet" {
                        if replaces.iter().any(|replace| match replace.input {
                            Input::Image { from: _, to: _ } => true,
                            _ => false,
                        }) {
                            let drawing = sheet::drawing {
                                r_id: "rId1".to_string(),
                            };
                            let _ = writer.write_event(Event::Text(BytesText::from_escaped(
                                to_string(&drawing)?,
                            )));
                        }
                        let _ = writer.write_event(Event::End(e.clone()));
                    } else {
                        let _ = writer.write_event(Event::End(e.clone()));
                    }
                }
                Ok(Event::Eof) => break,
                Ok(e) => {
                    let _ = writer.write_event(e);
                }
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            }
            buf.clear();
        }
        Ok(ReplaceXml {file_name: self.file_name.clone(), xml: writer.into_inner().into_inner()})
    }
}
