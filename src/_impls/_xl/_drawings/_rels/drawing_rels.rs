use crate::_structs::_xl::_drawings::_rels::drawing_rels::{
    Relationship, Relationships, RELATION_SHIPS_XMLNS, RELATION_TYPE, XML_DECLARATION,
};
use crate::_structs::input::Input;
use crate::_structs::replace::ReplaceXml;
use crate::_structs::xml::XmlReader;
use crate::_traits::replace::Replace;
use crate::_traits::xlsx_reader::XlsxArchive;
use anyhow::Context;
use quick_xml::de::from_str;
use quick_xml::events::{BytesText, Event};
use quick_xml::se::to_string;
use quick_xml::writer::Writer;
#[allow(unused_imports)]
use std::io::{BufWriter, Cursor, Write};

impl Relationships {
    pub fn new<R>(num: u32, reader: &mut R) -> anyhow::Result<Relationships>
    where
        R: XlsxArchive,
    {
        let mut buf: String = String::new();
        let file_name = format!("xl/drawings/_rels/drawing{}.xml.rels", num);
        reader.get_file(&file_name, &mut buf)?;
        let mut drawing_rels: Relationships = Relationships {
            xmlns: RELATION_SHIPS_XMLNS.to_string(),
            relationships: None,
            xml: None,
            file_name,
        };
        if buf.len() == 0 {
            buf = to_string(&drawing_rels)?;
        } else {
            drawing_rels = from_str::<Relationships>(buf.as_str())?;
        }
        drawing_rels.xml = Some(format!("{}{}", XML_DECLARATION, buf));
        Ok(drawing_rels)
    }
}

impl Relationship {
    pub fn new(id: String, target: String) -> Relationship {
        Relationship {
            id,
            target,
            relation_type: RELATION_TYPE.to_string(),
        }
    }
}

impl Replace for Relationships {
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
                    if e.name().as_ref() == b"Relationships" {
                        for replace in replaces.iter() {
                            match replace.input {
                                Input::Text { from: _, to: _ } => (),
                                Input::Image { from: _, to: _ } => match &replace.image {
                                    None => (),
                                    Some(image) => {
                                        let id = image.r_id.clone().context("r_id is empty")?;
                                        let target =
                                            image.to_name.clone().context("to_name is empty")?;
                                        let relationship = Relationship::new(id, target);
                                        let _ = writer.write_event(Event::Text(
                                            BytesText::from_escaped(to_string(&relationship)?),
                                        ));
                                    }
                                },
                            }
                        }
                        let _ = writer.write_event(Event::End(e.clone()));
                    } else {
                        let _ = writer.write_event(Event::End(e.clone()));
                    }
                }
                Ok(Event::Empty(_)) => (),
                Ok(Event::Eof) => break,
                Ok(e) => {
                    let _ = writer.write_event(e);
                }
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            }
            buf.clear();
        }
        Ok(ReplaceXml {
            file_name: self.file_name.clone(),
            xml: writer.into_inner().into_inner(),
        })
    }
}
