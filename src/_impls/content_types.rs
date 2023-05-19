use crate::_structs::content_types::{
    ContentTypes, Default, Defaults, Override, Overrides, Types, CONTENT_TYPES,
};
use crate::_structs::replace::Replaces;
use crate::_structs::xml::XmlReader;
use crate::_structs::zip::XlsxReader;
use crate::_traits::content_types::{AddType, Contains};
use crate::_traits::replace::Replace;
use crate::_traits::zip::XlsxArchive;
use anyhow::Context;
use quick_xml::de::from_str;
use quick_xml::events::{BytesText, Event};
use quick_xml::se::to_string;
use quick_xml::writer::Writer;
use std::fs;
#[allow(unused_imports)]
use std::io::{BufWriter, Cursor, Write};

impl Types {
    pub fn new(reader: &mut XlsxReader) -> anyhow::Result<Types> {
        let mut buf: String = String::new();
        reader.get_file("[Content_Types].xml", &mut buf)?;
        let mut types: Types = from_str(buf.as_str())?;
        types.xml = Some(buf);
        types.set_content_types()?;
        types.reset_defaults()?;
        Ok(types)
    }

    fn set_content_types(&mut self) -> anyhow::Result<()> {
        let toml = fs::read_to_string(CONTENT_TYPES)?;
        let content_types: ContentTypes = toml::from_str::<ContentTypes>(&toml)?;
        self.content_types = content_types;
        Ok(())
    }

    fn reset_defaults(&mut self) -> anyhow::Result<()> {
        let lst = vec!["rels", "xml", "jpeg"];
        self.defaults = Vec::<Default>::new();
        for ext in lst.iter() {
            self.add_default(ext)?;
        }
        Ok(())
    }
}

impl Contains for Defaults {
    fn contains(&self, key: &str) -> bool {
        for default in self.iter() {
            if default.extension.eq(&key.to_string()) {
                return true;
            }
        }
        false
    }
}

impl Contains for Overrides {
    fn contains(&self, key: &str) -> bool {
        for default in self.iter() {
            if default.part_name.eq(&key.to_string()) {
                return true;
            }
        }
        false
    }
}

impl AddType for Types {
    fn add_default(&mut self, extension: &str) -> anyhow::Result<()> {
        if self.defaults.contains(extension) {
            return Ok(());
        }
        if self.content_types.Default.contains_key(extension) {
            let content_type_value: &toml::Value = self
                .content_types
                .Default
                .get(extension)
                .expect("not found");
            let content_type = match content_type_value {
                toml::Value::String(v) => v.to_string(),
                _ => "".to_string(),
            };
            self.defaults.push(Default {
                extension: extension.to_string(),
                content_type: content_type,
            });
        }
        Ok(())
    }

    fn add_override(&mut self, path_name: &str) -> anyhow::Result<()> {
        let path = format!("/{}", path_name);
        if self.overrides.contains(&path) {
            return Ok(());
        }
        if self.content_types.Override.contains_key(&path) {
            let content_type_value: &toml::Value =
                self.content_types.Override.get(&path).expect("not found");
            let content_type = match content_type_value {
                toml::Value::String(v) => v.to_string(),
                _ => "".to_string(),
            };
            self.overrides.push(Override {
                part_name: path,
                content_type: content_type,
            });
        }
        Ok(())
    }
}

impl Replace for Types {
    fn replace(&mut self, _replaces: &Replaces) -> anyhow::Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::<u8>::new()));
        let xml: String = self.xml.clone().context("xml is empty")?;
        let mut reader = XmlReader::new(&xml); // xml文字からreader生成
        reader.trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"Types" {
                        let _ = writer.write_event(Event::Start(e));
                        for item in self.defaults.iter() {
                            let _ = writer.write_event(Event::Text(BytesText::from_escaped(
                                to_string(item)?,
                            )));
                        }
                        for item in self.overrides.iter() {
                            let _ = writer.write_event(Event::Text(BytesText::from_escaped(
                                to_string(item)?,
                            )));
                        }
                    }
                    ()
                }
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() != b"Default" && e.name().as_ref() != b"Override" {
                        let _ = writer.write_event(Event::Empty(e));
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
        Ok(writer.into_inner().into_inner())
    }
}

// #[cfg(test)]
// mod tests {
//     #[test]
//     fn test_replace() {}
// }
