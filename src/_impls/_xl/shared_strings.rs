use crate::_structs::_xl::shared_strings::{self, si};
use crate::_structs::input::Input;
use crate::_structs::replace::ReplaceXml;
use crate::_structs::replace::Replaces;
use crate::_structs::xml::XmlReader;
use crate::_traits::_xl::shared_strings::{Si, Sst, T};
use crate::_traits::replace::Replace;
use crate::_traits::xlsx_reader::XlsxArchive;
use anyhow::Context;
use quick_xml::de::from_str;
use quick_xml::events::{BytesText, Event};
use quick_xml::se::to_string;
use quick_xml::writer::Writer;
#[allow(unused_imports)]
use std::io::{BufWriter, Cursor, Write};

impl shared_strings::sst {
    pub fn new<R>(reader: &mut R) -> anyhow::Result<shared_strings::sst>
    where
        R: XlsxArchive,
    {
        let mut buf: String = String::new();
        reader.get_file(shared_strings::FILE_NAME, &mut buf)?;
        let mut sst: shared_strings::sst = from_str(buf.as_str())?;
        sst.xml = Some(buf);
        sst.file_name = shared_strings::FILE_NAME.to_string();
        Ok(sst)
    }
}

impl Replace for shared_strings::sst {
    fn replace(&mut self, replaces: &Replaces) -> anyhow::Result<ReplaceXml> {
        let mut writer = Writer::new(Cursor::new(Vec::<u8>::new()));
        let xml: String = self.xml.clone().context("xml is empty")?;
        let mut reader = XmlReader::new(&xml); // xml文字からreader生成
        reader.trim_text(true);
        let mut buf = Vec::new();
        let mut index: u32 = 0;
        let mut flg = true;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"si" {
                        for replace in replaces.iter() {
                            if replace.index == Some(index) {
                                match replace.input {
                                    Input::Image { from: _, to: _ } => (),
                                    Input::Text { from: _, to } => {
                                        flg = false;
                                        let si: si = si::new(to.to_string());
                                        let _ = writer.write_event(Event::Text(
                                            BytesText::from_escaped(to_string(&si)?),
                                        ));
                                    }
                                }
                            }
                        }
                        index += 1;
                    }
                    if flg {
                        let _ = writer.write_event(Event::Start(e));
                    }
                    ()
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"si" && !flg {
                        flg = true;
                    } else if flg {
                        let _ = writer.write_event(Event::End(e));
                    }
                    ()
                }
                Ok(Event::Eof) => break,
                Ok(e) => {
                    if flg {
                        let _ = writer.write_event(e);
                    }
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

impl Sst for shared_strings::sst {
    fn index(&self, target: &str) -> Option<u32> {
        match &self.si {
            None => None,
            Some(vec_si) => {
                for (i, si) in vec_si.iter().enumerate() {
                    match si.value() {
                        None => (),
                        Some(val) => {
                            if &val == target {
                                return Some(i as u32);
                            }
                        }
                    }
                }
                None
            }
        }
    }
}

impl Si for shared_strings::si {
    fn new(value: String) -> Self {
        shared_strings::si {
            t: Some(shared_strings::T::new(value)),
            r: None,
        }
    }

    fn value(&self) -> Option<String> {
        match &self.r {
            None => match &self.t {
                None => None,
                Some(t) => Some(t.t.to_string()),
            },
            Some(r) => {
                let str: String = r.iter().fold(
                    String::new(),
                    |a: String, b: &shared_strings::R| -> String { a + &b.t.t },
                );
                Some(str)
            }
        }
    }
}

impl T for shared_strings::T {
    fn new(t: String) -> shared_strings::T {
        shared_strings::T { t }
    }
}

#[cfg(test)]
mod tests {
    use crate::{
        _structs::_xl::shared_strings::{si, sst, R, T},
        _traits::_xl::shared_strings::{Si, Sst},
    };
    use quick_xml::de::from_str;

    const XML: &str = r#"
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="19" uniqueCount="11">
        <si>
            <t>pref_name</t>
        </si>
        <si>
            <t>sex</t>
        </si>
        <si>
            <t>population</t>
        </si>
        <si>
            <t>徳島</t>
        </si>
        <si>
            <t>香川</t>
        </si>
        <si>
            <t>愛媛</t>
        </si>
        <si>
            <t>高知</t>
        </si>
        <si>
            <t>福岡</t>
        </si>
        <si>
            <t>佐賀</t>
        </si>
        <si>
            <t>長崎</t>
        </si>
        <si>
            <t>東京</t>
        </si>
        <si>
            <t>#&lt;情報&gt;&lt;ほげ&gt;</t>
        </si>
        <si>
            <r>
                <rPr>
                    <sz val="10"/>
                    <color rgb="FF000000"/>
                    <rFont val="Arial"/>
                    <family val="0"/>
                    <charset val="1"/>
                </rPr>
                <t xml:space="preserve">#&lt;</t>
            </r>
            <r>
                <rPr>
                    <sz val="10"/>
                    <color rgb="FF000000"/>
                    <rFont val="Noto Sans CJK JP"/>
                    <family val="2"/>
                </rPr>
                <t xml:space="preserve">情報</t>
            </r>
            <r>
                <rPr>
                    <sz val="10"/>
                    <color rgb="FF000000"/>
                    <rFont val="Arial"/>
                    <family val="0"/>
                    <charset val="1"/>
                </rPr>
                <t xml:space="preserve">&gt;&lt;</t>
            </r>
            <r>
                <rPr>
                    <sz val="10"/>
                    <color rgb="FF000000"/>
                    <rFont val="Noto Sans CJK JP"/>
                    <family val="2"/>
                </rPr>
                <t xml:space="preserve">氏名</t>
            </r>
            <r>
                <rPr>
                    <sz val="10"/>
                    <color rgb="FF000000"/>
                    <rFont val="Arial"/>
                    <family val="0"/>
                    <charset val="1"/>
                </rPr>
                <t xml:space="preserve">&gt;</t>
            </r>
        </si>
    </sst>
"#;

    #[test]
    fn index() {
        let sst: sst = from_str::<sst>(XML).unwrap();
        assert_eq!(sst.index("#<情報><ほげ>"), Some(11));
        assert_eq!(sst.index("#<情報><ふが>"), None);
    }

    #[test]
    fn si_t_value() {
        let si = si {
            t: Some(T {
                t: "DUMMY".to_string(),
            }),
            r: None,
        };
        assert_eq!("DUMMY", si.value().unwrap());
    }
    #[test]
    fn si_r_value() {
        let r: R = R {
            t: T {
                t: "DUMMY".to_string(),
            },
        };
        let si = si {
            t: None,
            r: Some(vec![r]),
        };
        assert_eq!("DUMMY", si.value().unwrap());
    }
}
