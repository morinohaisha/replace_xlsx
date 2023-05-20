use crate::_structs::_xl::_drawings::drawing::{
    AAvLst, ABlip, AExt, ALn, AOff, APrstGeom, AStretch, AXfrm, XdrBlipFill, XdrCNvPicPr, XdrCNvPr,
    XdrClientData, XdrCol, XdrColOff, XdrExt, XdrFrom, XdrNvPicPr, XdrOneCellAnchor, XdrPic,
    XdrRow, XdrRowOff, XdrSpPr, XdrWsDr, XMLNS_A, XMLNS_R, XMLNS_XDR, XML_DECLARATION,
};
use crate::_structs::input::Input;
use crate::_structs::replace::{Replaces, ReplaceXml};
use crate::_structs::xml::XmlReader;
use crate::_structs::xlsx_reader::XlsxReader;
use crate::_traits::replace::Replace;
use crate::_traits::xlsx_reader::XlsxArchive;
use anyhow::Context;
use quick_xml::de::from_str;
use quick_xml::events::{BytesText, Event};
use quick_xml::se::to_string;
use quick_xml::writer::Writer;
#[allow(unused_imports)]
use std::io::{BufWriter, Cursor, Write};

impl XdrWsDr {
    pub fn new(num: u32, reader: &mut XlsxReader) -> anyhow::Result<XdrWsDr> {
        let mut buf: String = String::new();
        let file_name = format!("xl/drawings/drawing{}.xml", num);
        reader.get_file(&file_name, &mut buf)?;
        let mut drawing: XdrWsDr = XdrWsDr {
            xmlns_xdr: Some(XMLNS_XDR.to_string()),
            xmlns_a: Some(XMLNS_A.to_string()),
            xmlns_r: Some(XMLNS_R.to_string()),
            oneCellAnchor: None,
            xml: None,
            file_name,
        };
        if buf.len() == 0 {
            buf = to_string(&drawing)?;
        } else {
            drawing = from_str::<XdrWsDr>(buf.as_str())?;
        }
        drawing.xml = Some(format!("{}{}", XML_DECLARATION, buf));
        Ok(drawing)
    }
}

impl Replace for XdrWsDr {
    fn replace(&mut self, replaces: &Replaces) -> anyhow::Result<ReplaceXml> {
        let mut writer: Writer<Cursor<Vec<u8>>> = Writer::new(Cursor::new(Vec::<u8>::new()));
        let xml: String = self.xml.clone().context("xml is empty")?;
        let mut reader = XmlReader::new(&xml); // xml文字からreader生成
        reader.trim_text(true);
        let mut buf = Vec::new();

        fn replace_exec(
            replaces: &Replaces,
            writer: &mut Writer<Cursor<Vec<u8>>>,
        ) -> anyhow::Result<()> {
            for replace in replaces.iter() {
                match replace.input {
                    Input::Text { from: _, to: _ } => (),
                    Input::Image { from: _, to: _ } => match &replace.image {
                        None => (),
                        Some(image) => match &image.r_id {
                            None => (),
                            Some(id) => {
                                for cell in replace.cells.iter() {
                                    let image_name = image.name.clone().expect("Image name empty");
                                    let emu = 914400;
                                    let dpi = 72;
                                    let mut ratio = 1.0;
                                    let mut w_ratio = 0.997;
                                    if image.width > 300 {
                                        ratio = 0.608;
                                        w_ratio = 1.0;
                                    }

                                    let wi = ((emu / dpi) as f64 * ratio * w_ratio).round() as u32;
                                    let hi = ((emu / dpi) as f64 * ratio).round() as u32;
                                    // let i = 12708;
                                    let cx = wi * image.width;
                                    let cy = hi * image.height;
                                    let xdr_one_cell_anchor = XdrOneCellAnchor::new(
                                        cell.col,
                                        0,
                                        cell.row,
                                        0,
                                        cx,
                                        cy,
                                        image_name,
                                        id.to_string(),
                                    );
                                    let _ = writer.write_event(Event::Text(
                                        BytesText::from_escaped(to_string(&xdr_one_cell_anchor)?),
                                    ));
                                }
                            }
                        },
                    },
                }
            }
            Ok(())
        }

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"xdr:wsDr" {
                        let _ = writer.write_event(Event::Start(e.borrow()));
                        let _ = replace_exec(replaces, &mut writer)?;
                        let _ = writer.write_event(Event::End(e.to_end()));
                    } else if e.name().as_ref() == b"xdr:oneCellAnchor" {
                        ()
                    } else {
                        let _ = writer.write_event(Event::Empty(e.borrow()));
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"xdr:wsDr" {
                        let _ = replace_exec(replaces, &mut writer)?;
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

impl XdrOneCellAnchor {
    pub fn new(
        col: u32,
        col_off: i32,
        row: u32,
        row_off: i32,
        cx: u32,
        cy: u32,
        name: String,
        embed: String,
    ) -> XdrOneCellAnchor {
        XdrOneCellAnchor {
            from: XdrFrom {
                col: XdrCol { value: col },
                colOff: XdrColOff { value: col_off },
                row: XdrRow { value: row },
                rowOff: XdrRowOff { value: row_off },
            },
            ext: XdrExt { cx, cy },
            clientData: XdrClientData {
                fLocksWithSheet: None,
            },
            pic: XdrPic {
                nvPicPr: XdrNvPicPr {
                    cNvPr: XdrCNvPr {
                        id: 0,
                        name: name,
                        descr: "".to_string(),
                        v: "".to_string(),
                    },
                    cNvPicPr: XdrCNvPicPr::default(),
                },
                blipFill: XdrBlipFill {
                    blip: ABlip {
                        cstate: None,
                        embed: embed,
                        v: "".to_string(),
                    },
                    stretch: AStretch { fillRect: None },
                },
                spPr: XdrSpPr {
                    axfrm: Some(AXfrm {
                        aoff: AOff { x: 0, y: 0 },
                        aext: AExt { cx, cy },
                    }),
                    prstGeom: APrstGeom {
                        prst: "rect".to_string(),
                        avLst: AAvLst {},
                    },
                    a_ln: Some(ALn { w: 0, noFill: None }),
                },
            },
        }
    }
}
