use crate::_structs::_xl::_drawings::drawing::{
    AAvLst, ABlip, AFillRect, APrstGeom, AStretch, XdrBlipFill, XdrCNvPicPr, XdrCNvPr,
    XdrClientData, XdrCol, XdrColOff, XdrExt, XdrFrom, XdrNvPicPr, XdrOneCellAnchor, XdrPic,
    XdrRow, XdrRowOff, XdrSpPr, XdrWsDr, XMLNS_A, XMLNS_C, XMLNS_CX, XMLNS_CX1, XMLNS_DGM,
    XMLNS_MC, XMLNS_R, XMLNS_SLE15, XMLNS_X3UNK, XMLNS_XDR,
};
use crate::_structs::input::Input;
use crate::_structs::replace::Replaces;
use crate::_structs::xml::XmlReader;
use crate::_structs::zip::XlsxReader;
use crate::_traits::replace::Replace;
use crate::_traits::zip::XlsxArchive;
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
        let file = format!("xl/drawings/drawing{}.xml", num);
        reader.get_file(&file, &mut buf)?;
        let mut drawing: XdrWsDr = XdrWsDr {
            xmlns_xdr: Some(XMLNS_XDR.to_string()),
            xmlns_a: Some(XMLNS_A.to_string()),
            xmlns_r: Some(XMLNS_R.to_string()),
            xmlns_c: Some(XMLNS_C.to_string()),
            xmlns_cx: Some(XMLNS_CX.to_string()),
            xmlns_cx1: Some(XMLNS_CX1.to_string()),
            xmlns_mc: Some(XMLNS_MC.to_string()),
            xmlns_dgm: Some(XMLNS_DGM.to_string()),
            xmlns_x3Unk: Some(XMLNS_X3UNK.to_string()),
            xmlns_sle15: Some(XMLNS_SLE15.to_string()),
            oneCellAnchor: None,
            xml: None,
        };
        if buf.len() == 0 {
            buf = to_string(&drawing)?;
        } else {
            drawing = from_str::<XdrWsDr>(buf.as_str())?;
        }
        drawing.xml = Some(buf);
        Ok(drawing)
    }

}

impl Replace for XdrWsDr {
    fn replace(&mut self, replaces: &Replaces) -> anyhow::Result<Vec<u8>> {
        let mut writer: Writer<Cursor<Vec<u8>>> = Writer::new(Cursor::new(Vec::<u8>::new()));
        let xml: String = self.xml.clone().context("xml is empty")?;
        let mut reader = XmlReader::new(&xml); // xml文字からreader生成
        reader.trim_text(true);
        let mut buf = Vec::new();

        fn replace_exec(replaces: &Replaces, writer: &mut Writer<Cursor<Vec<u8>>>) -> anyhow::Result<()> {
            for replace in replaces.iter() {
                match replace.input {
                    Input::Text { from: _, to: _ } => (),
                    Input::Image { from: _, to: _ } => match &replace.image {
                        None => (),
                        Some(image) => match &image.r_id {
                            None => (),
                            Some(id) => {
                                for cell in replace.cells.iter() {
                                    let image_name =
                                        image.name.clone().expect("Image name empty");
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
                                        BytesText::from_escaped(to_string(
                                            &xdr_one_cell_anchor,
                                        )?),
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
                    } else {
                        let _ = writer.write_event(Event::Empty(e.borrow()));
                    }
                },
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
        Ok(writer.into_inner().into_inner())
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
            pic: XdrPic {
                nvPicPr: XdrNvPicPr {
                    cNvPr: XdrCNvPr {
                        id: 0,
                        name: name,
                        title: "".to_string(),
                    },
                    cNvPicPr: XdrCNvPicPr {
                        preferRelativeResize: 0,
                    },
                },
                blipFill: XdrBlipFill {
                    blip: ABlip {
                        cstate: "".to_string(),
                        embed: embed,
                    },
                    stretch: AStretch {
                        fillRect: AFillRect {},
                    },
                },
                spPr: XdrSpPr {
                    prstGeom: APrstGeom {
                        prst: "".to_string(),
                        avLst: AAvLst {},
                    },
                    noFill: None,
                },
            },
            clientData: XdrClientData { fLocksWithSheet: 0 },
        }
    }
}
