#![allow(non_snake_case)]
use serde::{Deserialize, Serialize};

pub const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
pub const XMLNS_XDR: &str =
    r#"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"#;
pub const XMLNS_A: &str = r#"http://schemas.openxmlformats.org/drawingml/2006/main"#;
pub const XMLNS_R: &str = r#"http://schemas.openxmlformats.org/officeDocument/2006/relationships"#;

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:wsDr")]
pub struct XdrWsDr {
    #[serde(rename = "@xmlns:xdr")]
    pub xmlns_xdr: Option<String>,
    #[serde(rename = "@xmlns:a")]
    pub xmlns_a: Option<String>,
    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: Option<String>,
    #[serde(rename(serialize = "xdr:oneCellAnchor"))]
    pub oneCellAnchor: Option<Vec<XdrOneCellAnchor>>,
    #[serde(skip)]
    pub xml: Option<String>,
    #[serde(skip)]
    pub file_name: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:oneCellAnchor")]
pub struct XdrOneCellAnchor {
    #[serde(rename(serialize = "xdr:from"))]
    pub from: XdrFrom,
    #[serde(rename(serialize = "xdr:ext"))]
    pub ext: XdrExt,
    #[serde(rename(serialize = "xdr:pic"))]
    pub pic: XdrPic,
    #[serde(rename(serialize = "xdr:clientData"))]
    pub clientData: XdrClientData,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:from")]
pub struct XdrFrom {
    #[serde(rename(serialize = "xdr:col"))]
    pub col: XdrCol,
    #[serde(rename(serialize = "xdr:colOff"))]
    pub colOff: XdrColOff,
    #[serde(rename(serialize = "xdr:row"))]
    pub row: XdrRow,
    #[serde(rename(serialize = "xdr:rowOff"))]
    pub rowOff: XdrRowOff,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:col")]
pub struct XdrCol {
    #[serde(rename = "$text")]
    pub value: u32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:colOff")]
pub struct XdrColOff {
    #[serde(rename = "$text")]
    pub value: i32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:row")]
pub struct XdrRow {
    #[serde(rename = "$text")]
    pub value: u32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:rowOff")]
pub struct XdrRowOff {
    #[serde(rename = "$text")]
    pub value: i32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:ext")]
pub struct XdrExt {
    #[serde(rename = "@cx")]
    pub cx: u32,
    #[serde(rename = "@cy")]
    pub cy: u32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:pic")]
pub struct XdrPic {
    #[serde(rename(serialize = "xdr:nvPicPr"))]
    pub nvPicPr: XdrNvPicPr,
    #[serde(rename(serialize = "xdr:blipFill"))]
    pub blipFill: XdrBlipFill,
    #[serde(rename(serialize = "xdr:spPr"))]
    pub spPr: XdrSpPr,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:nvPicPr")]
pub struct XdrNvPicPr {
    #[serde(rename(serialize = "xdr:cNvPr"))]
    pub cNvPr: XdrCNvPr,
    #[serde(rename(serialize = "xdr:cNvPicPr"))]
    pub cNvPicPr: XdrCNvPicPr,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:cNvPr")]
pub struct XdrCNvPr {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@descr")]
    pub descr: String,
    #[serde(rename = "$text")]
    pub v: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:cNvPicPr")]
pub struct XdrCNvPicPr {
    #[serde(
        rename = "@preferRelativeResize",
        skip_serializing_if = "Option::is_none"
    )]
    pub preferRelativeResize: Option<u32>,
    //
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:blipFill")]
pub struct XdrBlipFill {
    #[serde(rename(serialize = "a:blip"))]
    pub blip: ABlip,
    #[serde(rename(serialize = "a:stretch"))]
    pub stretch: AStretch,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:blip")]
pub struct ABlip {
    #[serde(rename = "@cstate", skip_serializing_if = "Option::is_none")]
    pub cstate: Option<String>,
    #[serde(rename(deserialize = "@embed", serialize = "@r:embed"))]
    pub embed: String,
    #[serde(rename = "$text")]
    pub v: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:stretch")]
pub struct AStretch {
    #[serde(rename(serialize = "a:fillRect"))]
    pub fillRect: Option<AFillRect>,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:fillRect")]
pub struct AFillRect {}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:spPr")]
pub struct XdrSpPr {
    #[serde(rename(serialize = "a:xfrm"))]
    pub axfrm: Option<AXfrm>,
    #[serde(rename(serialize = "a:prstGeom"))]
    pub prstGeom: APrstGeom,
    #[serde(rename(serialize = "a:ln"))]
    pub a_ln: Option<ALn>,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:ln")]
pub struct ALn {
    #[serde(rename = "@w")]
    pub w: u32,
    #[serde(rename(serialize = "a:noFill"))]
    pub noFill: Option<ANoFill>,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:xfrm")]
pub struct AXfrm {
    #[serde(rename(serialize = "a:off"))]
    pub aoff: AOff,
    #[serde(rename(serialize = "a:ext"))]
    pub aext: AExt,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:off")]
pub struct AOff {
    #[serde(rename = "@x")]
    pub x: u32,
    #[serde(rename = "@y")]
    pub y: u32,
}
#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:ext")]
pub struct AExt {
    #[serde(rename = "@cx")]
    pub cx: u32,
    #[serde(rename = "@cy")]
    pub cy: u32,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:prstGeom")]
pub struct APrstGeom {
    #[serde(rename = "@prst")]
    pub prst: String,
    #[serde(rename(serialize = "a:avLst"))]
    pub avLst: AAvLst,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:avLst")]
pub struct AAvLst {}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "a:noFill")]
pub struct ANoFill {}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "xdr:clientData")]
pub struct XdrClientData {
    #[serde(rename = "@fLocksWithSheet", skip_serializing_if = "Option::is_none")]
    pub fLocksWithSheet: Option<usize>,
}

#[cfg(test)]
mod tests {
    use crate::_structs::_xl::_drawings::drawing::XdrWsDr;
    use quick_xml::de::from_str;

    const XML: &str = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <xdr:oneCellAnchor>
        <xdr:from>
            <xdr:col>1</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>20</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:ext cx="987" cy="1181"/>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id="0" name="image" title="画像"/>
                <xdr:cNvPicPr preferRelativeResize="0"/>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip cstate="print" r:embed="embed"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln w="0">
                    <a:noFill/>
                </a:ln>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData fLocksWithSheet="0"/>
    </xdr:oneCellAnchor>
</xdr:wsDr>
"#;

    #[test]
    fn deserialize() {
        assert!(from_str::<XdrWsDr>(XML).is_ok());
    }
}
