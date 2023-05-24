use serde::{Deserialize, Serialize};

pub const FILE_NAME: &'static str = "xl/drawings/_rels/drawing{}.xml.rels";

pub const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
pub const RELATION_TYPE: &str =
    r#"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"#;
pub const RELATION_SHIPS_XMLNS: &str =
    r#"http://schemas.openxmlformats.org/package/2006/relationships"#;

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_snake_case)]
pub struct Relationships {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,
    #[serde(rename(serialize = "Relationship", deserialize = "Relationship"))]
    pub relationships: Option<Vec<Relationship>>,
    #[serde(skip)]
    pub xml: Option<String>,
    #[serde(skip)]
    pub file_name: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
pub struct Relationship {
    #[serde(rename = "@Id")]
    pub id: String,
    #[serde(rename = "@Type")]
    pub relation_type: String,
    #[serde(rename = "@Target")]
    pub target: String,
}

#[cfg(test)]
mod tests {
    use crate::_structs::_xl::_drawings::_rels::drawing_rels::Relationships;
    use quick_xml::de::from_str;

    const XML: &str = r#"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpg"/>
</Relationships>
"#;

    #[test]
    fn deserialize() {
        assert!(from_str::<Relationships>(XML).is_ok());
    }
}
