use serde::{Deserialize, Serialize};

pub const CONTENT_TYPES: &str = "ContentTypes.toml";

#[derive(Debug, Default, PartialEq, Deserialize)]
#[allow(non_snake_case)]
pub struct ContentTypes {
    pub Default: toml::Table,
    pub Override: toml::Table,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename(serialize = "Types"))]
#[allow(non_snake_case)]
pub struct Types {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,
    #[serde(rename(serialize = "Default", deserialize = "Default"))]
    pub defaults: Defaults,
    #[serde(rename(serialize = "Override", deserialize = "Override"))]
    pub overrides: Overrides,
    #[serde(skip)]
    pub xml: Option<String>,
    #[serde(skip)]
    pub content_types: ContentTypes,
}

pub type Defaults = Vec<Default>;
pub type Overrides = Vec<Override>;

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_snake_case)]
pub struct Default {
    #[serde(rename = "@Extension")]
    pub extension: String,
    #[serde(rename = "@ContentType")]
    pub content_type: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_snake_case)]
pub struct Override {
    #[serde(rename = "@PartName")]
    pub part_name: String,
    #[serde(rename = "@ContentType")]
    pub content_type: String,
}

pub enum DefaultTypes {
    Jpeg,
    Png,
    Xml,
    Rels,
}

#[cfg(test)]
mod tests {
    use crate::_structs::content_types::Types;
    use quick_xml::de::from_str;

    const XML: &str = r#"
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    </Types>
"#;

    #[test]
    fn deserialize() {
        assert!(from_str::<Types>(XML).is_ok());
        let types: Types = from_str::<Types>(XML).unwrap();
        assert_eq!(2, types.defaults.len());
        assert_eq!(7, types.overrides.len());
    }
}
