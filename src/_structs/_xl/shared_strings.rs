use serde::{Deserialize, Serialize};

pub const FILE_NAME: &'static str = "xl/sharedStrings.xml";

/// `<t>`
#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
pub struct T {
    #[serde(rename = "$text")]
    pub t: String,
}

/// `<si>`
#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_camel_case_types)]
pub struct si {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub t: Option<T>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r: Option<Vec<R>>,
}

/// `<r>`
#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
pub struct R {
    pub t: T,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_camel_case_types)]
pub struct sst {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "@count")]
    count: Option<u32>,
    #[serde(rename = "@uniqueCount")]
    unique_count: Option<u32>,

    pub si: Option<Vec<si>>,
    #[serde(skip)]
    pub xml: Option<String>,
    #[serde(skip)]
    pub file_name: String,
}

#[cfg(test)]
mod tests {
    use crate::_structs::_xl::shared_strings::sst;
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
    fn deserialize() {
        assert!(from_str::<sst>(XML).is_ok());
    }
}
