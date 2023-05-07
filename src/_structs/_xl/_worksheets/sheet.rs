use serde::{Deserialize, Serialize};

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_camel_case_types, non_snake_case)]
pub struct worksheet {
    #[serde(rename = "sheetData")]
    pub sheetData: SheetData,
    pub drawing: Option<Vec<drawing>>,
    #[serde(skip)]
    pub xml: Option<String>,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
pub struct SheetData {
    pub row: Vec<Row>,
}

//drawing
#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[allow(non_camel_case_types)]
pub struct drawing {
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    pub r_id: String,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
pub struct Row {
    #[serde(rename = "@r")]
    pub r: String,
    pub c: Option<Vec<C>>,
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct C {
    #[serde(rename = "@r")]
    pub r: String,
    #[serde(rename = "@s")]
    s: String,
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
    pub v: Option<V>,
}

#[derive(Debug, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "v")]
pub struct V {
    #[serde(rename = "$text")]
    pub v: String,
}

#[cfg(test)]
mod tests {
    use crate::_structs::_xl::_worksheets::sheet::worksheet;
    use quick_xml::de::from_str;
    const XML: &str = r#"
<worksheet>
    <sheetData>
        <row r="1">
            <c r="A1" s="1" t="s">
                <v>0</v>
            </c>
            <c r="B1" s="1" t="s">
                <v>1</v>
            </c>
            <c r="C1" s="1" t="s">
                <v>2</v>
            </c>
        </row>
        <row r="2">
            <c r="A2" s="1" t="s">
                <v>3</v>
            </c>
            <c r="B2" s="1">
                <v>1.0</v>
            </c>
            <c r="C2" s="1">
                <v>60.0</v>
            </c>
        </row>
        <row r="3">
            <c r="A3" s="1" t="s">
                <v>3</v>
            </c>
            <c r="B3" s="1">
                <v>2.0</v>
            </c>
            <c r="C3" s="1">
                <v>40.0</v>
            </c>
        </row>
        <row r="4">
            <c r="A4" s="1" t="s">
                <v>4</v>
            </c>
            <c r="B4" s="1">
                <v>1.0</v>
            </c>
            <c r="C4" s="1">
                <v>100.0</v>
            </c>
        </row>
        <row r="5">
            <c r="A5" s="1" t="s">
                <v>4</v>
            </c>
            <c r="B5" s="1">
                <v>2.0</v>
            </c>
            <c r="C5" s="1">
                <v>100.0</v>
            </c>
        </row>
        <row r="6">
            <c r="A6" s="1" t="s">
                <v>5</v>
            </c>
            <c r="B6" s="1">
                <v>1.0</v>
            </c>
            <c r="C6" s="1">
                <v>100.0</v>
            </c>
        </row>
        <row r="7">
            <c r="A7" s="1" t="s">
                <v>5</v>
            </c>
            <c r="B7" s="1">
                <v>2.0</v>
            </c>
            <c r="C7" s="1">
                <v>50.0</v>
            </c>
        </row>
        <row r="8">
            <c r="A8" s="1" t="s">
                <v>6</v>
            </c>
            <c r="B8" s="1">
                <v>1.0</v>
            </c>
            <c r="C8" s="1">
                <v>100.0</v>
            </c>
        </row>
    </sheetData>
    <drawing r:id="rId1"/>
</worksheet>
"#;
    #[test]
    fn deserialize() {
        assert!(from_str::<worksheet>(XML).is_ok());
    }
}
