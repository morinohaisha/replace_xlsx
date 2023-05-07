use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type")]
pub enum Input {
    #[serde(rename = "text")]
    Text { from: String, to: String },
    #[serde(rename = "image")]
    Image {
        from: String,
        to: String,
        // #[serde(skip)]
        // index: Option<u32>,
        // #[serde(skip)]
        // cell: Option<Cell>,
        // #[serde(skip)]
        // image: Option<Image>,
    },
}

pub type Inputs = Vec<Input>;

#[cfg(test)]
mod tests {
    use crate::_structs::input::Input;
    use serde_json::from_str;

    const JSON: &str = r#"
[
    {"type":"text","from":"氏名","to":"もりのはいしゃ"},
    {"type":"image","from":"職業","to":"はいしゃ"}
]"#;

    #[test]
    fn deserialize() {
        assert!(from_str::<Vec<Input>>(JSON.replace("\n", "").as_str()).is_ok());
    }
}
