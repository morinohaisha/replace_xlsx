use crate::_structs::input::Inputs;
use crate::_structs::replace::Replace;
use crate::_structs::replace::Replaces;
use crate::_traits::input::Convert;

impl Convert for Inputs {
    fn convert(&self) -> Replaces {
        let mut replaces = Vec::<Replace>::new();
        for input in self.into_iter() {
            replaces.push(Replace::new(input));
        }
        replaces
    }
}

#[cfg(test)]
mod tests {
    use crate::_structs::input::{Input, Inputs};
    use crate::_structs::replace::Replaces;
    use crate::_traits::input::Convert;
    use serde_json::from_str;

    const JSON: &str = r#"
[
    {"type":"text","from":"氏名","to":"もりのはいしゃ"},
    {"type":"image","from":"職業","to":"はいしゃ"}
]"#;

    #[test]
    fn test_convert() {
        let inputs: Inputs = from_str::<Vec<Input>>(JSON.replace("\n", "").as_str()).unwrap();
        let replaces: Replaces = inputs.convert();
        assert_eq!(replaces.len(), 2);
        for replace in replaces.iter() {
            match replace.input {
                Input::Text { from, to } => {
                    assert_eq!(from, "氏名");
                    assert_eq!(to, "もりのはいしゃ");
                }
                Input::Image { from, to } => {
                    assert_eq!(from, "職業");
                    assert_eq!(to, "はいしゃ");
                }
            }
        }
    }
}
