use regex::Regex;

#[derive(Debug, PartialEq)]
pub struct Cell {
    pub col: u32,
    pub row: u32,
}

impl Cell {
    pub fn new(str: &str) -> Self {
        let re = Regex::new(r"^(?P<col>[A-Z]+)(?P<row>\d+)$").unwrap();
        let cap = re.captures(str).unwrap();
        Cell {
            col: Self::col(cap.name("col").unwrap().as_str()),
            row: Self::row(cap.name("row").unwrap().as_str()),
        }
    }
    pub fn col(str: &str) -> u32 {
        let lst = str
            .chars()
            .map(|a| -> u32 { a.to_digit(36).unwrap() - 9 })
            .rev();
        lst.enumerate()
            .map(|(i, x)| -> u32 {
                let digit = i as u32;
                (digit * 26) - digit + x
            })
            .sum::<u32>()
            - 1
    }

    pub fn row(str: &str) -> u32 {
        let row = str.parse::<u32>().unwrap();
        row - 1
    }
}
