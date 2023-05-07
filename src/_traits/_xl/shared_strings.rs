pub trait Sst {
    fn index(&self, target: &str) -> Option<u32>;
    // fn indexes(&self, replaces: &mut Replaces) -> anyhow::Result<String>;
}

pub trait Si {
    fn new(value: String) -> Self;
    fn value(&self) -> Option<String>;
}

pub trait T {
    fn new(t: String) -> Self;
}
