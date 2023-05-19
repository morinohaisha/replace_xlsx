pub trait AddType {
    fn add_default(&mut self, extension: &str) -> anyhow::Result<()>;
    fn add_override(&mut self, path_name: &str) -> anyhow::Result<()>;
}

pub trait Contains {
    fn contains(&self, key: &str) -> bool;
}
