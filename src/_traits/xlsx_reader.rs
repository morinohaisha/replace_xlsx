#[cfg(test)]
use mockall::automock;
use zip::read::ZipFile;

pub trait XlsxArchive: XlsxGetFile {
    // fn get_file(&mut self, file_name: &str, buf: &mut String) -> anyhow::Result<String>;
    // fn is_empty(&self) -> bool;
    // fn len(&self) -> usize;
    fn by_index(&mut self, file_number: usize) -> anyhow::Result<ZipFile>;
    fn by_name(&mut self, name: &str) -> anyhow::Result<ZipFile>;
}

#[cfg_attr(test, automock)]
pub trait XlsxGetFile {
    fn get_file(&mut self, file_name: &str, buf: &mut String) -> anyhow::Result<String>;
    fn is_empty(&self) -> bool;
    fn len(&self) -> usize;
}
