use zip::{read::ZipFile, write::FileOptions};

pub trait XlsxArchive {
    fn get_file(&mut self, file_name: &str, buf: &mut String) -> anyhow::Result<String>;
    fn is_empty(&self) -> bool;
    fn len(&self) -> usize;
    fn by_index(&mut self, file_number: usize) -> anyhow::Result<ZipFile>;
    fn by_name(&mut self, name: &str) -> anyhow::Result<ZipFile>;
}

pub trait XlsxWrite {
    fn start_file(&mut self, file_name: &str, options: FileOptions) -> anyhow::Result<()>;
    fn write_all(&mut self, buf: &[u8]) -> anyhow::Result<()>;
    fn raw_copy_file(&mut self, file: ZipFile) -> anyhow::Result<()>;
    fn finish(&mut self) -> anyhow::Result<()>;
}
