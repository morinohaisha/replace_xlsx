use zip::{read::ZipFile, write::FileOptions};

pub trait XlsxWrite {
    fn start_file(&mut self, file_name: &str, options: FileOptions) -> anyhow::Result<()>;
    fn write_all(&mut self, buf: &[u8]) -> anyhow::Result<()>;
    fn raw_copy_file(&mut self, file: ZipFile) -> anyhow::Result<()>;
    fn add_directory(&mut self, file_name: &str, options: FileOptions) -> anyhow::Result<()>;
    fn finish(&mut self) -> anyhow::Result<()>;
}

pub trait XmlReplace {
    fn replace_file(&mut self,  file_name: &str, xml: Vec<u8>, options: FileOptions) -> anyhow::Result<()>;
}
