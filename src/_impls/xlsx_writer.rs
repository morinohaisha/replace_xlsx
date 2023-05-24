use crate::{
    _structs::xlsx_writer::XlsxWriter,
    _traits::xlsx_writer::{XlsxWrite, XmlReplace},
};
use std::io::Write;
use zip::{read::ZipFile, write::FileOptions};

impl XlsxWriter {
    pub fn new(out_file: &str) -> anyhow::Result<XlsxWriter> {
        let outputzipfile = std::fs::File::create(std::path::Path::new(&out_file))?;
        let writer = zip::ZipWriter::new(outputzipfile);
        Ok(XlsxWriter { writer })
    }
}

impl XlsxWrite for XlsxWriter {
    fn start_file(&mut self, file_name: &str, options: FileOptions) -> anyhow::Result<()> {
        self.writer.start_file(file_name, options)?;
        Ok(())
    }

    fn add_directory(&mut self, file_name: &str, options: FileOptions) -> anyhow::Result<()> {
        self.writer.add_directory(file_name, options)?;
        Ok(())
    }

    fn write_all(&mut self, buf: &[u8]) -> anyhow::Result<()> {
        self.writer.write_all(buf)?;
        Ok(())
    }

    fn raw_copy_file(&mut self, file: ZipFile) -> anyhow::Result<()> {
        self.writer.raw_copy_file(file)?;
        Ok(())
    }

    fn finish(&mut self) -> anyhow::Result<()> {
        self.writer.finish()?;
        Ok(())
    }
}

impl XmlReplace for XlsxWriter {
    fn replace_file(
        &mut self,
        file_name: &str,
        xml: Vec<u8>,
        options: FileOptions,
    ) -> anyhow::Result<()> {
        self.start_file(file_name, options)?;
        self.write_all(xml.as_slice())?;
        Ok(())
    }
}
