use crate::{
    _structs::zip::{XlsxReader, XlsxWriter},
    _traits::zip::{XlsxArchive, XlsxWrite},
};
use std::{
    fs::File,
    io::{Read, Write},
    path::Path,
};
use zip::{read::ZipFile, write::FileOptions, ZipArchive};

impl XlsxReader {
    pub fn new(template: &str) -> anyhow::Result<XlsxReader> {
        let path = Path::new(template);
        let zipfile = File::open(path)?;
        let reader = ZipArchive::new(zipfile)?;
        Ok(XlsxReader { reader })
    }
}

impl XlsxArchive for XlsxReader {
    fn get_file(&mut self, file_name: &str, buf: &mut String) -> anyhow::Result<String> {
        for i in 0..self.reader.len() {
            let mut file: ZipFile = self.reader.by_index(i)?;
            let name: &str = file.name();
            if name == file_name {
                let _ = file.read_to_string(buf);
                break;
            }
        }
        Ok("ok".to_string())
    }
    fn is_empty(&self) -> bool {
        self.reader.len() == 0
    }

    fn len(&self) -> usize {
        self.reader.len()
    }

    fn by_index(&mut self, file_number: usize) -> anyhow::Result<ZipFile> {
        Ok(self.reader.by_index(file_number)?)
    }

    fn by_name(&mut self, name: &str) -> anyhow::Result<ZipFile> {
        Ok(self.reader.by_name(name)?)
    }
}

impl XlsxWriter {
    pub fn new(out_file: &'static str) -> anyhow::Result<XlsxWriter> {
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
