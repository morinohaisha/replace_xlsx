use std::fs::File;
use zip::ZipArchive;
use zip::ZipWriter;

pub struct XlsxReader {
    pub reader: ZipArchive<File>,
}

pub struct XlsxWriter {
    pub writer: ZipWriter<File>,
}
