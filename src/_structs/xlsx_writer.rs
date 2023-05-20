use std::fs::File;
use zip::ZipWriter;

pub struct XlsxWriter {
    pub writer: ZipWriter<File>,
}
