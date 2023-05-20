use std::fs::File;
use zip::ZipArchive;

pub struct XlsxReader {
    pub reader: ZipArchive<File>,
}
