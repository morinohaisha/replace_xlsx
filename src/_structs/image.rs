use crate::_structs::zip::XlsxWriter;
use crate::_traits::zip::XlsxWrite;
use image::io::Reader;
use std::ffi::OsStr;
use std::fs::File;
use std::io::Read;
use std::path::Path;
use zip::write::FileOptions;

#[derive(Debug, PartialEq)]
pub struct Image {
    pub width: u32,
    pub height: u32,
    pub ext: String,
    pub name: Option<String>,
    pub to_name: Option<String>,
    pub dist_name: Option<String>,
    pub r_id: Option<String>,
}

impl Image {
    pub fn new(name: &str) -> Image {
        let ext = Path::new(&name)
            .extension()
            .and_then(OsStr::to_str)
            .unwrap();
        let reader = Reader::open(&name).expect("Image read failed");
        let image = reader.decode().expect("Image read decode failed");
        Image {
            width: image.width(),
            height: image.height(),
            ext: ext.to_string(),
            name: None,
            to_name: None,
            dist_name: None,
            r_id: None,
        }
    }

    pub fn set_index(&mut self, index: u32) {
        let rid = format!("rId{}", index);
        self.r_id = Some(rid);
        let name = format!("image{}.{}", index, self.ext);
        self.name = Some(name);
        let to_name = format!("../media/image{}.{}", index, self.ext);
        self.to_name = Some(to_name);
        let dist_name = format!("xl/media/image{}.{}", index, self.ext);
        self.dist_name = Some(dist_name);
    }

    pub fn dist(
        &self,
        to: String,
        writer: &mut XlsxWriter,
        options: FileOptions,
    ) -> anyhow::Result<()> {
        match &self.dist_name {
            None => (),
            Some(file_name) => {
                writer.start_file(&file_name, options)?;
                let mut file = File::open(to)?;
                let mut buf = Vec::new();
                let _ = file.read_to_end(&mut buf)?;
                writer.write_all(&buf)?;
            }
        };
        Ok(())
    }
}
