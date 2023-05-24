pub mod _impls;
pub mod _structs;
pub mod _traits;

use crate::_structs::_xl::_drawings::_rels::drawing_rels::Relationships as DrawingRelationships;
use crate::_structs::_xl::_drawings::drawing::XdrWsDr;
use crate::_structs::_xl::_worksheets::_rels::sheet_rels::Relationships as SheetRelationships;
use crate::_structs::_xl::_worksheets::sheet::worksheet;
use crate::_structs::_xl::shared_strings;
use crate::_structs::content_types::Types;
use crate::_structs::input::Input;
use crate::_structs::input::Inputs;
use crate::_structs::replace::ReplaceXmls;
use crate::_structs::replace::Replaces;
use crate::_structs::xlsx_reader::XlsxReader;
use crate::_structs::xlsx_writer::XlsxWriter;
use crate::_traits::content_types::AddType;
use crate::_traits::replace::{IsSkip, Replace};
use crate::_traits::xlsx_reader::{XlsxArchive, XlsxGetFile};
use crate::_traits::xlsx_writer::XlsxWrite;
use crate::_traits::xlsx_writer::XmlReplace;
use _traits::input::Convert;
use _traits::replace::Extract;
use serde_json::from_str;
use std::io;
use zip::write::FileOptions;

pub fn exec_replace(template: &str, inputs: Inputs, output: &str) -> anyhow::Result<()> {
    let mut reader: XlsxReader = get_reader(template)?;
    let mut replaces = get_replaces(&inputs, &mut reader)?;

    let replace_xmls = get_replace_xmls(&mut replaces, &mut reader)?;

    let mut content_types: Types = Types::new(&mut reader)?;

    // Write
    let mut writer: XlsxWriter = XlsxWriter::new(output)?;

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(None)
        .unix_permissions(0o755);

    // 書き出し処理
    for i in 0..reader.len() {
        let file = reader.by_index(i)?;
        let _ = content_types.add_override(file.name());
        if file.name() == "[Content_Types].xml" {
            continue;
        } else if replace_xmls.is_skip(file.name()) {
            continue;
        } else {
            writer.raw_copy_file(file)?;
        }
    }

    for replace_xml in replace_xmls {
        writer.replace_file(&replace_xml.file_name, replace_xml.xml, options)?;
        let _ = content_types.add_override(&replace_xml.file_name);
    }

    let _ = write_images(&mut replaces, &mut content_types, &mut writer, options);

    // Content_Types.xml書き込み
    let _ = write_content_types(&mut replaces, &mut content_types, &mut writer, options);

    writer.finish()?;
    Ok(())
}

pub fn get_inputs() -> anyhow::Result<Inputs> {
    let mut buffer: String = String::new();
    io::stdin().read_line(&mut buffer)?;
    let input: String = buffer.replace("\n", "");
    let inputs: Inputs = from_str::<Inputs>(input.replace("\n", "").as_str())?;
    Ok(inputs)
}

fn get_reader(template: &str) -> anyhow::Result<XlsxReader> {
    let reader: XlsxReader = XlsxReader::new(template)?;
    Ok(reader)
}

fn get_replaces<'a>(inputs: &'a Inputs, reader: &mut XlsxReader) -> anyhow::Result<Replaces<'a>> {
    let mut replaces: Replaces = inputs.convert();
    replaces.extract(reader)?;
    Ok(replaces)
}

fn get_replace_xmls(
    replaces: &mut Replaces,
    reader: &mut XlsxReader,
) -> anyhow::Result<ReplaceXmls> {
    let mut replace_xmls: ReplaceXmls = Vec::new();

    let mut shared_strings: shared_strings::sst =
        shared_strings::sst::new(reader).expect("Faild to create shared strings");
    replace_xmls.push(
        shared_strings
            .replace(&replaces)
            .expect("Faild to replace shared strings"),
    );

    let mut drawing_relationships: DrawingRelationships =
        DrawingRelationships::new(1, reader).expect("Faild to create drawing relationships");
    replace_xmls.push(
        drawing_relationships
            .replace(&replaces)
            .expect("Faild to replace drawing relationships"),
    );

    let mut sheet_relationships: SheetRelationships =
        SheetRelationships::new(1, reader).expect("Faild to create sheet relationships");
    replace_xmls.push(
        sheet_relationships
            .replace(&replaces)
            .expect("Faild to replace sheet relationships"),
    );

    let mut xdr_ws_dr: XdrWsDr = XdrWsDr::new(1, reader).expect("Faild to create xdr_ws_dr");
    replace_xmls.push(
        xdr_ws_dr
            .replace(&replaces)
            .expect("Faild to replace xdr_ws_dr"),
    );

    let mut worksheet: worksheet = worksheet::new(1, reader).expect("Faild to create xdr_ws_dr");
    replace_xmls.push(
        worksheet
            .replace(&replaces)
            .expect("Faild to replace xdr_ws_dr"),
    );

    Ok(replace_xmls)
}

fn write_images(
    replaces: &mut Replaces,
    content_types: &mut Types,
    writer: &mut XlsxWriter,
    options: FileOptions,
) -> anyhow::Result<()> {
    for replace in replaces.iter() {
        match replace.input {
            Input::Text { from: _, to: _ } => (),
            Input::Image { from: _, to } => match &replace.image {
                None => (),
                Some(image) => {
                    let _ = content_types.add_default(&image.ext);
                    image.dist(to.to_string(), writer, options)?;
                }
            },
        }
    }
    Ok(())
}

fn write_content_types(
    replaces: &mut Replaces,
    content_types: &mut Types,
    writer: &mut XlsxWriter,
    options: FileOptions,
) -> anyhow::Result<()> {
    let content_types_xml = content_types
        .replace(replaces)
        .expect("Faild to replace content_types");
    writer
        .start_file(&content_types_xml.file_name, options)
        .unwrap();
    writer.write_all(&content_types_xml.xml).unwrap();
    Ok(())
}
