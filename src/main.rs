use clap::Parser;
use replace_xlsx::_structs::_xl::_drawings::_rels::drawing_rels::Relationships as DrawingRelationships;
use replace_xlsx::_structs::_xl::_drawings::drawing::XdrWsDr;
use replace_xlsx::_structs::_xl::_worksheets::_rels::sheet_rels::Relationships as SheetRelationships;
use replace_xlsx::_structs::_xl::_worksheets::sheet::worksheet;
use replace_xlsx::_structs::_xl::shared_strings;
use replace_xlsx::_structs::content_types::Types;
use replace_xlsx::_structs::input::{Input, Inputs};
use replace_xlsx::_structs::replace::{Replaces, ReplaceXmls};
use replace_xlsx::_structs::xlsx_reader::XlsxReader;
use replace_xlsx::_structs::xlsx_writer::XlsxWriter;
use replace_xlsx::_traits::content_types::AddType;
use replace_xlsx::_traits::input::Convert;
use replace_xlsx::_traits::replace::{Extract, Replace, IsSkip};
use replace_xlsx::_traits::xlsx_reader::XlsxArchive;
use replace_xlsx::_traits::xlsx_writer::XmlReplace;
use replace_xlsx::_traits::xlsx_writer::XlsxWrite;
use serde_json::from_str;
use std::io;
use zip::write::FileOptions;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long)]
    template: String,

    #[arg(short, long, default_value = "out.xlsx")]
    out: String,
}

fn main() {

    let args: Args = Args::parse();
    let template: &str = args.template.as_str();
    let output: &str = args.out.as_str();
    let inputs = get_inputs().expect("Faild to input json error");

    let mut reader: XlsxReader = XlsxReader::new(template).expect("Faild to Xlsx reader error");

    let mut replaces: Replaces = inputs.convert();
    replaces
        .extract(&mut reader)
        .expect("Faild to extract error");

    let mut replace_xmls: ReplaceXmls = Vec::new();

    let mut shared_strings: shared_strings::sst =
        shared_strings::sst::new(&mut reader).expect("Faild to create shared strings");
    replace_xmls.push(shared_strings
        .replace(&replaces)
        .expect("Faild to replace shared strings"));

    let mut drawing_relationships: DrawingRelationships =
        DrawingRelationships::new(1, &mut reader).expect("Faild to create drawing relationships");
    replace_xmls.push(drawing_relationships
        .replace(&replaces)
        .expect("Faild to replace drawing relationships"));

    let mut sheet_relationships: SheetRelationships =
        SheetRelationships::new(1, &mut reader).expect("Faild to create sheet relationships");
    replace_xmls.push(sheet_relationships
        .replace(&replaces)
        .expect("Faild to replace sheet relationships"));

    let mut xdr_ws_dr: XdrWsDr = XdrWsDr::new(1, &mut reader).expect("Faild to create xdr_ws_dr");
    replace_xmls.push(xdr_ws_dr
        .replace(&replaces)
        .expect("Faild to replace xdr_ws_dr"));

    let mut worksheet: worksheet =
        worksheet::new(1, &mut reader).expect("Faild to create xdr_ws_dr");
    replace_xmls.push(worksheet
        .replace(&replaces)
        .expect("Faild to replace xdr_ws_dr"));

    let mut content_types: Types = Types::new(&mut reader).expect("Faild to create content_types");

    // Write
    let mut writer: XlsxWriter = XlsxWriter::new(output).expect("XlsxWriter error");

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(None)
        .unix_permissions(0o755);

    // 書き出し処理
    for i in 0..reader.len() {
        let file = reader.by_index(i).unwrap();
        let _ = content_types.add_override(file.name());
        if file.name() == "[Content_Types].xml" {
            continue;
        } else if replace_xmls.is_skip(file.name()) {
            continue;
        } else {
            writer.raw_copy_file(file).unwrap();
        }
    }

    for replace_xml in replace_xmls {
        writer.replace_file(&replace_xml.file_name, replace_xml.xml, options).unwrap();
        let _ = content_types.add_override(&replace_xml.file_name);
    }

    for replace in replaces.iter() {
        match replace.input {
            Input::Text { from: _, to: _ } => (),
            Input::Image { from: _, to } => match &replace.image {
                None => (),
                Some(image) => {
                    let _ = content_types.add_default(&image.ext);
                    image
                        .dist(to.to_string(), &mut writer, options)
                        .expect("image dist error");
                }
            },
        }
    }

    // Content_Types.xml書き込み
    let content_types_xml = content_types
        .replace(&replaces)
        .expect("Faild to replace content_types");
    writer.start_file(&content_types_xml.file_name, options).unwrap();
    writer.write_all(&content_types_xml.xml).unwrap();

    writer.finish().unwrap();
}

fn get_inputs() -> anyhow::Result<Inputs> {
    let mut buffer: String = String::new();
    io::stdin().read_line(&mut buffer)?;
    let input: String = buffer.replace("\n", "");
    let inputs: Inputs = from_str::<Inputs>(input.replace("\n", "").as_str())?;
    Ok(inputs)
}
