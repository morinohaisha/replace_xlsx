use clap::Parser;
use replace_xlsx::_structs::_xl::_drawings::_rels::drawing_rels::Relationships as DrawingRelationships;
use replace_xlsx::_structs::_xl::_drawings::drawing::XdrWsDr;
use replace_xlsx::_structs::_xl::_worksheets::_rels::sheet_rels::Relationships as SheetRelationships;
use replace_xlsx::_structs::_xl::_worksheets::sheet::worksheet;
use replace_xlsx::_structs::_xl::shared_strings;
use replace_xlsx::_structs::input::{Input, Inputs};
use replace_xlsx::_structs::replace::Replaces;
use replace_xlsx::_structs::zip::XlsxReader;
use replace_xlsx::_structs::zip::XlsxWriter;
use replace_xlsx::_traits::input::Convert;
use replace_xlsx::_traits::replace::Extract;
use replace_xlsx::_traits::replace::Replace;
use replace_xlsx::_traits::zip::XlsxArchive;
use replace_xlsx::_traits::zip::XlsxWrite;
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
    let inputs = get_inputs().expect("Faild to input json error");
    let mut replaces: Replaces = inputs.convert();
    let args: Args = Args::parse();
    let template: &str = args.template.as_str();

    let mut reader: XlsxReader = XlsxReader::new(template).expect("Faild to Xlsx reader error");
    replaces
        .extract(&mut reader)
        .expect("Faild to extract error");

    let mut shared_strings: shared_strings::sst =
        shared_strings::sst::new(&mut reader).expect("Faild to create shared strings");
    let shared_strings_xml = shared_strings
        .replace(&replaces)
        .expect("Faild to replace shared strings");

    let mut drawing_relationships: DrawingRelationships =
        DrawingRelationships::new(1, &mut reader).expect("Faild to create drawing relationships");
    let drawing_relationships_xml = drawing_relationships
        .replace(&replaces)
        .expect("Faild to replace drawing relationships");

    let mut sheet_relationships: SheetRelationships =
        SheetRelationships::new(1, &mut reader).expect("Faild to create sheet relationships");
    let sheet_relationships_xml = sheet_relationships
        .replace(&replaces)
        .expect("Faild to replace sheet relationships");

    let mut xdr_ws_dr: XdrWsDr = XdrWsDr::new(1, &mut reader).expect("Faild to create xdr_ws_dr");
    let xdr_ws_dr_xml = xdr_ws_dr
        .replace(&replaces)
        .expect("Faild to replace xdr_ws_dr");

    let mut worksheet: worksheet =
        worksheet::new(1, &mut reader).expect("Faild to create xdr_ws_dr");
    let worksheet_xml = worksheet
        .replace(&replaces)
        .expect("Faild to replace xdr_ws_dr");

    // Write
    let mut writer: XlsxWriter = XlsxWriter::new("out.xlsx").expect("XlsxWriter error");

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    // 書き出し処理
    for i in 0..reader.len() {
        let file = reader.by_index(i).unwrap();
        writer.start_file(file.name(), options).unwrap();
        if file.name() == "xl/sharedStrings.xml" {
            writer.write_all(shared_strings_xml.as_slice()).unwrap();
        } else if file.name() == "xl/drawings/_rels/drawing1.xml.rels" {
            writer
                .write_all(drawing_relationships_xml.as_slice())
                .unwrap();
        } else if file.name() == "xl/drawings/drawing1.xml" {
            writer.write_all(xdr_ws_dr_xml.as_slice()).unwrap();
        } else if file.name() == "xl/worksheets/sheet1.xml" {
            writer.write_all(worksheet_xml.as_slice()).unwrap();
        } else {
            writer.raw_copy_file(file).unwrap();
        }
    }

    if reader.by_name("xl/drawings/drawing1.xml").is_err() {
        writer
            .start_file("xl/drawings/drawing1.xml", options)
            .unwrap();
        writer.write_all(xdr_ws_dr_xml.as_slice()).unwrap();
    }
    if reader
        .by_name("xl/drawings/_rels/drawing1.xml.rels")
        .is_err()
    {
        writer
            .start_file("xl/drawings/_rels/drawing1.xml.rels", options)
            .unwrap();
        writer
            .write_all(drawing_relationships_xml.as_slice())
            .unwrap();
    }
    if reader
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")
        .is_err()
    {
        writer
            .start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        writer
            .write_all(sheet_relationships_xml.as_slice())
            .unwrap();
    }

    for replace in replaces.iter() {
        match replace.input {
            Input::Text { from: _, to: _ } => (),
            Input::Image { from: _, to } => match &replace.image {
                None => (),
                Some(image) => image
                    .dist(to.to_string(), &mut writer, options)
                    .expect("image dist error"),
            },
        }
    }
    writer.finish().unwrap();
}

fn get_inputs() -> anyhow::Result<Inputs> {
    let mut buffer: String = String::new();
    io::stdin().read_line(&mut buffer)?;
    let input: String = buffer.replace("\n", "");
    let inputs: Inputs = from_str::<Inputs>(input.replace("\n", "").as_str())?;
    Ok(inputs)
}
