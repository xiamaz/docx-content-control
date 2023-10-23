use std::collections::HashMap;
use std::fs;
use std::io::BufReader;
use std::io::prelude::*;

use quick_xml::name::QName;
use zip::write::FileOptions;

use quick_xml::reader::Reader;

type ZipData = HashMap<String, String>;

fn list_zip_contents(reader: impl Read + Seek) -> zip::result::ZipResult<ZipData> {
    let mut zip = zip::ZipArchive::new(reader)?;

    let mut data: ZipData = HashMap::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i)?;
        let mut data_str = String::new();
        let _ = file.read_to_string(&mut data_str);
        // std::io::copy(&mut file, &mut std::io::stdout());
        data.insert(file.name().into(), data_str);
    }

    Ok(data)
}

fn zip_dir(
    data: &HashMap<String, String>,
    path_str: &str,
    method: zip::CompressionMethod,
) -> zip::result::ZipResult<()> {
    let path = std::path::Path::new(path_str);
    let file = fs::File::create(path).unwrap();
    let mut writer = zip::ZipWriter::new(file);
    let options = FileOptions::default().compression_method(method).unix_permissions(0o755);

    for (key, value) in data {
        writer.start_file(key, options)?;
        let _ = writer.write_all(value.as_bytes());
    }
    Ok(())
}

/**
 * Check if the string contains an sdt tag (Ruby Inline-Level Structured Document Tag)
 */
fn has_content_control(text: &String) -> bool {
    return text.contains("<w:sdt>")
}

enum ContentControlType {
    Unsupported,
    RichText,
    Text,
    ComboBox,
    DropdownList,
    Date,
}

enum FontStyleSpecifier {
    Normal,
    Bold,
    Italic,
    Superscript,
    Subscript,
}

struct FontFormatting {
    size: i32,
    style: FontStyleSpecifier,
}

impl ContentControlType {
    fn from_value(val: i32) -> ContentControlType {
        match val {
            0 => ContentControlType::RichText,
            1 => ContentControlType::Text,
            3 => ContentControlType::ComboBox,
            4 => ContentControlType::DropdownList,
            6 => ContentControlType::Date,
            _ => ContentControlType::Unsupported,
        }
    }
}

fn get_content_controls(data: &ZipData) {
    for (name, string) in data {
        if has_content_control(&string) {
            let mut reader = Reader::from_str(string);
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    Ok(quick_xml::events::Event::Eof) => break,
                    Ok(quick_xml::events::Event::Start(e)) => {
                        if e.name() == QName(b"w:sdt") {
                            // TODO: match sdtPr and sdtContent
                            // - edit sdtContent if we have a matching entry
                            println!("{:?}", e.name());
                        }
                    }
                    Ok(e) => {},
                }
            }
            println!("{}", name);
        }
    }
}

fn main() {
    let fname = std::path::Path::new("tests/data/content_controlled_document.docx");
    let file = fs::File::open(fname).unwrap();
    let reader = BufReader::new(file);
    if let Ok(data) = list_zip_contents(reader) {
        get_content_controls(&data);
        let _ = zip_dir(&data, "test.docx", zip::CompressionMethod::Deflated);
    }
}
