use std::collections::HashMap;
use std::fs;
use std::io::Cursor;
use std::io::prelude::*;
use std::io::BufReader;

use quick_xml::Writer;
use quick_xml::events::Event;
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
    let options = FileOptions::default()
        .compression_method(method)
        .unix_permissions(0o755);

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
    return text.contains("<w:sdt>");
}

#[derive(Debug)]
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
    fn parse_string(value: &String) -> Option<ContentControlType> {
        match value.as_str() {
            "w:richText" => Some(ContentControlType::RichText),
            "w:text" => Some(ContentControlType::Text),
            "w:comboBox" => Some(ContentControlType::ComboBox),
            "w:dropDownList" => Some(ContentControlType::DropdownList),
            "w:date" => Some(ContentControlType::Date),
            _ => None
        }
    }
}

struct ContentControl {
    tag: String,
    value: String,
    ct_type: ContentControlType,
    params: HashMap<String, String>,
}

impl ContentControl {
    fn new() -> ContentControl {
        ContentControl { tag: "".into(), value: "".into(), ct_type: ContentControlType::Unsupported, params: HashMap::new() }
    }
    fn infer_from_params(&mut self) {
        // if no type is set, RichText is default
        self.ct_type = ContentControlType::RichText;

        for (k, v) in &self.params {
            if k == "w:tag" {
                self.tag = v.to_string();
            } else if let Some(t) = ContentControlType::parse_string(&k) {
                self.ct_type = t;
            }
        }
    }
}

fn get_content_controls(data: &ZipData) -> Vec<ContentControl> {
    let mut controls: Vec<ContentControl> = Vec::new();
    for (name, string) in data {
        if has_content_control(&string) {
            let mut reader = Reader::from_str(string);
            let mut inPr = false;
            let mut inContent = false;
            let mut isControl = false;
            let mut control = ContentControl::new();
            loop {
                let ev = reader.read_event();
                match &ev {
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    Ok(quick_xml::events::Event::Eof) => break,
                    Ok(quick_xml::events::Event::Start(e)) => match e.name() {
                        QName(b"w:sdt") => {
                            for attr in e.attributes() {
                                println!("{:?}", attr);
                            }
                        }
                        QName(b"w:sdtPr") => {
                            inPr = true;
                        }
                        QName(b"w:sdtContent") => {
                            inContent = true;
                        }
                        QName(n) => {
                            if inPr {
                                control.params.insert(String::from_utf8_lossy(n).to_string(), "".into());
                            }
                        }
                    },
                    Ok(quick_xml::events::Event::Empty(e)) => {
                        if inPr {
                            let mut vwal: String = "".into();
                            for rattr in e.attributes() {
                                if let Ok(attr) = rattr {
                                if attr.key == QName(b"w:val") {
                                    vwal = String::from_utf8_lossy(&attr.value).into();
                                }
                                }
                            }
                            control.params.insert(String::from_utf8_lossy(e.name().into_inner()).to_string(), vwal);
                        }
                    }
                    Ok(quick_xml::events::Event::End(e)) => match e.name() {
                        QName(b"w:sdt") => {
                            control.infer_from_params();
                            println!("{} {:?}", control.tag, control.ct_type);
                            controls.push(control);
                            control = ContentControl::new();
                        }
                        QName(b"w:sdtPr") => {
                            inPr = false;
                        }
                        QName(b"w:sdtContent") => {
                            inContent = false;
                        }
                        _ => {}
                    },
                    _ => {}
                }
            }
        }
    }
    controls
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
