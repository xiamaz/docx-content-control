use std::collections::HashMap;
use std::fs;
use std::io::prelude::*;
use std::io::BufReader;
use std::io::Cursor;

use quick_xml::events::BytesText;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Writer;
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
            _ => None,
        }
    }
}

struct ContentControl {
    tag: String,
    value: String,
    ct_type: ContentControlType,
    params: HashMap<String, String>,
    contains_paragraph: bool,
    content_begin: i64,
    content_end: i64,
}

impl ContentControl {
    fn new() -> ContentControl {
        ContentControl {
            tag: "".into(),
            value: "".into(),
            ct_type: ContentControlType::Unsupported,
            params: HashMap::new(),
            contains_paragraph: false,
            content_begin: -1,
            content_end: -1,
        }
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

fn write_content<'a, W>(control: &'a ContentControl, writer: &'a mut Writer<W>, content: &'a str) -> Result<(), &'a str> where W: std::io::Write {
    if control.contains_paragraph {
            let _ = writer.create_element("w:p").write_inner_content(|writer| {
                let _ = writer.create_element("w:r").write_inner_content(|writer| {
                    let _ = writer
                        .create_element("w:t")
                        .write_text_content(BytesText::new(content));
                    Ok(())
                });
                Ok(())
            });
    } else {
            let _ = writer.create_element("w:r").write_inner_content(|writer| {
                let _ = writer
                    .create_element("w:t")
                    .write_text_content(BytesText::new(content));
                Ok(())
            });
    }
    Ok(())
}

struct DocumentData<'a> {
    events: Vec<Event<'a>>,
    controls: Vec<ContentControl>,
}

impl<'a> DocumentData<'a> {
    fn new() -> DocumentData<'a> {
        DocumentData {
            events: Vec::new(),
            controls: Vec::new(),
        }
    }
}

type ParsedDocuments<'a> = HashMap<String, DocumentData<'a>>;

struct DocumentState {
    states: HashMap<String, i32>,
    is_eof: bool,
}

impl DocumentState {
    fn new() -> DocumentState {
        DocumentState {
            states: HashMap::new(),
            is_eof: false,
        }
    }

    fn is_in(&self, key: &str) -> bool {
        self.states.get(key).unwrap_or(&0) > &0
    }

    fn consume<'b>(&mut self, event: &'b Event) {
        match event {
            Event::Start(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.states.insert(name, current + 1);
            }
            Event::End(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.states.insert(name, current - 1);
            }
            Event::Eof => self.is_eof = true,
            _ => {}
        }
    }
}

fn get_content_controls(data: &ZipData) -> ParsedDocuments {
    let mut documents = HashMap::new();
    for (filename, string) in data {
        if has_content_control(&string) {
            let mut reader = Reader::from_str(string);
            let mut state = DocumentState::new();
            let mut controls: Vec<ContentControl> = Vec::new();
            let mut control = ContentControl::new();
            let mut events: Vec<Event> = Vec::new();
            while !state.is_eof {
                let event = reader.read_event();
                match event {
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    Ok(e) => {
                        state.consume(&e);
                        match &e {
                            Event::Start(v) => {
                                if state.is_in("w:sdtPr") {
                                    control.params.insert(
                                        String::from_utf8_lossy(v.name().into_inner()).to_string(),
                                        "".into(),
                                    );
                                }
                                if v.name() == QName(b"w:sdtContent") {
                                    control.content_begin = events.len() as i64;
                                }
                                if state.is_in("w:sdtContent") && v.name() == QName(b"w:p") {
                                    control.contains_paragraph = true;
                                }
                            }
                            Event::End(v) => {
                                if v.name() == QName(b"w:sdt") {
                                    control.infer_from_params();
                                    controls.push(control);
                                    control = ContentControl::new();
                                } else if v.name() == QName(b"w:sdtContent") {
                                    control.content_end = events.len() as i64;
                                }
                            }
                            Event::Empty(v) => {
                                if state.is_in("w:sdtPr") {
                                    let mut vwal: String = "".into();
                                    for rattr in v.attributes() {
                                        if let Ok(attr) = rattr {
                                            if attr.key == QName(b"w:val") {
                                                vwal = String::from_utf8_lossy(&attr.value).into();
                                            }
                                        }
                                    }
                                    control.params.insert(
                                        String::from_utf8_lossy(v.name().into_inner()).to_string(),
                                        vwal,
                                    );
                                }
                            }
                            Event::Text(v) => {
                                if state.is_in("w:t") && state.is_in("w:sdtContent") {
                                    control.value.push_str(&v.unescape().unwrap().to_string());
                                }
                            }
                            _ => {}
                        }
                        events.push(e.clone());
                    }
                }
            }
            documents.insert(filename.into(), DocumentData { events, controls });
        }
    }
    documents
}

fn clear_content_controls(data: &ZipData, controlled: &ParsedDocuments) -> ZipData {
    let mut mapped_data = ZipData::new();
    for (filename, data) in data {
        if let Some(doc) = controlled.get(filename) {
            for control in &doc.controls {
                println!("{} {:?} {}", control.tag, control.ct_type, control.value);
            }
            let events: Vec<Event> = doc
                .events
                .iter()
                .enumerate()
                .filter(|&(i, _)| {
                    !doc.controls
                        .iter()
                        .map(|c| (i as i64) > c.content_begin && (i as i64) < c.content_end)
                        .reduce(|a, b| a || b)
                        .unwrap_or(true)
                })
                .map(|(_, v)| v.clone())
                .collect();
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            for event in events {
                let _ = writer.write_event(event);
            }
            let new_data = String::from_utf8(writer.into_inner().into_inner()).unwrap();
            mapped_data.insert(filename.into(), new_data);
        } else {
            mapped_data.insert(filename.into(), data.into());
        }
    }
    mapped_data
}

fn get_intersecting_control(index: i64, controls: &Vec<ContentControl>) -> Option<&ContentControl> {
    for control in controls {
        if index >= control.content_begin && index < control.content_end {
            return Some(control);
        }
    }
    return None;
}

fn map_content_controls(
    data: &ZipData,
    controlled: &ParsedDocuments,
    mappings: &HashMap<&str, &str>,
) -> ZipData {
    let mut mapped_data = ZipData::new();
    for (filename, data) in data {
        if let Some(doc) = controlled.get(filename) {
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            for (i, event) in doc.events.iter().enumerate() {
                if let Some(control) = get_intersecting_control(i as i64, &doc.controls) {
                    if (i as i64) == control.content_begin {
                        let _ = writer.write_event(event);
                        let mapped = mappings.get(&control.tag.as_str()).unwrap_or(&"");
                        write_content(&control, &mut writer, *mapped);
                        // let _ = writer.create_element("w:r").write_inner_content(|writer| {
                        //     let _ = writer
                        //         .create_element("w:t")
                        //         .write_text_content(BytesText::new(mapped));
                        //     Ok(())
                        // });
                    }
                } else {
                    let _ = writer.write_event(event);
                }
            }
            let new_data = String::from_utf8(writer.into_inner().into_inner()).unwrap();
            mapped_data.insert(filename.into(), new_data);
        } else {
            mapped_data.insert(filename.into(), data.into());
        }
    }
    mapped_data
}

fn main() {
    let fname = std::path::Path::new("tests/data/content_controlled_document.docx");
    let file = fs::File::open(fname).unwrap();
    let reader = BufReader::new(file);
    let mappings = HashMap::from([
        ("Title", "Brave New World"),
        ("Sidematter", "Into a brave new world"),
        ("WritingDate", "12.12.2012"),
        ("Author", "Bruce Wayne"),
        ("MainContent", "This is rich coming from you."),
    ]);
    if let Ok(data) = list_zip_contents(reader) {
        let controlled_documents = get_content_controls(&data);
        // let new_data = clear_content_controls(&data, &controlled_documents);
        let new_data = map_content_controls(&data, &controlled_documents, &mappings);
        let _ = zip_dir(&new_data, "test.docx", zip::CompressionMethod::Deflated);
    }
}
