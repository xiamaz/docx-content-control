use std::collections::HashMap;
use std::error::Error;
use std::fmt;
use std::str;
use std::io::prelude::*;
use std::io::Cursor;

use quick_xml::events::BytesText;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Writer;
use zip::write::FileOptions;

use quick_xml::reader::Reader;

pub type ZipData = HashMap<String, Vec<u8>>;

#[derive(Debug)]
struct ParserError {
    data: String,
}

impl fmt::Display for ParserError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Parser error, probably malformed xml tags")
    }
}

impl Error for ParserError {}

pub fn list_zip_contents(reader: impl Read + Seek) -> zip::result::ZipResult<ZipData> {
    let mut zip = zip::ZipArchive::new(reader)?;

    let mut data: ZipData = HashMap::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i)?;
        let mut buf = Vec::new();
        let _ = file.read_to_end(&mut buf);
        // std::io::copy(&mut file, &mut std::io::stdout());
        data.insert(file.name().into(), buf);
    }

    Ok(data)
}

pub fn zip_dir<W: Write + Seek>(
    data: &HashMap<String, Vec<u8>>,
    file: &mut W,
) -> zip::result::ZipResult<()> {
    let mut writer = zip::ZipWriter::new(file);
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o755);

    for (key, value) in data {
        writer.start_file(key, options)?;
        let _ = writer.write_all(value);
    }
    Ok(())
}

fn find_subsequence(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|window| window == needle)
}

/**
 * Check if the string contains an sdt tag (Ruby Inline-Level Structured Document Tag)
 */
fn has_content_control(text: &Vec<u8>) -> bool {
    find_subsequence(text, b"<w:sdt>").is_some()
}

#[derive(Debug)]
pub enum ContentControlType {
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
    pub fn parse_string(value: &String) -> Option<ContentControlType> {
        match value.as_str() {
            "w:richText" => Some(ContentControlType::RichText),
            "w:text" => Some(ContentControlType::Text),
            "w:comboBox" => Some(ContentControlType::ComboBox),
            "w:dropDownList" => Some(ContentControlType::DropdownList),
            "w:date" => Some(ContentControlType::Date),
            _ => None,
        }
    }

    pub fn to_string(&self) -> String {
        match &self {
            ContentControlType::RichText => "w:richText".to_string(),
            ContentControlType::Text => "w:text".to_string(),
            ContentControlType::ComboBox => "w:comboBox".to_string(),
            ContentControlType::DropdownList => "w:dropDownList".to_string(),
            ContentControlType::Date => "w:date".to_string(),
            ContentControlType::Unsupported => "unsupported".to_string(),
        }
    }
}

pub struct ContentControl<'a> {
    pub tag: String,
    value: String,
    ct_type: ContentControlType,
    params: HashMap<String, String>,
    contains_paragraph: bool,
    paragraph_params: Vec<Event<'a>>,
    run_params: Vec<Event<'a>>,
    content_begin: i64,
    content_end: i64,
}

impl<'a> ContentControl<'a> {
    fn new() -> ContentControl<'a> {
        ContentControl {
            tag: "".into(),
            value: "".into(),
            ct_type: ContentControlType::Unsupported,
            params: HashMap::new(),
            paragraph_params: Vec::new(),
            run_params: Vec::new(),
            contains_paragraph: false,
            content_begin: -1,
            content_end: -1,
        }
    }

    pub fn get_control_type(&self) -> &ContentControlType {
        &self.ct_type
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

fn write_content<'a, W>(
    control: &'a ContentControl,
    writer: &'a mut Writer<W>,
    content: &'a str,
) -> Result<(), &'a str>
where
    W: std::io::Write,
{
    if control.contains_paragraph {
        let _ = writer.create_element("w:p").write_inner_content(|writer| {
            for ev in &control.paragraph_params {
                let _ = writer.write_event(ev);
            }
            let _ = writer.create_element("w:r").write_inner_content(|writer| {
                for ev in &control.run_params {
                    let _ = writer.write_event(ev);
                }
                let _ = writer
                    .create_element("w:t")
                    .write_text_content(BytesText::new(content));
                Ok(())
            });
            Ok(())
        });
    } else {
        let _ = writer.create_element("w:r").write_inner_content(|writer| {
            for ev in &control.run_params {
                let _ = writer.write_event(ev);
            }
            let _ = writer
                .create_element("w:t")
                .write_text_content(BytesText::new(content));
            Ok(())
        });
    }
    Ok(())
}

pub struct DocumentData<'a, 'b> {
    events: Vec<Event<'a>>,
    pub controls: Vec<ContentControl<'b>>,
}

type ParsedDocuments<'a, 'b> = HashMap<String, DocumentData<'a, 'b>>;

pub struct DocumentState {
    states: HashMap<String, i32>,
    is_eof: bool,
    last_seen_closed: String,
}

impl DocumentState {
    fn new() -> DocumentState {
        DocumentState {
            states: HashMap::new(),
            is_eof: false,
            last_seen_closed: "".into(),
        }
    }

    fn is_in(&self, key: &str) -> bool {
        self.states.get(key).unwrap_or(&0) > &0
    }

    fn is_at(&self, key: &str) -> bool {
        self.is_in(key) || key == self.last_seen_closed
    }

    fn consume<'b>(&mut self, event: &'b Event) {
        // reset last seen closing tag, as we only want that to cover the closing tag
        if self.last_seen_closed != "" {
            self.last_seen_closed = "".into();
        }
        match event {
            Event::Start(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.states.insert(name, current + 1);
            }
            Event::End(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.last_seen_closed = name.clone();
                self.states.insert(name, current - 1);
            }
            Event::Eof => self.is_eof = true,
            _ => {}
        }
    }
}

pub fn get_content_controls(data: &ZipData) -> ParsedDocuments {
    let mut documents = HashMap::new();
    for (filename, string) in data {
        if has_content_control(&string) {
            let enc_str = str::from_utf8(&string).expect("should be utf-8 encoded string");
            let mut reader = Reader::from_str(enc_str);
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
                        if state.is_in("w:sdtContent") && state.is_in("w:p") && state.is_at("w:pPr")
                        {
                            control.paragraph_params.push(e.clone());
                        }
                        if state.is_in("w:sdtContent") && state.is_in("w:r") && state.is_at("w:rPr")
                        {
                            control.run_params.push(e.clone());
                        }
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
            mapped_data.insert(filename.into(), writer.into_inner().into_inner());
        } else {
            mapped_data.insert(filename.into(), data.clone());
        }
    }
    mapped_data
}

fn get_intersecting_control<'a>(
    index: i64,
    controls: &'a Vec<ContentControl<'a>>,
) -> Option<&'a ContentControl<'a>> {
    for control in controls {
        if index >= control.content_begin && index < control.content_end {
            return Some(control);
        }
    }
    return None;
}

pub fn map_content_controls(
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
                        let _ = write_content(&control, &mut writer, *mapped);
                    }
                } else {
                    let _ = writer.write_event(event);
                }
            }
            mapped_data.insert(filename.into(), writer.into_inner().into_inner());
        } else {
            mapped_data.insert(filename.into(), data.clone());
        }
    }
    mapped_data
}

fn parse_mappings(mappings: &HashMap<&str, &str>) -> Result<i8, ParserError> {
    for (key, value) in mappings {
        let mut reader = Reader::from_str(value);
        let mut txt = Vec::new();
        loop {
            match reader.read_event() {
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                Ok(Event::Eof) => break,
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"tag1" => println!(
                        "attributes values: {:?}",
                        e.attributes().map(|a| a.unwrap().value).collect::<Vec<_>>()
                    ),
                    _ => (),
                },
                Ok(Event::Text(e)) => txt.push(e.unescape().unwrap().into_owned()),
                _ => ()
            }
        }
        println!("{} {:?}", key, txt)
    }
    Err(ParserError {
        data: "Nothing has been implemented".into(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use std::fs;
    use std::io::{BufReader, BufWriter};
    use tempfile::tempfile;

    fn load_path(path: &str) -> ZipData {
        let fname = std::path::Path::new(&path);
        let file = fs::File::open(fname).unwrap();
        let reader = BufReader::new(file);
        list_zip_contents(reader).unwrap()
    }

    #[test]
    fn full_operation() {
        let input_data = load_path("tests/data/content_controlled_document.docx");
        let expected_data = load_path("tests/data/content_controlled_document_expected.docx");

        let mappings = HashMap::from([
            ("Title", "Brave New World"),
            ("Sidematter", "Into a brave new world"),
            ("WritingDate", "12.12.2012"),
            ("Author", "Bruce Wayne"),
            ("MainContent", "This is rich coming from you."),
        ]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);

        let mut outfile = tempfile().unwrap();
        let _ = zip_dir(&mapped_data, &mut outfile);
        let nreader = BufReader::new(outfile);
        let result_data = list_zip_contents(nreader).unwrap();

        for (e_k, e_v) in expected_data {
            assert_eq!(e_v, result_data[&e_k]);
        }
    }

    #[test]
    fn run_with_params() {
        let input_data = load_path("tests/data/run_with_params.docx");
        let expected_data = load_path("tests/data/run_with_params_expected.docx");
        let mappings = HashMap::from([
            ("RunField", "Something new"),
        ]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);
        let mut outfile = tempfile().unwrap();
        let _ = zip_dir(&mapped_data, &mut outfile);
        let nreader = BufReader::new(outfile);
        let result_data = list_zip_contents(nreader).unwrap();

        for (e_k, e_v) in expected_data {
            assert_eq!(e_v, result_data[&e_k]);
        }
    }

    #[test]
    fn preserve_images() {
        let input_data = load_path("tests/data/run_with_params_imgs.docx");
        let mappings = HashMap::from([
            ("RunField", "Something new"),
        ]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);

        let file = fs::File::create("test.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }
}
