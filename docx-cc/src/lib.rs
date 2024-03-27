use std::collections::HashMap;
use std::collections::HashSet;
use std::error::Error;
use std::fmt;
use std::io::prelude::*;
use std::io::Cursor;
use std::str;

use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Writer;
use zip::write::FileOptions;

use quick_xml::reader::Reader;

pub type ZipData = HashMap<String, Vec<u8>>;

#[derive(Debug)]
struct ParserError {}

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
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/**
 * Check if the string contains an sdt tag (Ruby Inline-Level Structured Document Tag)
 */
fn has_content_control(text: &[u8]) -> bool {
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

impl ContentControlType {
    pub fn parse_string(value: &str) -> Option<ContentControlType> {
        match value {
            "w:richText" => Some(ContentControlType::RichText),
            "w:text" => Some(ContentControlType::Text),
            "w:comboBox" => Some(ContentControlType::ComboBox),
            "w:dropDownList" => Some(ContentControlType::DropdownList),
            "w:date" => Some(ContentControlType::Date),
            _ => None,
        }
    }
}

impl fmt::Display for ContentControlType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{}",
            match &self {
                ContentControlType::RichText => "w:richText".to_string(),
                ContentControlType::Text => "w:text".to_string(),
                ContentControlType::ComboBox => "w:comboBox".to_string(),
                ContentControlType::DropdownList => "w:dropDownList".to_string(),
                ContentControlType::Date => "w:date".to_string(),
                ContentControlType::Unsupported => "unsupported".to_string(),
            }
        )
    }
}

fn get_tag_types(content: &str) -> HashSet<String> {
    let mut content_reader = Reader::from_str(content);
    let mut tag_names = HashSet::new();
    loop {
        let event = content_reader
            .read_event()
            .expect("should be a well formatted xml");
        match event {
            Event::Eof => break,
            Event::Start(e) => {
                tag_names.insert(String::from_utf8_lossy(e.name().into_inner()).to_string());
            }
            Event::Empty(e) => {
                tag_names.insert(String::from_utf8_lossy(e.name().into_inner()).to_string());
            }
            _ => {}
        };
    }
    tag_names
}

pub struct ContentControl<'a> {
    pub tag: String,
    value: String,
    ct_type: ContentControlType,
    params: HashMap<String, String>,
    contains_paragraph: bool,
    paragraph_position: i32,
    paragraph_params: Vec<Event<'a>>,
    run_position: i32,
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
            paragraph_position: -1,
            run_params: Vec::new(),
            run_position: -1,
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
            } else if let Some(t) = ContentControlType::parse_string(k) {
                self.ct_type = t;
            }
        }
    }
}

fn write_parsed_content<W>(writer: &mut Writer<W>, content: &str) -> Result<(), quick_xml::Error>
where
    W: std::io::Write,
{
    let mut content_reader = Reader::from_str(content);
    loop {
        let event = content_reader
            .read_event()
            .expect("should be a well formatted xml");
        let _ = match event {
            Event::Eof => break,
            _ => writer.write_event(event),
        };
    }
    Ok(())
}

fn write_wrap_tags<W>(
    writer: &mut Writer<W>,
    control: &ContentControl,
    content: &str,
    tags: &[&str],
) -> Result<(), quick_xml::Error>
where
    W: std::io::Write,
{
    let content_tags = get_tag_types(content);
    if !tags.is_empty() {
        let tag = tags[0];
        if content_tags.contains(tag) {
            write_parsed_content(writer, content)?
        } else {
            let _ = writer.create_element(tag).write_inner_content(|writer| {
                match tag {
                    "w:p" => {
                        for ev in &control.paragraph_params {
                            let _ = writer.write_event(ev);
                        }
                    }
                    "w:r" => {
                        for ev in &control.run_params {
                            let _ = writer.write_event(ev);
                        }
                    }
                    _ => {}
                }
                write_wrap_tags(writer, control, content, &tags[1..])
            });
        }
    } else {
        write_parsed_content(writer, content)?
    }
    Ok(())
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
        let _ = write_wrap_tags(writer, control, content, &["w:p", "w:r", "w:t"]);
    } else {
        let _ = write_wrap_tags(writer, control, content, &["w:r", "w:t"]);
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
    positions: HashMap<String, i32>,
    is_eof: bool,
    last_seen_closed: String,
    counter: i32,
}

impl DocumentState {
    fn new() -> DocumentState {
        DocumentState {
            states: HashMap::new(),
            positions: HashMap::new(),
            is_eof: false,
            last_seen_closed: "".into(),
            counter: 0,
        }
    }

    fn is_in(&self, key: &str) -> bool {
        self.states.get(key).unwrap_or(&0) > &0
    }

    fn is_at(&self, key: &str) -> bool {
        self.is_in(key) || key == self.last_seen_closed
    }

    fn last_seen_position(&self, key: &str) -> &i32 {
        self.positions.get(key).unwrap()
    }

    fn consume(&mut self, event: &Event) {
        // reset last seen closing tag, as we only want that to cover the closing tag
        if !self.last_seen_closed.is_empty() {
            self.last_seen_closed = "".into();
        }
        match event {
            Event::Start(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.states.insert(name.clone(), current + 1);
                self.positions.insert(name.clone(), self.counter);
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
        self.counter += 1;
    }
}

pub fn get_content_controls(data: &ZipData) -> ParsedDocuments {
    let mut documents = HashMap::new();
    for (filename, string) in data {
        if has_content_control(string) {
            let enc_str = str::from_utf8(string).expect("should be utf-8 encoded string");
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
                            if control.paragraph_position < 0 {
                                control.paragraph_position = *state.last_seen_position("w:p");
                            }
                            if *state.last_seen_position("w:p") == control.paragraph_position {
                                control.paragraph_params.push(e.clone());
                            }
                        }
                        if state.is_in("w:sdtContent") && state.is_in("w:r") && state.is_at("w:rPr")
                        {
                            if control.run_position < 0 {
                                control.run_position = *state.last_seen_position("w:r");
                            }
                            if *state.last_seen_position("w:r") == control.run_position {
                                control.run_params.push(e.clone());
                            }
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
                                    for attr in v.attributes().flatten() {
                                        if attr.key == QName(b"w:val") {
                                            vwal = String::from_utf8_lossy(&attr.value).into();
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
                                    control.value.push_str(v.unescape().unwrap().as_ref());
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

/**
 * Remove all content controls while retaining content.
 */
pub fn remove_content_controls(data: &ZipData) -> ZipData {
    let mut cleared_data = ZipData::new();
    for (filename, doc_string) in data {
        if has_content_control(doc_string) {
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            let doc_string_enc =
                str::from_utf8(doc_string).expect("should be utf-8 encoded string");
            let mut reader = Reader::from_str(doc_string_enc);
            let mut state = DocumentState::new();
            while !state.is_eof {
                let event = reader.read_event();
                match event {
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    Ok(e) => {
                        state.consume(&e);
                        match &e {
                            Event::Start(v) => {
                                if v.name() != QName(b"w:sdtContent")
                                    && v.name() != QName(b"w:sdt")
                                    && !state.is_at("w:sdtPr")
                                {
                                    let _ = writer.write_event(e);
                                }
                            }
                            Event::End(v) => {
                                if v.name() != QName(b"w:sdtContent")
                                    && v.name() != QName(b"w:sdt")
                                    && !state.is_at("w:sdtPr")
                                {
                                    let _ = writer.write_event(e);
                                }
                            }
                            _ => {
                                if !state.is_at("w:sdtPr") {
                                    let _ = writer.write_event(e);
                                }
                            }
                        }
                    }
                }
            }
            cleared_data.insert(filename.into(), writer.into_inner().into_inner());
        } else {
            cleared_data.insert(filename.into(), doc_string.clone());
        }
    }
    cleared_data
}

fn get_intersecting_control<'a>(
    index: i64,
    controls: &'a [ContentControl<'a>],
) -> Option<&'a ContentControl<'a>> {
    controls
        .iter()
        .find(|&control| index >= control.content_begin && index < control.content_end)
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
                        let _ = write_content(control, &mut writer, mapped);
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
        let mappings = HashMap::from([("RunField", "Something new")]);
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
        let mappings = HashMap::from([("RunField", "Something new")]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);

        let file = fs::File::create("test.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }

    #[test]
    fn complex_replacement() {
        let input_data = load_path("tests/data/run_with_params_imgs.docx");
        let mappings = HashMap::from([("RunField", "<w:t>Something</w:t><w:cr/><w:t>new</w:t>")]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);

        let file = fs::File::create("test2.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }

    #[test]
    fn table_replacement() {
        let input_data = load_path("tests/data/content_controlled_document.docx");
        let mappings = HashMap::from([(
            "MainContent",
            "<w:tbl>
<w:tblPr>
<w:tblStyle w:val=\"TableGrid\"/>
<w:tblW w:w=\"5000\" w:type=\"pct\"/>
</w:tblPr>
<w:tblGrid>
<w:gridCol w:w=\"2880\"/>
<w:gridCol w:w=\"2880\"/>
<w:gridCol w:w=\"2880\"/>
</w:tblGrid>
<w:tr>
<w:tc>
<w:tcPr>
<w:tcW w:w=\"2880\" w:type=\"dxa\"/>
</w:tcPr>
<w:p>
<w:r>
<w:t>AAA</w:t>
</w:r>
</w:p>
</w:tc>
<w:tc>
<w:tcPr>
<w:tcW w:w=\"2880\" w:type=\"dxa\"/>
</w:tcPr>
<w:p>
<w:r>
<w:t>BBB</w:t>
</w:r>
</w:p>
</w:tc>
<w:tc>
<w:tcPr>
<w:tcW w:w=\"2880\" w:type=\"dxa\"/>
</w:tcPr>
<w:p>
<w:r>
<w:t>CCC</w:t>
</w:r>
</w:p>
</w:tc>
</w:tr>
</w:tbl>",
        )]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(&input_data, &controlled_documents, &mappings);

        let file = fs::File::create("test3.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }
}
