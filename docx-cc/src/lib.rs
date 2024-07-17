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

use serde::Serialize;

static MISSING_STR: &str = "MISSING";

pub type ZipData = HashMap<String, Vec<u8>>;
pub type Mapping = HashMap<String, String>;
pub type RepeatMapping = HashMap<String, Vec<Mapping>>;

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

#[derive(Debug, PartialEq, Serialize)]
pub enum ContentControlType {
    Unsupported,
    RichText,
    Text,
    ComboBox,
    DropdownList,
    Date,
    RepeatingSection,
    RepeatingSectionItem,
}

impl ContentControlType {
    pub fn parse_string(value: &str) -> Option<ContentControlType> {
        match value {
            "w:richText" => Some(ContentControlType::RichText),
            "w:text" => Some(ContentControlType::Text),
            "w:comboBox" => Some(ContentControlType::ComboBox),
            "w:dropDownList" => Some(ContentControlType::DropdownList),
            "w:date" => Some(ContentControlType::Date),
            "w15:repeatingSection" => Some(ContentControlType::RepeatingSection),
            "w15:repeatingSectionItem" => Some(ContentControlType::RepeatingSectionItem),
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
                ContentControlType::RepeatingSection => "w15:repeatingSection".to_string(),
                ContentControlType::RepeatingSectionItem => "w15:repeatingSectionItem".to_string(),
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

fn get_intersecting_control_position(
    index: i64,
    controls: &[ContentControlPosition],
) -> Option<&ContentControlPosition> {
    controls
        .iter()
        .find(|&control| control.intersects_content(index as i32))
}

fn get_contained_control_at<'a>(
    controls: &'a [ContentControlPosition],
    control: &'a ContentControlPosition,
    index: i32,
) -> Option<&'a ContentControlPosition> {
    get_contained_control(controls, control).find(|&c| c.intersects_content(index))
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
    control: &ContentControlPosition,
    content: &str,
    tags: &[&str],
    events: &[Event],
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
                        if control.has_paragraph_params() {
                            for ev in &events[control.paragraph_params_start as usize
                                ..control.paragraph_params_end as usize]
                            {
                                let _ = writer.write_event(ev.clone());
                            }
                        }
                    }
                    "w:r" => {
                        if control.has_run_params() {
                            for ev in &events
                                [control.run_params_start as usize..control.run_params_end as usize]
                            {
                                let _ = writer.write_event(ev.clone());
                            }
                        }
                    }
                    _ => {}
                }
                write_wrap_tags(writer, control, content, &tags[1..], events)
            });
        }
    } else {
        write_parsed_content(writer, content)?
    }
    Ok(())
}

fn write_content<'a, W>(
    control: &ContentControlPosition,
    writer: &'a mut Writer<W>,
    content: &'a str,
    events: &[Event],
) -> Result<(), &'a str>
where
    W: std::io::Write,
{
    if control.contains_paragraph {
        let _ = write_wrap_tags(writer, control, content, &["w:p", "w:r", "w:t"], events);
    } else {
        let _ = write_wrap_tags(writer, control, content, &["w:r", "w:t"], events);
    }
    Ok(())
}

pub struct DocumentData<'a> {
    events: Vec<Event<'a>>,
    pub control_positions: Vec<ContentControlPosition>,
}

type ParsedDocuments<'a> = HashMap<String, DocumentData<'a>>;

#[derive(Debug, Serialize)]
pub struct ContentControlPosition {
    r#type: ContentControlType,
    tag: String,
    begin: i32,
    end: i32,
    content_begin: i32,
    content_end: i32,
    paragraph_params_start: i32,
    paragraph_params_end: i32,
    contains_paragraph: bool,
    run_params_start: i32,
    run_params_end: i32,
}

impl ContentControlPosition {
    fn new() -> Self {
        ContentControlPosition {
            r#type: ContentControlType::Unsupported,
            tag: "".into(),
            begin: -1,
            end: -1,
            content_begin: -1,
            content_end: -1,
            paragraph_params_start: -1,
            paragraph_params_end: -1,
            contains_paragraph: false,
            run_params_start: -1,
            run_params_end: -1,
        }
    }

    fn intersects_header(&self, index: i32) -> bool {
        self.begin != -1 && index > self.begin && self.content_begin == -1 && self.end == -1
    }

    fn intersects_content(&self, index: i32) -> bool {
        index >= self.content_begin && index < self.content_end
    }

    fn content_opened(&self) -> bool {
        self.content_begin != -1
    }

    fn closed(&self) -> bool {
        self.end != -1
    }

    fn content_closed(&self) -> bool {
        self.content_end != -1
    }

    fn has_paragraph_params(&self) -> bool {
        self.paragraph_params_start >= 0 && self.paragraph_params_end >= 0
    }

    fn has_run_params(&self) -> bool {
        self.run_params_start >= 0 && self.run_params_end >= 0
    }

    pub fn get_tag(&self) -> &str {
        &self.tag
    }

    pub fn get_type(&self) -> &ContentControlType {
        &self.r#type
    }
}

impl Default for ContentControlPosition {
    fn default() -> Self {
        Self::new()
    }
}

pub struct DocumentState {
    states: HashMap<String, i32>,
    positions: HashMap<String, i32>,
    controls: Vec<ContentControlPosition>,
    is_eof: bool,
    last_seen_closed: String,
    counter: i32,
}

impl DocumentState {
    fn new() -> DocumentState {
        DocumentState {
            states: HashMap::new(),
            positions: HashMap::new(),
            controls: Vec::new(),
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
                match name.as_str() {
                    "w:sdt" => {
                        self.controls.push(ContentControlPosition {
                            begin: self.counter,
                            ..Default::default()
                        });
                    }
                    "w:sdtContent" => {
                        for ctrl in self.controls.iter_mut().rev() {
                            if !ctrl.content_opened() {
                                ctrl.content_begin = self.counter;
                                break;
                            }
                        }
                    }
                    "w:p" => {
                        if self.is_in("w:sdtContent") {
                            if let Some(ctrl) = self.controls.iter_mut().next_back() {
                                ctrl.contains_paragraph = true;
                            }
                        }
                    }
                    "w:rPr" => {
                        if self.is_in("w:sdtContent") && self.is_in("w:r") {
                            if let Some(ctrl) = self.controls.iter_mut().next_back() {
                                if ctrl.run_params_start < 0 {
                                    ctrl.run_params_start = self.counter;
                                }
                            }
                        }
                    }
                    "w:pPr" => {
                        if self.is_in("w:sdtContent") && self.is_in("w:p") {
                            if let Some(ctrl) = self.controls.iter_mut().next_back() {
                                if ctrl.paragraph_params_start < 0 {
                                    ctrl.paragraph_params_start = self.counter;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                let current = self.states.get(&name).unwrap_or(&0);
                self.last_seen_closed = name.clone();
                match name.as_str() {
                    "w:sdt" => {
                        for ctrl in self.controls.iter_mut().rev() {
                            if !ctrl.closed() {
                                ctrl.end = self.counter;
                                // Content Control defaults to RichText if no type has been given.
                                if let ContentControlType::Unsupported = ctrl.r#type {
                                    ctrl.r#type = ContentControlType::RichText
                                }
                                break;
                            }
                        }
                    }
                    "w:sdtContent" => {
                        for ctrl in self.controls.iter_mut().rev() {
                            if !ctrl.content_closed() {
                                ctrl.content_end = self.counter;
                                break;
                            }
                        }
                    }
                    "w:rPr" => {
                        if self.is_in("w:sdtContent") && self.is_in("w:r") {
                            if let Some(ctrl) = self.controls.iter_mut().next_back() {
                                if ctrl.run_params_end < 0 {
                                    ctrl.run_params_end = self.counter + 1;
                                }
                            }
                        }
                    }
                    "w:pPr" => {
                        if self.is_in("w:sdtContent") && self.is_in("w:p") {
                            if let Some(ctrl) = self.controls.iter_mut().next_back() {
                                if ctrl.paragraph_params_end < 0 {
                                    ctrl.paragraph_params_end = self.counter + 1;
                                }
                            }
                        }
                    }
                    _ => {}
                }
                self.states.insert(name, current - 1);
            }
            Event::Empty(e) => {
                let name = String::from_utf8_lossy(e.name().into_inner()).to_string();
                if self.is_in("w:sdtPr") {
                    if let Some(t) = ContentControlType::parse_string(&name) {
                        for ctrl in self.controls.iter_mut().rev() {
                            if ctrl.intersects_header(self.counter) {
                                ctrl.r#type = t;
                                break;
                            }
                        }
                    } else if name == "w:tag" {
                        if let Some(ctrl) = self.controls.iter_mut().next_back() {
                            for attr in e.attributes().flatten() {
                                if attr.key == QName(b"w:val") {
                                    let vwal = String::from_utf8_lossy(&attr.value).into();
                                    ctrl.tag = vwal;
                                }
                            }
                        }
                    }
                }
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
            let mut events: Vec<Event> = Vec::new();
            while !state.is_eof {
                let event = reader.read_event();
                match event {
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    Ok(e) => {
                        state.consume(&e);
                        events.push(e.clone());
                    }
                }
            }
            documents.insert(
                filename.into(),
                DocumentData {
                    events,
                    control_positions: state.controls,
                },
            );
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

pub fn get_contained_control<'a>(
    controls: &'a [ContentControlPosition],
    control: &'a ContentControlPosition,
) -> impl Iterator<Item = &'a ContentControlPosition> + 'a {
    controls
        .iter()
        .filter(|c| c.begin >= control.content_begin && c.end <= control.content_end)
}

pub fn map_content_controls(
    data: &ZipData,
    controlled: &ParsedDocuments,
    mappings: &Mapping,
    repeat_mappings: &RepeatMapping,
) -> ZipData {
    let mut mapped_data = ZipData::new();
    for (filename, data) in data {
        if let Some(doc) = controlled.get(filename) {
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            for (i, event) in doc.events.iter().enumerate() {
                if let Some(control) =
                    get_intersecting_control_position(i as i64, &doc.control_positions)
                {
                    if control.content_begin == i as i32 {
                        let _ = writer.write_event(event);
                        match control.r#type {
                            ContentControlType::RepeatingSection => {
                                let default_values = Vec::new();
                                let new_values = repeat_mappings.get(control.tag.as_str()).unwrap_or(&default_values);
                                for new_value in new_values.iter() {
                                    if let Some(section_item) = get_contained_control(
                                        &doc.control_positions,
                                        control,
                                    )
                                    .find(|c| c.r#type == ContentControlType::RepeatingSectionItem)
                                    {
                                        for i_item in section_item.begin..section_item.end + 1 {
                                            let ev_item = &doc.events[i_item as usize];
                                            if let Some(ctrl_item) = get_contained_control_at(
                                                &doc.control_positions,
                                                section_item,
                                                i_item,
                                            ) {
                                                if ctrl_item.content_begin == i_item {
                                                    let _ = writer.write_event(ev_item);
                                                    let new_value = new_value
                                                        .get(&ctrl_item.tag).map(String::as_str)
                                                        .unwrap_or(MISSING_STR);
                                                    let _ = write_content(
                                                        ctrl_item,
                                                        &mut writer,
                                                        new_value,
                                                        &doc.events,
                                                    );
                                                }
                                            } else {
                                                let _ = writer.write_event(ev_item);
                                            }
                                        }
                                    }
                                }
                            }
                            _ => {
                                let new_value =
                                    mappings.get(&control.tag).map(String::as_str).unwrap_or(MISSING_STR);
                                let _ = write_content(control, &mut writer, new_value, &doc.events);
                            }
                        }
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
    use pretty_assertions::assert_eq;
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
    fn document_state() {
        let input_data = load_path("tests/data/content_controlled_document.docx");
        for (_filename, string) in input_data {
            if has_content_control(&string) {
                let enc_str = str::from_utf8(&string).expect("should be utf-8 encoded string");
                let mut reader = Reader::from_str(enc_str);
                let mut state = DocumentState::new();
                while !state.is_eof {
                    let event = reader.read_event();
                    match event {
                        Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                        Ok(e) => state.consume(&e),
                    }
                }
                assert!(!state.controls.is_empty());
                assert!(state.controls.iter().any(|c| c.begin >= 0
                    && c.end >= 0
                    && c.content_begin >= 0
                    && c.content_end >= 0));
            }
        }
    }

    #[test]
    fn full_operation() {
        let input_data = load_path("tests/data/content_controlled_document.docx");
        let expected_data = load_path("tests/data/content_controlled_document_expected.docx");

        let mappings = HashMap::from([
            ("Title".into(), "Brave New World".into()),
            ("Sidematter".into(), "Into a brave new world".into()),
            ("WritingDate".into(), "12.12.2012".into()),
            ("Author".into(), "Bruce Wayne".into()),
            ("MainContent".into(), "This is rich coming from you.".into()),
        ]);
        let repeat_mappings = HashMap::from([]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );

        let mut outfile = tempfile().unwrap();
        let _ = zip_dir(&mapped_data, &mut outfile);
        let nreader = BufReader::new(outfile);
        let result_data = list_zip_contents(nreader).unwrap();

        for (e_k, e_v) in expected_data {
            assert_eq!(
                String::from_utf8_lossy(&e_v),
                String::from_utf8_lossy(&result_data[&e_k])
            );
        }
    }

    #[test]
    fn run_with_params() {
        let input_data = load_path("tests/data/run_with_params.docx");
        let expected_data = load_path("tests/data/run_with_params_expected.docx");
        let mappings = HashMap::from([("RunField".into(), "Something new".into())]);
        let repeat_mappings = HashMap::from([]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );
        let mut outfile = tempfile().unwrap();
        let _ = zip_dir(&mapped_data, &mut outfile);
        let nreader = BufReader::new(outfile);
        let result_data = list_zip_contents(nreader).unwrap();

        for (e_k, e_v) in expected_data {
            assert_eq!(
                String::from_utf8_lossy(&e_v),
                String::from_utf8_lossy(&result_data[&e_k])
            );
        }
    }

    #[test]
    fn preserve_images() {
        let input_data = load_path("tests/data/run_with_params_imgs.docx");
        let mappings = HashMap::from([("RunField".into(), "Something new".into())]);
        let repeat_mappings = HashMap::from([]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );

        let file = fs::File::create("test.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }

    #[test]
    fn complex_replacement() {
        let input_data = load_path("tests/data/run_with_params_imgs.docx");
        let mappings = HashMap::from([(
            "RunField".into(),
            "<w:t>Something</w:t><w:cr/><w:t>new</w:t>".into(),
        )]);
        let repeat_mappings = HashMap::from([]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );

        let file = fs::File::create("test2.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }

    #[test]
    fn repeat_replacement() {
        let input_data = load_path("tests/data/TownLandRiver.docx");
        let expected_data = load_path("tests/data/TownLandRiver_expected.docx");
        let mappings = HashMap::from([]);
        let data = r#"
        {
            "Entry": [
            {
                "Town": "Cottbus",
                "Land": "Brandenburg",
                "River": "Dahme"
            },
            {
                "Town": "Aachen",
                "Land": "NRW",
                "River": "Wurm"
            }
            ]
        }
            "#;
        let repeat_mappings: RepeatMapping = serde_json::from_str(data).unwrap();
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );

        let mut outfile = tempfile().unwrap();
        let _ = zip_dir(&mapped_data, &mut outfile);
        let nreader = BufReader::new(outfile);
        let result_data = list_zip_contents(nreader).unwrap();

        for (e_k, e_v) in expected_data {
            assert_eq!(
                String::from_utf8_lossy(&e_v),
                String::from_utf8_lossy(&result_data[&e_k])
            );
        }
    }

    #[test]
    fn table_replacement() {
        let input_data = load_path("tests/data/content_controlled_document.docx");
        let mappings = HashMap::from([(
            "MainContent".into(),
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
</w:tbl>"
                .into(),
        )]);
        let repeat_mappings = HashMap::from([]);
        let controlled_documents = get_content_controls(&input_data);
        let mapped_data = map_content_controls(
            &input_data,
            &controlled_documents,
            &mappings,
            &repeat_mappings,
        );

        let file = fs::File::create("test3.docx").unwrap();
        let mut writer = BufWriter::new(file);
        let _ = zip_dir(&mapped_data, &mut writer);
    }
}
