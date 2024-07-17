use std::collections::HashMap;
use std::io;
use std::borrow::Cow;
use pyo3::prelude::*;

#[pyfunction]
fn map_content_controls<'a>(template_data: Vec<u8>, mappings: docx_cc::Mapping, repeat_mappings: docx_cc::RepeatMapping) -> Cow<'a, [u8]> {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let controlled_docs = docx_cc::get_content_controls(&data);
    let mapped_data = docx_cc::map_content_controls(&data, &controlled_docs, &mappings, &repeat_mappings);
    let mut buffer: Vec<u8> = Vec::new();
    let mut outc = io::Cursor::new(&mut buffer);
    let _ = docx_cc::zip_dir(&mapped_data, &mut outc);

    Cow::Owned(buffer)
}

#[pyfunction]
fn remove_content_controls<'a>(template_data: Vec<u8>) -> Cow<'a, [u8]> {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let result = docx_cc::remove_content_controls(&data);
    let mut buffer: Vec<u8> = Vec::new();
    let mut outc = io::Cursor::new(&mut buffer);
    let _ = docx_cc::zip_dir(&result, &mut outc);

    Cow::Owned(buffer)
}

#[pyclass(get_all)]
pub struct ContentControlMetadata {
    pub types: Vec<String>,
    pub children_tags: Vec<String>
}

impl ContentControlMetadata {
    fn new() -> Self {
        ContentControlMetadata {
            types: Vec::new(), children_tags: Vec::new()
        }
    }

    fn add_type(&mut self, type_name: String) {
        self.types.push(type_name);
    }

    fn add_child(&mut self, child_tag: String) {
        self.children_tags.push(child_tag)
    }
}

#[pyfunction]
fn get_content_controls(template_data: Vec<u8>) -> PyResult<HashMap<String, ContentControlMetadata>> {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let controlled_docs = docx_cc::get_content_controls(&data);

    let mut result = HashMap::new();
    for (_name, docdata) in controlled_docs {
        for control in docdata.control_positions.iter() {
            let entry = result.entry(control.get_tag().to_string()).or_insert(ContentControlMetadata::new());
            entry.add_type(control.get_type().to_string());
            for contained_control in docx_cc::get_contained_control(&docdata.control_positions, control) {
                entry.add_child(contained_control.get_tag().to_string())
            }
        }
    }
    Ok(result)
}


/// A Python module implemented in Rust.
#[pymodule]
#[pyo3(name = "py_docx_cc")]
fn py_docx_cc(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(remove_content_controls, m)?)?;
    m.add_function(wrap_pyfunction!(map_content_controls, m)?)?;
    m.add_function(wrap_pyfunction!(get_content_controls, m)?)?;
    Ok(())
}
