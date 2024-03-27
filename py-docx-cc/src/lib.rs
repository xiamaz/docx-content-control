use std::collections::HashMap;
use std::io;
use std::borrow::Cow;
use pyo3::prelude::*;

#[pyfunction]
fn map_content_controls<'a>(template_data: Vec<u8>, mappings: HashMap<String, String>) -> Cow<'a, [u8]> {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let controlled_docs = docx_cc::get_content_controls(&data);
    let mappings_str = mappings.iter().map(|(k, v)| (k.as_str(), v.as_str())).collect::<HashMap<_, _>>();
    let mapped_data = docx_cc::map_content_controls(&data, &controlled_docs, &mappings_str);
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

#[pyfunction]
fn get_content_controls(template_data: Vec<u8>) -> PyResult<HashMap<String, Vec<String>>> {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let controlled_docs = docx_cc::get_content_controls(&data);

    let mut result = HashMap::new();
    for (_name, docdata) in controlled_docs {
        for control in docdata.controls {
            result.entry(control.tag.clone()).or_insert(Vec::new()).push(control.get_control_type().to_string());
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
