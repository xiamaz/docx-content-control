use std::collections::HashMap;
use std::io;
use pyo3::types::PyBytes;
use pyo3::prelude::*;

#[pyfunction]
fn map_content_controls<'a>(py: Python<'a>, template_data: Vec<u8>, mappings: HashMap<&str, &str>) -> &'a PyBytes {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let controlled_docs = docx_cc::get_content_controls(&data);
    let mapped_data = docx_cc::map_content_controls(&data, &controlled_docs, &mappings);
    let mut buffer: Vec<u8> = Vec::new();
    let mut outc = io::Cursor::new(&mut buffer);
    let _ = docx_cc::zip_dir(&mapped_data, &mut outc);

    PyBytes::new(py, &buffer)
}

#[pyfunction]
fn remove_content_controls(py: Python, template_data: Vec<u8>) -> &PyBytes {
    let cursor = io::Cursor::new(template_data);
    let reader = io::BufReader::new(cursor);
    let data = docx_cc::list_zip_contents(reader).unwrap();
    let result = docx_cc::remove_content_controls(&data);
    let mut buffer: Vec<u8> = Vec::new();
    let mut outc = io::Cursor::new(&mut buffer);
    let _ = docx_cc::zip_dir(&result, &mut outc);

    PyBytes::new(py, &buffer)
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
fn py_docx_cc(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(remove_content_controls, m)?)?;
    m.add_function(wrap_pyfunction!(map_content_controls, m)?)?;
    m.add_function(wrap_pyfunction!(get_content_controls, m)?)?;
    Ok(())
}
