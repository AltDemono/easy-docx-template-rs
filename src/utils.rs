use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read};
use serde_json::Value;
use zip::{ZipArchive};
use crate::DOCX;

pub fn read_raw_docx(path: &str) -> HashMap<String, String> {
    let file = File::open(path).unwrap();
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader).unwrap();
    let mut hashmap = HashMap::<String, String>::new();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();

        if name == "word/document.xml"
            || name.starts_with("word/header")
            || name.starts_with("word/footer")
        {
            let mut xml_content = String::new();
            file.read_to_string(&mut xml_content).unwrap();
            hashmap.insert(name, xml_content);
        }
    }

    hashmap
}

pub fn process_text(docx: &DOCX, text: &mut String) {
    for (placeholder, value) in &docx.placeholders {
        *text = text.replace(placeholder, value);
    }
}

pub fn process_json_map(docx: &mut DOCX, prefix: &str, map: &serde_json::Map<String, Value>) {
    for (k, v) in map {
        let full_key = if prefix.is_empty() {
            k.to_string()
        } else {
            format!("{}.{}", prefix, k)
        };

        match v {
            Value::String(s) => {
                let placeholder = format!("{{{{{}}}}}", full_key);
                docx.add_placeholder(&placeholder, s);
            }
            Value::Object(obj) => {
                process_json_map(docx, &full_key, obj);
            }
            Value::Number(n) => {
                let placeholder = format!("{{{{{}}}}}", full_key);
                docx.add_placeholder(&placeholder, &n.to_string());
            }
            Value::Bool(b) => {
                let placeholder = format!("{{{{{}}}}}", full_key);
                docx.add_placeholder(&placeholder, &b.to_string());
            }
            _ => {}
        }
    }
}