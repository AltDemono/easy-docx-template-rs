use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Write};
use quick_xml::events::{BytesText, Event};
use quick_xml::Writer;
use serde_json::{value, Value};
use zip::{ZipArchive, ZipWriter};
use zip::write::FileOptions;
use crate::DOCX;

const PLACEHOLDER_START: &str = "{{";
const PLACEHOLDER_END: &str = "}}";

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

pub fn process_text<'a, T: ToString + std::fmt::Display>(placeholders: &HashMap<String, T>, text: &'a mut String) -> &'a String {
    for (placeholder, value) in placeholders {
        *text = text.replace(placeholder, value.to_string().as_str());
    }

    text
}

pub fn process_json_map<T>(
    docx: &mut DOCX,
    prefix: &str,
    map: &serde_json::Map<String, Value>,
)
where
    T: From<String>, T: std::fmt::Display
{
    for (k, v) in map {
        let full_key = if prefix.is_empty() {
            k.to_string()
        } else {
            format!("{}.{}", prefix, k)
        };

        let placeholder = format!("{{{{{}}}}}", full_key);

        match v {
            Value::String(s) => {
                docx.add_placeholder(&placeholder, &*s.clone());
            }
            Value::Number(n) => {
                docx.add_placeholder(&placeholder, &*n.to_string());
            }
            Value::Bool(b) => {
                docx.add_placeholder(&placeholder, &*b.to_string());
            }
            Value::Object(obj) => {
                process_json_map::<T>(docx, &full_key, obj);
            }
            _ => {}
        }
    }
}


pub fn render_docx(docx: &DOCX) -> Vec<u8> {
    let media_dir = "word/media/";

    // open original
    let input_file = File::open(&docx.file_name).unwrap();
    let mut archive = ZipArchive::new(input_file).unwrap();

    // write into memory buffer
    let buffer = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(buffer);
    let options = FileOptions::default();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();

        if let Some(content) = docx.content.get(&name) {
            zip.start_file(&name, options).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        } else if name.starts_with(media_dir)
            && (name.ends_with(".jpeg") || name.ends_with(".png") || name.ends_with(".jpg"))
            && docx.placeholder_images.contains_key(&name)
        {
            let image_path = docx.placeholder_images.get(&name).unwrap();
            let mut img_file = File::open(image_path).unwrap();
            let mut img_data = Vec::new();
            img_file.read_to_end(&mut img_data).unwrap();

            zip.start_file(&name, options).unwrap();
            zip.write_all(&img_data).unwrap();
        } else {
            zip.start_file(&name, options).unwrap();
            std::io::copy(&mut file, &mut zip).unwrap();
        }
    }

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

pub fn init_placeholders(placeholders: &HashMap<String, String>, content: &String) -> String {
    let mut xml_writer = Writer::new(Cursor::new(Vec::new()));
    let mut reader = quick_xml::Reader::from_str(&*content);
    reader.trim_text(false);
    let mut buf = Vec::new();

    let mut current_placeholder = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                xml_writer.write_event(Event::Start(e)).unwrap();
            }
            Ok(Event::Text(e)) => {
                let mut text = e.unescape().unwrap().into_owned();

                if !current_placeholder.is_empty() || text.contains(PLACEHOLDER_START) {
                    current_placeholder.push_str(&text);

                    if current_placeholder.contains(PLACEHOLDER_START) && current_placeholder.contains(PLACEHOLDER_END) {
                        process_text(placeholders, &mut current_placeholder);

                        xml_writer.write_event(Event::Text(BytesText::new(&current_placeholder))).unwrap();
                        current_placeholder.clear();
                    }
                } else {
                    process_text(placeholders, &mut text);
                    xml_writer.write_event(Event::Text(BytesText::new(&text))).unwrap();
                }
            }
            Ok(Event::End(e)) => {
                xml_writer.write_event(Event::End(e)).unwrap();
            }
            Ok(Event::Empty(e)) => {
                xml_writer.write_event(Event::Empty(e)).unwrap();
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                xml_writer.write_event(e).unwrap();
            }
            Err(_) => break,
        }
        buf.clear();
    }

    let result = xml_writer.into_inner().into_inner();
    let result_str = String::from_utf8(result).unwrap();
    result_str
}