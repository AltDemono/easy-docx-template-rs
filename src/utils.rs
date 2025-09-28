use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Write};
use quick_xml::events::{BytesText, Event};
use quick_xml::Writer;
use serde_json::{Value};
use zip::{ZipArchive, ZipWriter};
use zip::write::FileOptions;
use regex::Regex;
// Crate
use crate::DOCX;
use crate::placeholder_helpers::{len_helper, lower_helper, upper_helper};

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

pub fn process_text(placeholders: &HashMap<String, Value>, text: &mut String) -> String {
    for (placeholder, value) in placeholders {
        let replacement = match value {
            Value::String(s) => s.clone(),
            _ => value.to_string(),
        };
        *text = text.replace(placeholder, &*replacement);
    }

    text.to_string()
}

pub fn process_json_map(
    docx: &mut DOCX,
    prefix: &str,
    map: &serde_json::Map<String, Value>,
)
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
                docx.add_placeholder(&placeholder, Value::from(&*s.clone()));
            }
            Value::Number(n) => {
                docx.add_placeholder(&placeholder, Value::from(&*n.to_string()));
            }
            Value::Bool(b) => {
                docx.add_placeholder(&placeholder, Value::from(&*b.to_string()));
            }
            Value::Array(a) => {
                docx.add_placeholder(&placeholder, Value::from(a.clone()));
            }
            Value::Object(obj) => {
                process_json_map(docx, &full_key, obj);
            }
            _ => {}
        }
    }
}

/// Rendering output docx file
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

/// Adding basic placeholder helpers
pub fn add_placeholder_helpers(placeholders: &mut HashMap<String, Value>) -> HashMap<String, Value> {
    for (placeholder, value) in placeholders.clone() {
        let len_placeholder = placeholder.replace("{{", "{{#");
        placeholders.insert(len_placeholder.clone(), Value::from(len_helper(value.clone())));

        let lower_placeholder = placeholder.replace("{{", "{{lower ");
        placeholders.insert(lower_placeholder, Value::from(lower_helper(value.clone())));

        let upper_placeholder = placeholder.replace("{{", "{{upper ");
        placeholders.insert(upper_placeholder, Value::from(upper_helper(value.clone())));
    }
    placeholders.clone()
}

pub fn init_placeholders(placeholders: &mut HashMap<String, Value>, content: &String) -> String {
    let mut xml_writer = Writer::new(Cursor::new(Vec::new()));
    let mut reader = quick_xml::Reader::from_str(&*content);
    reader.trim_text(false);
    let mut buf = Vec::new();

    let mut current_placeholder = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                xml_writer.write_event(Event::Start(e.clone())).unwrap();
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



pub fn remove_paragraph_with_placeholder(xml_content: &str, placeholder: &str) -> String {
    let re = Regex::new(r"<w:p[\s\S]*?</w:p>").unwrap();

    re.replace_all(xml_content, |caps: &regex::Captures| {
        let paragraph = &caps[0];

        let text_only = Regex::new(r"<[^>]+>").unwrap().replace_all(paragraph, "");

        if text_only.contains(placeholder) {
            "".to_string()
        } else {
            paragraph.to_string()
        }
    }).to_string()
}

pub fn init_each_placeholders(xml_content: String, placeholders: &mut HashMap<String, Value>) -> String {
    let mut in_for: bool = false;
    let mut in_block_content: String = String::new();
    let mut variable: Value = Value::Array(Vec::new());
    let mut output = String::new();

    let blocked_symbols = vec!['<', '>', '{', '}'];
    let mut count: usize = 0;
    let mut placeholder_value: String = String::new();
    let mut is_open: bool = false;
    let mut placeholder_start = String::new();

    for letter in xml_content.chars() {
        if count == 2 {
            if !blocked_symbols.contains(&letter) && is_open {
                placeholder_value.push(letter);
            }
        } else if count == 0 && placeholder_value.ends_with("}}") {
            if placeholder_value.contains("{}") {
                placeholder_value = placeholder_value.replace("{}", "");
            }
            if placeholder_value.starts_with("{{#each ") {
                placeholder_start = placeholder_value.clone();
                let var_name = placeholder_value.replace("#each ", "");
                if !placeholders.contains_key(&var_name) {
                    return xml_content
                }
                variable = placeholders[&var_name].clone();
                in_for = true;
                in_block_content.clear();
            } else if placeholder_value.starts_with("{{/each") {
                in_for = false;

                if variable.is_array() {
                    if let Some(arr) = variable.as_array() {
                        for item in arr {
                            let mut block_copy = in_block_content.clone();
                            if let Some(map) = item.as_object() {
                                for (k, v) in map {
                                    let ph = format!("{{{{{}}}}}", k);
                                    block_copy = block_copy.replace(&ph, &v.to_string().replace('"', ""));
                                }
                            }
                            output.push_str(&block_copy);
                        }
                    }
                }
                in_block_content.clear();
            }
            placeholder_value.clear();
        }

        if in_for {
            in_block_content.push(letter);
        } else {
            output.push(letter);
        }

        match letter {
            '{' => {
                count += 1;
                placeholder_value.push(letter);
            }
            '}' => {
                count -= 1;
                placeholder_value.push(letter);
            }
            '<' => is_open = false,
            '>' => is_open = true,
            _ => {}
        }
    }

    output = remove_paragraph_with_placeholder(&output, placeholder_start.as_str());
    output = remove_paragraph_with_placeholder(&output, "{{/each}}");
    output
}