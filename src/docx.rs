use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use quick_xml::events::{BytesText, Event};
use quick_xml::Writer;
use serde_json::Value;
use zip::ZipArchive;
use crate::utils::*;

pub const PLACEHOLDER_START: &str = "{{";
pub const PLACEHOLDER_END: &str = "}}";
pub static WORD_PARAGRAPH_TAG: &[u8] = b"w:p";
pub const PLACEHOLDER_BLOCK_START: &str = "{%";
pub const PLACEHOLDER_BLOCK_END: &str = "%}";

pub struct DOCX {
    pub content: HashMap<String, String>,
    pub file_name: String,
    pub placeholders: HashMap<String, String>,
    pub placeholders_blocks: HashMap<String, String>,
    pub placeholder_images: HashMap<String, String>,
}

impl DOCX {
    /// Create new docx manager
    ///
    /// `file_name` input file name
    pub fn new(file_name: String) -> DOCX {
        DOCX {
            file_name,
            content: HashMap::new(),
            placeholders: HashMap::new(),
            placeholders_blocks: HashMap::new(),
            placeholder_images: HashMap::new(),
        }
    }

    /// Read DOCX file
    pub fn read(&mut self) {
        self.content = read_raw_docx(&self.file_name);
    }

    /// Add placeholders to DOCX
    ///
    /// `placeholder` - Placeholder
    ///
    /// `replaced_content` - value for placeholder
    pub fn add_placeholder(&mut self, placeholder: &str, replaced_content: &str) {
        self.placeholders.insert(placeholder.to_string(), replaced_content.to_owned());
    }

    /// Add placeholders from json
    ///
    /// `placeholders_content` - json data
    pub fn add_placeholders_from_json(&mut self, placeholders_content: &str) {
        let v: Value = serde_json::from_str(placeholders_content).unwrap();
        if let Value::Object(map) = v {
            process_json_map(self,"", &map);
        }
    }

    /// Add image placeholder
    ///
    /// `placeholder` id of image in docx
    /// `image` directory of image for replace
    pub fn add_image_placeholder(&mut self, placeholder: &str, image: &str) {
        let media_dir = "word/media/";
        self.placeholder_images.insert(format!("{}{}", media_dir, placeholder.to_string()), image.to_string());
    }

    /// Remove image placeholder
    pub fn remove_image_placeholder(&mut self, placeholder: &str) {
        let media_dir = "word/media/";
        self.placeholder_images.remove(&format!("{}{}", media_dir, placeholder));
    }

    /// Init placeholders
    pub fn init_placeholders(&mut self) {
        let mut new_content = HashMap::new();

        for (k, v) in &self.content {
            let mut xml_writer = Writer::new(Cursor::new(Vec::new()));
            let mut reader = quick_xml::Reader::from_str(v);
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
                                process_text(&self, &mut current_placeholder);

                                xml_writer.write_event(Event::Text(BytesText::new(&current_placeholder))).unwrap();
                                current_placeholder.clear();
                            }
                        } else {
                            process_text(&self, &mut text);
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
            new_content.insert(k.clone(), result_str);
        }

        self.content = new_content;
    }


    /// save output file
    ///
    /// `output` - output
    pub fn save(&self, output: &str) {
        use zip::{write::FileOptions, ZipWriter};

        let media_dir = "word/media/";
        let input_file = File::open(&self.file_name).unwrap();
        let mut archive = ZipArchive::new(input_file).unwrap();

        let output_file = File::create(output).unwrap();
        let mut zip = ZipWriter::new(output_file);
        let options = FileOptions::default();
        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            if self.content.contains_key(&name) {
                zip.start_file(&name, options).unwrap();
                zip.write_all(self.content.get(&name).unwrap().as_bytes())
                    .unwrap();
            } else if name.starts_with(media_dir) && self.placeholder_images.contains_key(&name)
                && name.ends_with(".jpeg") || name.ends_with(".png") {
                let mut file = File::open(self.placeholder_images.get(&name).unwrap()).unwrap();
                let mut file_data = Vec::new();
                file.read_to_end(&mut file_data).unwrap();
                zip.start_file(&name, options).unwrap();
                zip.write_all(&*file_data).expect("panic!");
            }
            else {
                zip.start_file(&name, options).unwrap();
                std::io::copy(&mut file, &mut zip).unwrap();
            }
        }

        zip.finish().unwrap();
    }
}
