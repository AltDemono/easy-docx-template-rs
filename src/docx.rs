use std::collections::HashMap;
use serde_json::Value;
use crate::utils::*;

const PLACEHOLDER_START: &str = "{{";
const PLACEHOLDER_END: &str = "}}";
static WORD_PARAGRAPH_TAG: &[u8] = b"w:p";
const PLACEHOLDER_BLOCK_START: &str = "{%";
const PLACEHOLDER_BLOCK_END: &str = "%}";

/// # DOCX struct
///
/// create template for docx file
///
/// `file_name` - docx file input name
///
/// `content` - content of word file
///
/// `placeholder`/`placeholder_block`/`placeholder_images` - saving your placeholders
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
        self.placeholders.insert(placeholder.to_string(), replaced_content.to_string());
    }

    /// Add placeholders from json
    ///
    /// `placeholders_content` - json data
    pub fn add_placeholders_from_json<T: ToString + std::fmt::Display + From<String>>(&mut self, placeholders_content: &str) {
        let v: Value = serde_json::from_str(placeholders_content).unwrap();
        if let Value::Object(map) = v {
            process_json_map::<T>(self,"", &map);
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
    ///
    /// `placeholder` - placeholder tag
    pub fn remove_image_placeholder(&mut self, placeholder: &str) {
        let media_dir = "word/media/";
        self.placeholder_images.remove(&format!("{}{}", media_dir, placeholder));
    }

    /// Init placeholders
    pub fn init_placeholders(&mut self) {
        let mut new_content = HashMap::new();

        for (k, v) in &self.content {
            new_content.insert(k.to_string(), init_placeholders(&self.placeholders, v));
        }

        self.content = new_content;
    }

    /// save output file
    ///
    /// `output` - output
    pub fn save(&self, output: &str) {
        let result = render_docx(&self);
        std::fs::write(output, result).unwrap();
    }

    /// Render docx file
    ///
    /// You can use this in web
    pub fn render(&self) -> Vec<u8> {
        render_docx(&self)
    }
}
