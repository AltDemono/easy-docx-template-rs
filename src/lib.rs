mod docx;
mod utils;

pub use docx::*;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_1() {
        // 1. Loading docx file
        let mut docx = DOCX::new("example/test.docx".to_string());
        docx.read();

        // 2. Adding placeholders
        docx.add_placeholders_from_json(r#"{
        "exam": {
                "level": "form 2-A",
                "variant": "1 variant",
                "title": "Math exam",
                "subject": "math"
            }
        }"#);
        
        // 4. Add image placeholder
        docx.add_image_placeholder("image1.jpeg", "example/replace_image1.png");

        // 5. Init placeholders
        docx.init_placeholders();

        // 6. Save our docx file
        docx.save("output.docx");
        println!("✅ File saved: output.docx")
    }
}
