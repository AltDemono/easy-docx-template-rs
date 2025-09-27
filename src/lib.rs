mod docx;
mod utils;

pub use docx::*;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_with_json() {
        // 1. Loading docx file
        let mut docx: DOCX = DOCX::new("example/test.docx".to_string());
        docx.read();

        // 2. Adding placeholders
        docx.add_placeholders_from_json::<String>(r#"{
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

    #[test]
    fn test_default() {
        // 1. Loading docx file
        let mut docx = DOCX::new("example/test.docx".to_string());
        docx.read();

        // 2. Adding placeholders
        docx.add_placeholder("{{exam.title}}", "Math exam");
        docx.add_placeholder("{{exam.variant}}", "1 variant");
        docx.add_placeholder("{{exam.subject}}", "Math");
        docx.add_placeholder("{{exam.level}}", "1-A form");
        docx.add_placeholder("{{example}}", "Hello world!");

        // 3. Init placeholders
        docx.init_placeholders();

        // 4. Save our docx file
        docx.save("output.docx");
        println!("✅ File saved: output.docx")
    }
}
