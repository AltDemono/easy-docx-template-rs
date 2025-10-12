mod docx;
mod utils;
mod placeholder_helpers;

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
        docx.add_placeholders_from_json(r#"{
        "exam": {
                "level": "form 2-A",
                "variant": "1 variant",
                "title": "Math exam",
                "subject": "math",
                "image_subtitle": "Hello world!",
                "users": [
                    { "first": "Ivan", "last": "Ivanov" },
                    { "first": "Petr", "last": "Petrov" }
                ],
                "down": "made with love",
                "test1": "2424"
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
        docx.add_placeholder("{{exam.image_subtitle}}", "Hello world!");
        docx.add_placeholder("{{exam.down}}", "made with love");

        // 3. Add image placeholder
        docx.add_image_placeholder("image1.jpeg", "example/replace_image1.png");

        // 4. Init placeholders
        docx.init_placeholders();

        // 5. Save our docx file
        docx.save("output.docx");
        println!("✅ File saved: output.docx")
    }

    #[test]
    fn test_with_each_blocks() {
        let mut docx: DOCX = DOCX::new("example/test.docx".to_string());
        docx.read();
        docx.add_placeholders_from_json(r#"{
        "exam": {
                "level": "form 2-A",
                "variant": "1 variant",
                "title": "Math exam",
                "subject": "math",
                "image_subtitle": "Hello world!",
                  "users": [
                    {
                      "name": "Alice",
                      "age": 28,
                      "pets": [
                        {
                          "type": "Cat",
                          "name": "Misty",
                          "toys": [
                            { "title": "Ball" },
                            { "title": "Mouse" }
                          ]
                        },
                        {
                          "type": "Dog",
                          "name": "Rex",
                          "toys": [
                            { "title": "Bone" }
                          ]
                        }
                      ]
                    },
                    {
                      "name": "Bob",
                      "age": 34,
                      "pets": [
                        {
                          "type": "Fish",
                          "name": "Nemo",
                          "toys": []
                        }
                      ]
                    }
                  ],
                "options": ["okak", "bubu"],
                "down": "made with love"
            }
        }"#);
        // 5. Init placeholders
        docx.init_placeholders();

        // 6. Save our docx file
        docx.save("output.docx");
        println!("✅ File saved: output.docx")
    }
}
