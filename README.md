
# 🚀Easy Docx Template

Create docx templates


## Usage/Examples

Example 1:
```Rust
fn main()  {
    // 1. Loading docx file
    let mut docx = DOCX::new("example/test.docx".to_string());
    docx.read();

    // 2. Adding placeholders
    docx.add_placeholder("{{exam.title}}", "Mathexam");
    docx.add_placeholder("{{exam.variant}}", "1 variant");
    docx.add_placeholder("{{exam.subject}}", "Math");
    docx.add_placeholder("{{exam.level}}", "1-A form");

    // 3. Init placeholders
    docx.init_placeholders();

    // 4. Save our docx file
    docx.save("output.docx");
    println!("✅ File saved: output.docx")
}
```

Example 2(Loading data from json):

```Rust
fn main() {
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

    // 3. Init placeholders
    docx.init_placeholders();

    // 4. Save our docx file
    docx.save("output.docx");
    println!("✅ File saved: output.docx")
}
```
#### Warning! Image placeholders are initialized when the final file is saved.

Example 3(Add image placeholder)

```Rust
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
```

![example1](/imgs/example1.png)

# Roadmap

- Add block constructions
- optimize
- Add list and table support

# 🔗Author
Created by Dmitry Dzhugov
`morfyalt@proton.me`
