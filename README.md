
# 🚀Easy Docx Template

Create docx templates

## ⚡Install to your project
`cargo add easy-docx-template`

## Changelog 0.1.4v
* Added each blocks
* Code refactoring
## Placeholder Helpers
`get array len: {{#exam.nums}} -> 3`

`set lowercase for text: {{lower exam.title}} -> "math exam"`

`set uppercase for text: {{upper exam.title}} -> "MATH EXAM"`

## Each block
```
{{#each exam.users}}

{{first}} {{last}}

{{/each}}

 |
\ /

Ivan Ivanov
Petr Petrov
```
## Usage/Examples

Example 1:
```Rust
use easy_docx_template::DOCX;

fn main()  {
    // 1. Loading docx file
    let mut docx = DOCX::new("example/test.docx".to_string());
    docx.read();

    // 2. Adding placeholders
    docx.add_placeholder("{{exam.title}}", "Math exam");
    docx.add_placeholder("{{exam.variant}}", "1 variant");
    docx.add_placeholder("{{exam.subject}}", "Math");
    docx.add_placeholder("{{exam.level}}", "1-A form");
    docx.add_placeholder("{{exam.image_subtitle}}", "Hello world!");
    docx.add_placeholder("{{exam.nums}}",vec!["1", "2", "3"]);

    // 3. Init placeholders
    docx.init_placeholders();

    // 4. Save our docx file
    docx.save("output.docx");
    println!("✅ File saved: output.docx")
}
```

Example 2(Loading data from json):

```Rust
use easy_docx_template::DOCX;

fn main() {
    // 1. Loading docx file
    let mut docx = DOCX::new("example/test.docx".to_string());
    docx.read();

    // 2. Adding placeholders
    docx.add_placeholders_from_json::<String>(r#"{
            "exam": {
                    "level": "form 2-A",
                    "variant": "1 variant",
                    "title": "Math exam",
                    "subject": "math",
                    "image_subtitle": "Hello world!",
                    "nums": ["1", "2"]
                }
            }"#);

    // 3. Init placeholders
    docx.init_placeholders();

    // 4. Save our docx file
    docx.save("output.docx");
    println!("✅ File saved: output.docx")
}
```
#### 🚨Warning! Image placeholders are initialized when the final file is saved.

Example 3(Add image placeholder)

```Rust
use easy_docx_template::DOCX;

fn test_1() {
        // 1. Loading docx file
        let mut docx: DOCX = DOCX::new("example/test.docx".to_string());
        docx.read();

        // 2. Adding placeholders
        docx.add_placeholders_from_json::<String>(r#"{
            "exam": {
                    "level": "form 2-A",
                    "variant": "1 variant",
                    "title": "Math exam",
                    "subject": "math",
                    "image_subtitle": "Hello world!",
                    "nums": ["1", "2"]
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

# 🚧 Roadmap

- Add block constructions
- optimize
- Add list and table support

# 🔗Author
Created by Dmitry Dzhugov
`morfyalt@proton.me`
