use serde_json::Value;

pub fn len_helper(value: Value) -> usize {
    match value {
        Value::Array(arr) => arr.len(),
        Value::String(s) => s.chars().count(),
        Value::Object(obj) => obj.len(),
        _ => 0
    }
}

pub fn lower_helper(value: Value) -> Value {
    match value {
        Value::String(s) => Value::String(s.to_lowercase()),
        Value::Array(arr) => Value::Array(arr.into_iter().map(lower_helper).collect()),
        _ => Value::from("")
    }
}

pub fn upper_helper(value: Value) -> Value {
    match value {
        Value::String(s) => Value::String(s.to_uppercase()),
        Value::Array(arr) => Value::Array(arr.into_iter().map(upper_helper).collect()),
        _ => Value::from("")
    }
}