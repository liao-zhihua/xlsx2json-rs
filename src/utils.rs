use std::path::Path;
use crate::error::{Result, XlsxError};
use serde_json::{Map, Value};
use crate::config::Target;


pub fn is_excel_file<P: AsRef<Path>>(path: P) -> bool {
    let path = path.as_ref();
    if let Some(file_name) = path.file_name() {
        if let Some(name) = file_name.to_str() {
            // 检查是否是Excel临时文件（以~$开头）
            if name.starts_with("~$") {
                return false;
            }
            // 检查是否是xlsx文件
            return name.to_lowercase().ends_with(".xlsx");
        }
    }
    false
}

pub fn convert_value_by_type(
    value: &str,
    type_str: &str,
    row: usize,
    header: &str,
) -> Result<serde_json::Value> {
    // 如果是空字符串，返回null
    if value.trim().is_empty() {
        return Ok(serde_json::Value::Null);
    }

    // 根据类型字符串进行转换
    match type_str.trim().to_lowercase().as_str() {
        "int" | "integer" => {
            value.parse::<i64>()
                .map(|n| serde_json::Value::Number(serde_json::Number::from(n)))
                .map_err(|_| XlsxError::ConvertError {
                    row,
                    header: header.to_string(),
                    value: value.to_string(),
                    type_name: "int".to_string(),
                    message: "不是有效的整数".to_string(),
                })
        }
        "float" | "double" | "number" => {
            match value.parse::<f64>() {
                Ok(n) => {
                    serde_json::Number::from_f64(n)
                        .map(serde_json::Value::Number)
                        .ok_or_else(|| XlsxError::ConvertError {
                            row,
                            header: header.to_string(),
                            value: value.to_string(),
                            type_name: "float".to_string(),
                            message: "不是有效的浮点数".to_string(),
                        })
                }
                Err(_) => Err(XlsxError::ConvertError {
                    row,
                    header: header.to_string(),
                    value: value.to_string(),
                    type_name: "float".to_string(),
                    message: "不是有效的浮点数".to_string(),
                })
            }
        }
        "bool" | "boolean" => {
            value.parse::<bool>()
                .map(serde_json::Value::Bool)
                .map_err(|_| XlsxError::ConvertError {
                    row,
                    header: header.to_string(),
                    value: value.to_string(),
                    type_name: "bool".to_string(),
                    message: "不是有效的布尔值".to_string(),
                })
        }
        "json" | "array" | "object" => {
            serde_json::from_str(value)
                .map_err(|_| XlsxError::ConvertError {
                    row,
                    header: header.to_string(),
                    value: value.to_string(),
                    type_name: "json".to_string(),
                    message: "不是有效的JSON格式".to_string(),
                })
        }
        "string" | "str" | "text" => Ok(serde_json::Value::String(value.to_string())),
        _ => Err(XlsxError::ConvertError {
            row,
            header: header.to_string(),
            value: value.to_string(),
            type_name: type_str.to_string(),
            message: "不支持的数据类型".to_string(),
        }),
    }
}

pub fn need_regenerate(
    excel_file: &Path,
    server_file: Option<&Path>,
    client_file: Option<&Path>,
) -> Result<bool> {
    let excel_mod_time = std::fs::metadata(excel_file)?.modified()?;

    // 检查逻辑：
    // 1. 如果 Option 是 None，说明这一端根本不需要，返回 false (不需要重新生成)
    // 2. 如果 Option 是 Some，检查文件是否存在及时间戳
    let check = |f: Option<&Path>| -> bool {
        f.map(|path| {
            std::fs::metadata(path)
                .map(|meta| meta.modified().unwrap_or(std::time::SystemTime::UNIX_EPOCH) < excel_mod_time)
                .unwrap_or(true) // 文件不存在，返回 true (需要重新生成)
        }).unwrap_or(false) // 配置不需要生成，返回 false
    };

    Ok(check(server_file) || check(client_file))
}

pub fn filter_data(
    headers: &[String],
    // types: &[String],
    marks: &[String],
    data: &Map<String, Value>,
    target: Target,
) -> (Vec<String>, Map::<String, Value>) {
    let mut filtered_headers = Vec::new();
    // let mut filtered_types = Vec::new();
    let mut filtered_data = Map::<String, Value>::new();

    // 创建列索引映射
    // let valid_indices: Vec<usize> = marks
    //     .iter()
    //     .enumerate()
    //     .filter_map(|(i, mark)| {
    //         let mark = mark.trim().to_lowercase();
    //         if mark == "b" || mark == target {
    //             Some(i)
    //         } else {
    //             None
    //         }
    //     })
    //     .collect();

    let valid_indices: Vec<usize> = marks
        .iter()
        .enumerate()
        .filter_map(|(i, mark)| {
            match mark.trim().to_lowercase().as_str() {
                "b" => Some(i),
                "s" if matches!(target, Target::Server | Target::Both) => Some(i),
                "c" if matches!(target, Target::Client | Target::Both) => Some(i),
                _ => None,
            }
        })
        .collect();

    // 过滤表头和类型
    for &idx in &valid_indices {
        filtered_headers.push(headers[idx].clone());
        // filtered_types.push(types[idx].clone());
    }

    // 过滤数据
    for (key, row_value) in data.iter() {
        if let Value::Object(row_map) = row_value {
            let mut new_row = Map::new();
            for &idx in &valid_indices {
                let header = &headers[idx];
                if let Some(v) = row_map.get(header) {
                    new_row.insert(header.clone(), v.clone());
                } 
                // else {
                //     new_row.insert(header.clone(), Value::Null);
                // }
            }
            filtered_data.insert(key.clone(), Value::Object(new_row));
        }
    }

    // (filtered_headers, filtered_types, filtered_data)
    (filtered_headers, filtered_data)
}
