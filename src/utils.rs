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

pub fn need_regenerate<P: AsRef<Path>>(
    excel_file: P,
    server_file: P,
    client_file: P,
) -> Result<bool> {
    let excel_info = std::fs::metadata(&excel_file)
        .map_err(|e| XlsxError::Io(e))?;
    let excel_mod_time = excel_info.modified()
        .map_err(|e| XlsxError::Io(e))?;

    // 检查服务端JSON文件
    let server_need_regen = match std::fs::metadata(&server_file) {
        Ok(server_info) => {
            let server_mod_time = server_info.modified()
                .map_err(|e| XlsxError::Io(e))?;
            server_mod_time < excel_mod_time
        }
        Err(_) => true,
    };

    // 检查客户端JSON文件
    let client_need_regen = match std::fs::metadata(&client_file) {
        Ok(client_info) => {
            let client_mod_time = client_info.modified()
                .map_err(|e| XlsxError::Io(e))?;
            client_mod_time < excel_mod_time
        }
        Err(_) => true,
    };

    Ok(server_need_regen || client_need_regen)
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