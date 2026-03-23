use calamine::{open_workbook, Reader, Xlsx};
use indicatif::ProgressBar;
use serde_json::{Map, Value};
use std::path::Path;

use crate::config::Target;
use crate::error::{Result, XlsxError};
use crate::json::save_json;
use crate::utils::{convert_value_by_type, filter_data};

pub struct ExcelProcessor {
    config: crate::config::Config,
    output_dir: std::path::PathBuf,
    pretty: bool,
    progress: ProgressBar,
}

impl ExcelProcessor {
    pub fn new(
        config: crate::config::Config,
        output_dir: std::path::PathBuf,
        pretty: bool,
        progress: ProgressBar,
    ) -> Self {
        Self {
            config,
            output_dir,
            pretty,
            progress,
        }
    }

    fn log(&self, message: impl Into<String>) {
        self.progress.println(message.into());
    }

    pub fn process_file<P: AsRef<Path>>(&self, excel_file: P) -> Result<()> {
        let excel_file = excel_file.as_ref();
        let output_name = self.config.get_output_name(
            excel_file
                .file_name()
                .and_then(|n| n.to_str())
                .ok_or_else(|| XlsxError::InvalidExcel("无效的文件名".to_string()))?,
        );

        let (server_path, client_path) = crate::json::get_output_paths(&self.output_dir, &output_name);

        let s_path = if self.config.gen_server(excel_file) {
            Some(server_path.as_path())
        } else {
            None
        };
        let c_path = if self.config.gen_client(excel_file) {
            Some(client_path.as_path())
        } else {
            None
        };

        if !crate::utils::need_regenerate(excel_file, s_path, c_path)? {
            self.log(format!("跳过文件 {}: JSON文件已是最新", excel_file.display()));
            return Ok(());
        }

        let mut workbook: Xlsx<_> = open_workbook(excel_file)
            .map_err(|e| XlsxError::Excel(format!("无法打开Excel文件: {}", e)))?;

        let sheet_name = workbook
            .sheet_names()
            .first()
            .ok_or_else(|| XlsxError::Excel("Excel文件没有工作表".to_string()))?
            .clone();

        let range = workbook
            .worksheet_range(&sheet_name)
            .map_err(|e| XlsxError::Excel(format!("读取工作表时出错: {}", e)))?;

        let rows: Vec<Vec<String>> = range
            .rows()
            .map(|row| row.iter().map(|cell| cell.to_string()).collect())
            .collect();

        if rows.len() < 4 {
            return Err(XlsxError::InvalidExcel(
                "Excel文件至少需要包含标记行、类型说明、表头和数据行".to_string(),
            ));
        }

        let marks: Vec<String> = rows[1].iter().map(|s| s.trim().to_lowercase()).collect();
        let types: Vec<String> = rows[2].iter().map(|s| s.trim().to_string()).collect();
        let headers: Vec<String> = rows[3].iter().map(|s| s.trim().to_string()).collect();

        let valid_header_count = headers
            .iter()
            .position(|h| h.is_empty())
            .unwrap_or(headers.len());

        let headers = headers[..valid_header_count].to_vec();
        let types = types[..valid_header_count].to_vec();
        let marks = marks[..valid_header_count].to_vec();

        let mut data = Map::<String, Value>::new();
        let mut has_error = false;

        for (i, row) in rows.iter().skip(4).enumerate() {
            if row.is_empty() || row.iter().all(|cell| cell.trim().is_empty()) {
                continue;
            }

            let mut row_obj = Map::<String, Value>::new();
            let mut id_value: Option<String> = None;
            let mut key_error_reported = false;

            for (j, cell) in row.iter().take(valid_header_count).enumerate() {
                match convert_value_by_type(cell, &types[j], i + 5, &headers[j]) {
                    Ok(value) => {
                        if j == 0 {
                            id_value = Some(match &value {
                                Value::Number(n) => n.to_string(),
                                Value::String(s) => s.clone(),
                                _ => {
                                    self.log(format!(
                                        "错误: 文件 {} 第 {} 行第一列无法作为 key",
                                        excel_file.display(),
                                        i + 5
                                    ));
                                    has_error = true;
                                    key_error_reported = true;
                                    continue;
                                }
                            });
                        }

                        if !value.is_null() {
                            row_obj.insert(headers[j].clone(), value);
                        }
                    }
                    Err(e) => {
                        self.log(format!("错误: 文件 {}: {}", excel_file.display(), e));
                        has_error = true;
                    }
                }
            }

            if let Some(key) = id_value {
                if data.contains_key(&key) {
                    self.log(format!(
                        "错误: 文件 {} key {} 重复，第 {} 行",
                        excel_file.display(),
                        key,
                        i + 5
                    ));
                    has_error = true;
                }
                data.insert(key, Value::Object(row_obj));
            } else {
                if !key_error_reported {
                    self.log(format!(
                        "错误: 文件 {} 第 {} 行第一列为空，无法作为 key",
                        excel_file.display(),
                        i + 5
                    ));
                    has_error = true;
                }
            }
        }

        if has_error {
            return Err(XlsxError::Excel(format!(
                "文件 {} 转换过程中出现错误，请检查上述错误信息",
                excel_file.display()
            )));
        }

        let (server_headers, server_data) = filter_data(&headers, &marks, &data, Target::Server);
        let (client_headers, client_data) = filter_data(&headers, &marks, &data, Target::Client);

        if !server_headers.is_empty() && self.config.gen_server(excel_file) {
            save_json(&server_path, server_data, self.pretty)?;
            self.log(format!("服务端数据已保存到: {}", server_path.display()));
        }

        if !client_headers.is_empty() && self.config.gen_client(excel_file) {
            save_json(&client_path, client_data, self.pretty)?;
            self.log(format!("客户端数据已保存到: {}", client_path.display()));
        }

        Ok(())
    }
}
