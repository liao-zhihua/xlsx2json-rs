use serde::Deserialize;
use std::path::Path;
use crate::error::{Result, XlsxError};
use std::collections::HashMap;

#[derive(Debug, Deserialize, Clone, Copy)]
#[serde(rename_all = "PascalCase")]
pub enum Target {
    Server,
    Client,
    Both,
}

impl Default for Target {
    fn default() -> Self {
        Target::Both
    }
}

#[derive(Debug, Deserialize)]
pub struct FileMapping {
    pub output: String,

    #[serde(default)]
    pub target: Target,
}


#[derive(Debug, Deserialize, Default)]
pub struct Config {
    #[serde(default)]
    pub file_mappings: HashMap<String, FileMapping>,
}

impl Config {
    pub fn load<P: AsRef<Path>>(path: P) -> Result<Self> {
        let content = std::fs::read_to_string(path)
            .map_err(|e| XlsxError::Config(format!("无法读取配置文件: {}", e)))?;
        
        toml::from_str(&content)
            .map_err(|e| XlsxError::Config(format!("无法解析配置文件: {}", e)))
    }

    pub fn get_output_name(&self, excel_file: &str) -> String {
        // 获取Excel文件名（不含路径和扩展名）
        let base_name = std::path::Path::new(excel_file)
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or(excel_file);

        // 如果配置中有映射，使用映射的名称
        self.file_mappings
            .get(base_name)
            .map(|m| m.output.clone())
            .unwrap_or_else(|| base_name.to_string())
    }


    pub fn gen_server<P: AsRef<Path>>(&self, excel_file: P) -> bool {
        self.get_target(excel_file)
            .map(|t| matches!(t, Target::Server | Target::Both))
            .unwrap_or(true)
    }

    pub fn gen_client<P: AsRef<Path>>(&self, excel_file: P) -> bool {
        self.get_target(excel_file)
            .map(|t| matches!(t, Target::Client | Target::Both))
            .unwrap_or(true)
    }

    fn get_target<P: AsRef<Path>>(&self, excel_file: P) -> Option<Target> {
        let base_name = excel_file
            .as_ref()
            .file_stem()
            .and_then(|s| s.to_str())?;

        self.file_mappings
            .get(base_name)
            .map(|m| m.target)
    }

    // pub fn gen_server(&self, excel_file: &str) -> bool {
    //     self.get_target(excel_file)
    //         .map(|t| matches!(t, Target::Server | Target::Both))
    //         .unwrap_or(true)
    // }
    //
    // pub fn gen_client(&self, excel_file: &str) -> bool {
    //     self.get_target(excel_file)
    //         .map(|t| matches!(t, Target::Client | Target::Both))
    //         .unwrap_or(true)
    // }
    //
    // fn get_target(&self, excel_file: &str) -> Option<Target> {
    //     let base_name = std::path::Path::new(excel_file)
    //         .file_stem()
    //         .and_then(|s| s.to_str())?;
    //
    //     self.file_mappings
    //         .get(base_name)
    //         .map(|m| m.target)
    // }
}

// impl Default for Config {
//     fn default() -> Self {
//         Self {
//             file_mappings: std::collections::HashMap::new(),
//         }
//     }
// }