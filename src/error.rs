use thiserror::Error;

#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("IO错误: {0}")]
    Io(#[from] std::io::Error),

    #[error("Excel文件错误: {0}")]
    Excel(String),

    #[error("JSON序列化错误: {0}")]
    Json(#[from] serde_json::Error),

    // #[error("TOML解析错误: {0}")]
    // Toml(#[from] toml::de::Error),

    #[error("类型转换错误: 第{row}行 '{header}'列: 无法将值 '{value}' 转换为 {type_name} 类型: {message}")]
    ConvertError {
        row: usize,
        header: String,
        value: String,
        type_name: String,
        message: String,
    },

    #[error("配置文件错误: {0}")]
    Config(String),

    #[error("无效的Excel文件: {0}")]
    InvalidExcel(String),

    #[error("其他错误: {0}")]
    Other(String),
}


pub type Result<T> = std::result::Result<T, XlsxError>; 
