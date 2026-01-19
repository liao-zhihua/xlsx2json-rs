use std::path::Path;
use crate::error::{Result, XlsxError};
use std::fs;

pub fn save_json<P: AsRef<Path>>(
    path: P,
    data: serde_json::Map<String, serde_json::Value>,
    pretty: bool,
) -> Result<()> {
        let json = if pretty {
            serde_json::to_string_pretty(&data)?
        } else {
            serde_json::to_string(&data)?
        };
        fs::write(path, json)?;
        Ok(())
}


pub fn ensure_output_dirs<P: AsRef<Path>>(output_dir: P) -> Result<()> {
    let server_dir = output_dir.as_ref().join("server");
    let client_dir = output_dir.as_ref().join("client");

    std::fs::create_dir_all(&server_dir)
        .map_err(|e| XlsxError::Io(e))?;
    std::fs::create_dir_all(&client_dir)
        .map_err(|e| XlsxError::Io(e))?;

    Ok(())
}

pub fn get_output_paths<P: AsRef<Path>>(
    output_dir: P,
    base_name: &str,
) -> (std::path::PathBuf, std::path::PathBuf) {
    let server_path = output_dir.as_ref().join("server").join(format!("{}.json", base_name));
    let client_path = output_dir.as_ref().join("client").join(format!("{}.json", base_name));
    (server_path, client_path)
} 