mod error;
mod config;
mod excel;
mod json;
mod utils;

use std::path::PathBuf;
use std::process;
use std::time::Instant;
use clap::Parser;
use walkdir::WalkDir;
use rayon::prelude::*;
use indicatif::{ProgressBar, ProgressStyle};
use num_cpus;
use crate::error::{Result, XlsxError};
use crate::utils::is_excel_file;

/// Excel转JSON工具
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// 输入Excel文件路径或目录
    #[arg(short, long)]
    input: PathBuf,

    /// 输出根目录
    #[arg(short, long, default_value = "output")]
    output_dir: PathBuf,

    /// 配置文件路径（TOML格式）
    #[arg(short, long)]
    config: Option<PathBuf>,

    /// 并行处理的工作线程数（默认为CPU核心数-1）
    #[arg(short, long)]
    workers: Option<usize>,

    /// 是否美化JSON输出
    #[arg(short, long, default_value = "false")]
    pretty: bool,
}

fn main() {
    if let Err(e) = run() {
        eprintln!("错误: {}", e);
        process::exit(1);
    }
}

fn run() -> Result<()> {
    // 初始化日志
    //env_logger::init();

    // 解析命令行参数
    let args = Args::parse();

    // 设置并行数
    let num_workers = args.workers.unwrap_or_else(|| {
        std::cmp::max(1, num_cpus::get() - 1)
    });
    rayon::ThreadPoolBuilder::new()
        .num_threads(num_workers)
        .build_global()
        .map_err(|e| XlsxError::Other(format!("无法初始化线程池: {}", e)))?;

    // 加载配置文件
    let config = if let Some(config_path) = args.config {
        crate::config::Config::load(config_path)?
    } else {
        crate::config::Config::default()
    };

    //println!("配置文件: {:?}", config);

    // 确保输出目录存在
    crate::json::ensure_output_dirs(&args.output_dir)?;

    // 获取所有需要处理的Excel文件
    let excel_files: Vec<PathBuf> = if args.input.is_dir() {
        WalkDir::new(&args.input)
            .into_iter()
            .filter_map(|e| e.ok())
            .filter(|e| !e.file_type().is_dir() && is_excel_file(e.path()))
            .filter(|e| {
            // 只导出 config 中配置过的文件
            e.path()
                .file_stem()
                .and_then(|s| s.to_str())
                .map(|stem| {
                    if config.file_mappings.contains_key(stem) {
                        true
                    } else {
                        println!(
                            "跳过文件 {}：未在 config 中配置",
                            e.path().display()
                        );
                        false
                    }
                })
                .unwrap_or(false)
        })
            .map(|e| e.path().to_path_buf())
            .collect()
    } else if is_excel_file(&args.input) {
        vec![args.input]
    } else {
        return Err(XlsxError::InvalidExcel(String::from("输入文件不是Excel文件")));
    };

    if excel_files.is_empty() {
        return Err(XlsxError::Other(String::from("未找到任何Excel文件")));
    }

    // 创建进度条
    let pb = ProgressBar::new(excel_files.len() as u64);
    pb.set_style(ProgressStyle::default_bar()
        .template("{spinner:.green} [{elapsed_precise}] [{bar:40.cyan/blue}] {pos}/{len} ({eta})")
        .map_err(|e| XlsxError::Other(format!("无法设置进度条样式: {}", e)))?
        .progress_chars("#>-"));

    // 创建处理器
    let processor = excel::ExcelProcessor::new(config, args.output_dir, args.pretty, pb.clone());

    // 记录开始时间
    let start_time = Instant::now();

    // 并行处理文件
    let results: Vec<Result<()>> = excel_files.par_iter()
        .map(|file| {
            let result = processor.process_file(file);
            pb.inc(1);
            result
        })
        .collect();

    // 完成进度条
    pb.finish_with_message("处理完成");

    // 统计结果
    let (success, errors): (Vec<_>, Vec<_>) = results.into_iter()
        .partition(Result::is_ok);

    // 打印统计信息
    println!("\n批量处理完成！");
    println!("成功处理文件数: {}", success.len());
    println!("失败文件数: {}", errors.len());
    println!("并行工作线程数: {}", num_workers);
    println!("总执行时间: {:?}", start_time.elapsed());

    // 如果有错误，打印错误信息
    if !errors.is_empty() {
        for err in errors {
            if let Err(e) = err {
                eprintln!("错误: {}", e);
            }
        }
        return Err(XlsxError::Other("处理过程中出现错误".to_string()));
    }

    Ok(())
} 
