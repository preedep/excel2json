use anyhow::{Context, Result};
use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use serde_json::{json, Value};
use std::fs::File;
use std::io::Write;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(name = "excel2json")]
#[command(about = "Convert Excel files to JSON format", long_about = None)]
struct Args {
    #[arg(help = "Input Excel file path (.xlsx)")]
    file: PathBuf,

    #[arg(help = "Sheet name to convert")]
    sheet: String,

    #[arg(short, long, help = "Visible column numbers to include (comma-separated, e.g., 1,2,3). Only counts columns with non-empty headers. If not specified, all visible columns are included")]
    columns: Option<String>,

    #[arg(short, long, help = "Output JSON file path")]
    output: PathBuf,
}

fn normalize_column_name(name: &str) -> String {
    let trimmed = name.trim();
    
    let result = match trimmed {
        "#" => "number".to_string(),
        "@" => "at".to_string(),
        "%" => "percent".to_string(),
        "$" => "usd".to_string(),
        "/" => "slash".to_string(),
        "&" => "and".to_string(),
        _ => {
            trimmed
                .to_lowercase()
                .replace(" & ", "_and_")
                .replace("&", "_and_")
                .replace("/", "_")
                .replace("@", "_at_")
                .replace("#", "_")
                .replace("%", "_percent")
                .replace("$", "_usd")
                .replace("(", "")
                .replace(")", "")
                .replace(" ", "_")
        }
    };
    
    result
        .split('_')
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>()
        .join("_")
}

fn get_visible_column_indices(header_row: &[calamine::Data]) -> Vec<usize> {
    header_row
        .iter()
        .enumerate()
        .filter_map(|(idx, cell)| {
            let cell_str = cell.to_string().trim().to_string();
            if !cell_str.is_empty() {
                Some(idx)
            } else {
                None
            }
        })
        .collect()
}

fn parse_visible_column_numbers(
    columns_str: &str,
    visible_indices: &[usize],
) -> Result<Vec<usize>> {
    columns_str
        .split(',')
        .map(|s| {
            s.trim()
                .parse::<usize>()
                .context("Invalid column number")
                .and_then(|n| {
                    if n == 0 {
                        anyhow::bail!("Column numbers must be greater than 0")
                    }
                    if n > visible_indices.len() {
                        anyhow::bail!(
                            "Column number {} exceeds visible column count ({})",
                            n,
                            visible_indices.len()
                        )
                    }
                    Ok(visible_indices[n - 1])
                })
        })
        .collect()
}

fn read_excel_sheet(file: &PathBuf, sheet: &str) -> Result<calamine::Range<calamine::Data>> {
    let mut workbook: Xlsx<_> = open_workbook(file)
        .context(format!("Failed to open Excel file: {:?}", file))?;

    workbook
        .worksheet_range(sheet)
        .context(format!("Sheet '{}' not found", sheet))
}

fn extract_headers(
    header_row: &[calamine::Data],
    column_indices: &[usize],
) -> Vec<String> {
    column_indices
        .iter()
        .map(|&i| {
            header_row
                .get(i)
                .map(|cell| normalize_column_name(&cell.to_string()))
                .unwrap_or_else(|| format!("column_{}", i + 1))
        })
        .collect()
}

fn convert_cell_to_json(cell: &calamine::Data) -> Value {
    json!(cell.to_string())
}

fn convert_rows_to_json<'a>(
    rows: impl Iterator<Item = &'a [calamine::Data]>,
    headers: &[String],
    column_indices: &[usize],
) -> Vec<Value> {
    rows.map(|row| {
        let json_obj: serde_json::Map<String, Value> = column_indices
            .iter()
            .enumerate()
            .map(|(header_idx, &col_idx)| {
                let value = row
                    .get(col_idx)
                    .map(convert_cell_to_json)
                    .unwrap_or(json!(null));
                (headers[header_idx].clone(), value)
            })
            .collect();
        json!(json_obj)
    })
    .collect()
}

fn write_json_to_file(json_array: &[Value], output: &PathBuf) -> Result<()> {
    let json_output = serde_json::to_string_pretty(json_array)
        .context("Failed to serialize JSON")?;

    let mut file = File::create(output)
        .context(format!("Failed to create output file: {:?}", output))?;

    file.write_all(json_output.as_bytes())
        .context("Failed to write to output file")?;

    Ok(())
}

fn main() -> Result<()> {
    let args = Args::parse();

    let range = read_excel_sheet(&args.file, &args.sheet)?;
    let mut rows = range.rows();

    let header_row = rows
        .next()
        .context("Excel sheet is empty, no header row found")?;

    let visible_indices = get_visible_column_indices(header_row);

    let column_indices: Vec<usize> = if let Some(ref cols_str) = args.columns {
        parse_visible_column_numbers(cols_str, &visible_indices)?
    } else {
        visible_indices
    };

    let headers = extract_headers(header_row, &column_indices);
    let json_array = convert_rows_to_json(rows, &headers, &column_indices);

    write_json_to_file(&json_array, &args.output)?;

    println!("Successfully converted Excel to JSON");
    println!("Input: {:?}", args.file);
    println!("Sheet: {}", args.sheet);
    println!("Output: {:?}", args.output);
    println!("Visible columns: {}", column_indices.len());
    println!("Total records: {}", json_array.len());

    Ok(())
}
