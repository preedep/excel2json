// External dependencies
use anyhow::{Context, Result}; // Error handling with context
use calamine::{open_workbook, Reader, Xlsx}; // Excel file reading library
use clap::Parser; // Command-line argument parser
use serde_json::{json, Value}; // JSON serialization
use std::fs::File; // File system operations
use std::io::Write; // Write trait for file output
use std::path::PathBuf; // Cross-platform file path handling

/// Command-line arguments structure
/// Defines all parameters that users can pass to the CLI tool
#[derive(Parser, Debug)]
#[command(name = "excel2json")]
#[command(about = "Convert Excel files to JSON format", long_about = None)]
struct Args {
    /// Path to the input Excel file (.xlsx format)
    #[arg(help = "Input Excel file path (.xlsx)")]
    file: PathBuf,

    /// Name of the sheet within the Excel file to convert
    #[arg(help = "Sheet name to convert")]
    sheet: String,

    /// Optional: Comma-separated list of visible column numbers to include
    /// Only columns with non-empty headers are counted
    /// Example: "1,2,3" will include the first three visible columns
    #[arg(short, long, help = "Visible column numbers to include (comma-separated, e.g., 1,2,3). Only counts columns with non-empty headers. If not specified, all visible columns are included")]
    columns: Option<String>,

    /// Path where the output JSON file will be saved
    #[arg(short, long, help = "Output JSON file path")]
    output: PathBuf,
}

/// Normalizes Excel column header names to valid JSON keys
/// 
/// Rules:
/// - Single special characters are converted to meaningful words (e.g., "#" -> "number")
/// - Converts to lowercase
/// - Replaces special characters with underscores or meaningful text
/// - Removes parentheses
/// - Removes consecutive underscores
/// 
/// # Arguments
/// * `name` - The original column header name from Excel
/// 
/// # Returns
/// A normalized string suitable for use as a JSON key
/// 
/// # Examples
/// - "First Name" -> "first_name"
/// - "#" -> "number"
/// - "Sales/Revenue" -> "sales_revenue"
/// - "Profit & Loss" -> "profit_and_loss"
fn normalize_column_name(name: &str) -> String {
    let trimmed = name.trim();
    
    // Handle single special characters with meaningful names
    let result = match trimmed {
        "#" => "number".to_string(),
        "@" => "at".to_string(),
        "%" => "percent".to_string(),
        "$" => "usd".to_string(),
        "/" => "slash".to_string(),
        "&" => "and".to_string(),
        _ => {
            // For all other cases, apply transformation rules
            trimmed
                .to_lowercase() // Convert to lowercase
                .replace(" & ", "_and_") // Replace " & " with "_and_"
                .replace("&", "_and_") // Replace "&" with "_and_"
                .replace("/", "_") // Replace "/" with "_"
                .replace("@", "_at_") // Replace "@" with "_at_"
                .replace("#", "_") // Replace "#" with "_"
                .replace("%", "_percent") // Replace "%" with "_percent"
                .replace("$", "_usd") // Replace "$" with "_usd"
                .replace("(", "") // Remove opening parenthesis
                .replace(")", "") // Remove closing parenthesis
                .replace(" ", "_") // Replace spaces with underscores
        }
    };
    
    // Clean up: remove consecutive underscores and empty segments
    result
        .split('_')
        .filter(|s| !s.is_empty()) // Remove empty segments
        .collect::<Vec<_>>()
        .join("_") // Join with single underscore
}

/// Identifies visible columns by filtering out columns with empty headers
/// 
/// This function helps distinguish between actual data columns and hidden/unused columns.
/// Only columns with non-empty header values are considered "visible".
/// 
/// # Arguments
/// * `header_row` - The first row of the Excel sheet containing column headers
/// 
/// # Returns
/// A vector of column indices (0-based) that have non-empty headers
/// 
/// # Example
/// If header row is: ["Name", "Age", "", "Email", "", "Phone"]
/// Returns: [0, 1, 3, 5] (indices of non-empty columns)
fn get_visible_column_indices(header_row: &[calamine::Data]) -> Vec<usize> {
    header_row
        .iter() // Iterate through all cells in the header row
        .enumerate() // Get index along with each cell
        .filter_map(|(idx, cell)| {
            // Convert cell to string and trim whitespace
            let cell_str = cell.to_string().trim().to_string();
            // Only include columns with non-empty headers
            if !cell_str.is_empty() {
                Some(idx) // Return the column index
            } else {
                None // Skip empty columns
            }
        })
        .collect() // Collect all visible column indices into a vector
}

/// Parses user-specified column numbers and maps them to actual visible column indices
/// 
/// Users specify columns using 1-based numbering (1, 2, 3, ...)
/// This function converts those to 0-based indices and validates them.
/// 
/// # Arguments
/// * `columns_str` - Comma-separated string of column numbers (e.g., "1,2,3")
/// * `visible_indices` - Vector of actual column indices that have non-empty headers
/// 
/// # Returns
/// A Result containing a vector of actual column indices to use
/// 
/// # Errors
/// - Returns error if column number is 0 or negative
/// - Returns error if column number exceeds the count of visible columns
/// - Returns error if the input string contains invalid numbers
/// 
/// # Example
/// If visible_indices = [0, 2, 5, 7] and columns_str = "1,3"
/// Returns: Ok([0, 5]) - maps user's 1st and 3rd visible columns to actual indices
fn parse_visible_column_numbers(
    columns_str: &str,
    visible_indices: &[usize],
) -> Result<Vec<usize>> {
    columns_str
        .split(',') // Split by comma
        .map(|s| {
            s.trim() // Remove whitespace
                .parse::<usize>() // Parse string to number
                .context("Invalid column number") // Add error context
                .and_then(|n| {
                    // Validate column number
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
                    // Convert 1-based user input to 0-based array index
                    // Then map to actual column index in the Excel sheet
                    Ok(visible_indices[n - 1])
                })
        })
        .collect() // Collect all results, will fail if any parsing failed
}

/// Opens an Excel file and reads a specific worksheet
/// 
/// # Arguments
/// * `file` - Path to the Excel file (.xlsx)
/// * `sheet` - Name of the worksheet to read
/// 
/// # Returns
/// A Result containing the Range of cells from the specified worksheet
/// 
/// # Errors
/// - Returns error if the file cannot be opened
/// - Returns error if the specified sheet name doesn't exist in the workbook
fn read_excel_sheet(file: &PathBuf, sheet: &str) -> Result<calamine::Range<calamine::Data>> {
    // Open the Excel workbook
    let mut workbook: Xlsx<_> = open_workbook(file)
        .context(format!("Failed to open Excel file: {:?}", file))?;

    // Get the specified worksheet range (all cells with data)
    workbook
        .worksheet_range(sheet)
        .context(format!("Sheet '{}' not found", sheet))
}

/// Extracts and normalizes column headers for the specified column indices
/// 
/// # Arguments
/// * `header_row` - The first row containing column headers
/// * `column_indices` - Vector of column indices to extract headers from
/// 
/// # Returns
/// A vector of normalized header names suitable for use as JSON keys
/// 
/// # Behavior
/// - Normalizes each header using normalize_column_name()
/// - If a column index is out of bounds, generates a default name "column_N"
fn extract_headers(
    header_row: &[calamine::Data],
    column_indices: &[usize],
) -> Vec<String> {
    column_indices
        .iter() // Iterate through selected column indices
        .map(|&i| {
            header_row
                .get(i) // Try to get the cell at this index
                .map(|cell| normalize_column_name(&cell.to_string())) // Normalize if found
                .unwrap_or_else(|| format!("column_{}", i + 1)) // Fallback name if not found
        })
        .collect() // Collect into a vector of strings
}

/// Converts an Excel cell value to a JSON value
/// 
/// Currently converts all cell values to strings to preserve formatting
/// and handle cases where numbers represent identifiers (like bullet numbers)
/// rather than numeric values.
/// 
/// # Arguments
/// * `cell` - Reference to a cell from the Excel sheet
/// 
/// # Returns
/// A serde_json::Value representing the cell content as a string
fn convert_cell_to_json(cell: &calamine::Data) -> Value {
    // Convert all values to strings to preserve formatting
    // This is useful for bullet numbers, IDs, and other non-numeric data
    json!(cell.to_string())
}

/// Converts Excel rows to JSON objects
/// 
/// Each row becomes a JSON object where keys are the normalized column headers
/// and values are the cell contents.
/// 
/// # Arguments
/// * `rows` - Iterator over Excel rows (excluding the header row)
/// * `headers` - Vector of normalized column header names
/// * `column_indices` - Vector of column indices to include in the output
/// 
/// # Returns
/// A vector of JSON values, where each value is an object representing one row
/// 
/// # Example
/// Input row: ["John", "25", "john@example.com"]
/// Headers: ["name", "age", "email"]
/// Output: {"name": "John", "age": "25", "email": "john@example.com"}
fn convert_rows_to_json<'a>(
    rows: impl Iterator<Item = &'a [calamine::Data]>,
    headers: &[String],
    column_indices: &[usize],
) -> Vec<Value> {
    rows.map(|row| {
        // Create a JSON object for this row
        let json_obj: serde_json::Map<String, Value> = column_indices
            .iter() // Iterate through selected columns
            .enumerate() // Get index for matching with headers
            .map(|(header_idx, &col_idx)| {
                // Get cell value or use null if cell doesn't exist
                let value = row
                    .get(col_idx) // Try to get the cell at this column index
                    .map(convert_cell_to_json) // Convert to JSON if found
                    .unwrap_or(json!(null)); // Use null if cell is missing
                // Create key-value pair: (header_name, cell_value)
                (headers[header_idx].clone(), value)
            })
            .collect(); // Collect into a Map
        json!(json_obj) // Convert Map to JSON Value
    })
    .collect() // Collect all row objects into a vector
}

/// Writes JSON data to a file with pretty formatting
/// 
/// # Arguments
/// * `json_array` - Array of JSON values to write
/// * `output` - Path where the JSON file should be created
/// 
/// # Returns
/// Result indicating success or failure
/// 
/// # Errors
/// - Returns error if JSON serialization fails
/// - Returns error if file cannot be created
/// - Returns error if writing to file fails
fn write_json_to_file(json_array: &[Value], output: &PathBuf) -> Result<()> {
    // Serialize JSON array to a pretty-printed string
    let json_output = serde_json::to_string_pretty(json_array)
        .context("Failed to serialize JSON")?;

    // Create the output file (overwrites if exists)
    let mut file = File::create(output)
        .context(format!("Failed to create output file: {:?}", output))?;

    // Write the JSON string to the file
    file.write_all(json_output.as_bytes())
        .context("Failed to write to output file")?;

    Ok(())
}

/// Main entry point for the Excel to JSON converter
/// 
/// Process flow:
/// 1. Parse command-line arguments
/// 2. Open Excel file and read specified sheet
/// 3. Identify visible columns (non-empty headers)
/// 4. Parse user-specified column selection (if provided)
/// 5. Extract and normalize column headers
/// 6. Convert all data rows to JSON objects
/// 7. Write JSON output to file
/// 8. Display summary statistics
/// 
/// # Returns
/// Result indicating success or failure of the conversion process
fn main() -> Result<()> {
    // Step 1: Parse command-line arguments
    let args = Args::parse();

    // Step 2: Open Excel file and read the specified sheet
    let range = read_excel_sheet(&args.file, &args.sheet)?;
    let mut rows = range.rows();

    // Step 3: Extract the header row (first row)
    let header_row = rows
        .next() // Get first row
        .context("Excel sheet is empty, no header row found")?;

    // Step 4: Identify which columns have non-empty headers (visible columns)
    let visible_indices = get_visible_column_indices(header_row);

    // Step 5: Determine which columns to include in the output
    // Either use user-specified columns or all visible columns
    let column_indices: Vec<usize> = if let Some(ref cols_str) = args.columns {
        // User specified specific columns - parse and validate them
        parse_visible_column_numbers(cols_str, &visible_indices)?
    } else {
        // No columns specified - use all visible columns
        visible_indices
    };

    // Step 6: Extract and normalize the column headers
    let headers = extract_headers(header_row, &column_indices);
    
    // Step 7: Convert all data rows to JSON objects
    let json_array = convert_rows_to_json(rows, &headers, &column_indices);

    // Step 8: Write the JSON array to the output file
    write_json_to_file(&json_array, &args.output)?;

    // Step 9: Display success message and statistics
    println!("Successfully converted Excel to JSON");
    println!("Input: {:?}", args.file);
    println!("Sheet: {}", args.sheet);
    println!("Output: {:?}", args.output);
    println!("Visible columns: {}", column_indices.len());
    println!("Total records: {}", json_array.len());

    Ok(())
}
