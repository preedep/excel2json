# Excel to JSON Converter

CLI tool สำหรับแปลงไฟล์ Excel (.xlsx) เป็นไฟล์ JSON โดยอัตโนมัติ

## Features

- ✅ แปลงไฟล์ Excel (.xlsx) เป็น JSON
- ✅ เลือก sheet ที่ต้องการแปลง
- ✅ เลือกเฉพาะ column ที่ต้องการ (optional)
- ✅ แปลงชื่อ column อัตโนมัติ:
  - ตัวพิมพ์ใหญ่ → ตัวพิมพ์เล็ก
  - เว้นวรรค → underscore (_)
- ✅ รองรับ data types: ตัวเลข, ข้อความ, boolean

## Installation

### ติดตั้งจาก Source

```bash
# Clone repository
git clone <repository-url>
cd excel2json

# Build
cargo build --release

# Binary จะอยู่ที่
./target/release/excel2json
```

### ติดตั้งด้วย Cargo

```bash
cargo install --path .
```

## Usage

### Basic Syntax

```bash
excel2json <FILE> <SHEET> --output <OUTPUT>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `<FILE>` | String | ✅ | ไฟล์ Excel ที่ต้องการแปลง (.xlsx) |
| `<SHEET>` | String | ✅ | ชื่อ sheet ที่ต้องการแปลง |
| `-o, --output` | String | ✅ | ชื่อไฟล์ output (.json) |
| `-c, --columns` | String | ❌ | เลือกเฉพาะ column ที่ต้องการ (comma-separated) |

### Examples

#### 1. แปลงทุก column

```bash
excel2json data.xlsx "Sheet1" --output result.json
```

หรือ

```bash
excel2json data.xlsx "Sheet1" -o result.json
```

#### 2. แปลงเฉพาะบาง column

```bash
# แปลงเฉพาะ column 1, 2, และ 3
excel2json data.xlsx "Sheet1" --columns 1,2,3 -o result.json
```

```bash
# แปลงเฉพาะ column 1, 3, และ 5
excel2json data.xlsx "Sales Data" -c 1,3,5 -o output.json
```

#### 3. ใช้กับ path ที่มีเว้นวรรค

```bash
excel2json "My Data.xlsx" "Sheet1" -o "output file.json"
```

### Column Name Normalization

โปรแกรมจะแปลงชื่อ column (row แรก) อัตโนมัติ:

| Excel Header | JSON Key |
|--------------|----------|
| `First Name` | `first_name` |
| `Last Name` | `last_name` |
| `Email Address` | `email_address` |
| `Phone Number` | `phone_number` |
| `Total Amount` | `total_amount` |

### Input/Output Example

**Excel File (data.xlsx):**

| Name | Age | Email Address |
|------|-----|---------------|
| John | 25 | john@example.com |
| Jane | 30 | jane@example.com |

**Command:**

```bash
excel2json data.xlsx "Sheet1" -o output.json
```

**Output (output.json):**

```json
[
  {
    "name": "John",
    "age": 25,
    "email_address": "john@example.com"
  },
  {
    "name": "Jane",
    "age": 30,
    "email_address": "jane@example.com"
  }
]
```

## Help

ดูคำสั่งทั้งหมด:

```bash
excel2json --help
```

Output:

```
Convert Excel files to JSON format

Usage: excel2json [OPTIONS] --output <OUTPUT> <FILE> <SHEET>

Arguments:
  <FILE>   Input Excel file path (.xlsx)
  <SHEET>  Sheet name to convert

Options:
  -c, --columns <COLUMNS>  Column numbers to include (comma-separated, e.g., 1,2,3)
  -o, --output <OUTPUT>    Output JSON file path
  -h, --help               Print help
```

## Error Handling

โปรแกรมจะแสดง error message ที่ชัดเจนเมื่อเกิดปัญหา:

- ไฟล์ Excel ไม่พบ
- Sheet ที่ระบุไม่มีในไฟล์
- Column number ไม่ถูกต้อง
- ไม่สามารถสร้างไฟล์ output ได้

## Requirements

- Rust 2024 edition หรือใหม่กว่า
- ไฟล์ Excel ต้องเป็นรูปแบบ .xlsx

## Dependencies

- `calamine` - อ่านไฟล์ Excel
- `clap` - จัดการ CLI arguments
- `serde_json` - สร้าง JSON output
- `anyhow` - จัดการ errors

## License

MIT
