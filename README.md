# Excel to JSON Converter

CLI tool สำหรับแปลงไฟล์ Excel (.xlsx) เป็นไฟล์ JSON โดยอัตโนมัติ

## Features

- ✅ แปลงไฟล์ Excel (.xlsx) เป็น JSON
- ✅ เลือก sheet ที่ต้องการแปลง
- ✅ เลือกเฉพาะ column ที่ต้องการ (optional)
- ✅ **นับเฉพาะ column ที่มี header** - column ที่ซ่อนหรือไม่มี header จะไม่ถูกนับ
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
| `-c, --columns` | String | ❌ | เลือกเฉพาะ column ที่มี header (นับเฉพาะ visible columns) |

### Examples

#### 1. แปลงทุก column

```bash
excel2json data.xlsx "Sheet1" --output result.json
```

หรือ

```bash
excel2json data.xlsx "Sheet1" -o result.json
```

#### 2. แปลงเฉพาะบาง column (นับเฉพาะ visible columns)

```bash
# แปลงเฉพาะ visible column 1, 2, และ 3
excel2json data.xlsx "Sheet1" --columns 1,2,3 -o result.json
```

```bash
# แปลงเฉพาะ visible column 1, 3, และ 5
excel2json data.xlsx "Sales Data" -c 1,3,5 -o output.json
```

**หมายเหตุ:** เลข column จะนับเฉพาะ column ที่มี header เท่านั้น (column ที่ซ่อนหรือไม่มี header จะไม่ถูกนับ)

#### 3. ใช้กับ path ที่มีเว้นวรรค

```bash
excel2json "My Data.xlsx" "Sheet1" -o "output file.json"
```

### Visible Columns Detection

โปรแกรมจะนับและประมวลผลเฉพาะ column ที่มี header (row แรกไม่ว่าง):

**ตัวอย่าง Excel:**

| Name | Age | (empty) | Email | (empty) | Phone |
|------|-----|---------|-------|---------|-------|
| John | 25  |         | john@example.com |    | 123-456 |

**Visible columns ที่จะถูกนับ:** 4 columns (Name, Age, Email, Phone)
- Column 1 = Name
- Column 2 = Age  
- Column 3 = Email
- Column 4 = Phone

Column ที่ไม่มี header จะถูกข้ามไป

### Column Name Normalization

โปรแกรมจะแปลงชื่อ column (row แรก) อัตโนมัติ:

#### Basic Normalization

| Excel Header | JSON Key |
|--------------|----------|
| `First Name` | `first_name` |
| `Last Name` | `last_name` |
| `Email Address` | `email_address` |
| `Phone Number` | `phone_number` |
| `Total Amount` | `total_amount` |

#### Special Characters Handling

โปรแกรมจัดการ special characters ด้วยกฎดังนี้:

**Single Special Character (ใช้ตัวเดียว):**

| Excel Header | JSON Key |
|--------------|----------|
| `#` | `number` |
| `@` | `at` |
| `%` | `percent` |
| `$` | `usd` |
| `/` | `slash` |
| `&` | `and` |

**Special Characters ในข้อความ:**

| Character | Replacement | Example |
|-----------|-------------|---------|
| `/` | `_` | `Sales/Revenue` → `sales_revenue` |
| `&` | `_and_` | `Profit & Loss` → `profit_and_loss` |
| `@` | `_at_` | `Email@Domain` → `email_at_domain` |
| `#` | `_` | `Customer#ID` → `customer_id` |
| `%` | `_percent` | `Tax%Rate` → `tax_percent_rate` |
| `$` | `_usd` | `Price ($)` → `price_usd` |
| `()` | ลบออก | `Amount (USD)` → `amount_usd` |
| Space | `_` | `Total Sales` → `total_sales` |

**ตัวอย่างเพิ่มเติม:**

| Excel Header | JSON Key |
|--------------|----------|
| `#` | `number` |
| `@` | `at` |
| `Sales/Revenue` | `sales_revenue` |
| `Profit & Loss` | `profit_and_loss` |
| `Customer#ID` | `customer_id` |
| `Price ($)` | `price_usd` |
| `Discount %` | `discount_percent` |
| `Email@Address` | `email_at_address` |
| `Total (USD)` | `total_usd` |

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
  -c, --columns <COLUMNS>  Visible column numbers to include (comma-separated, e.g., 1,2,3). Only counts columns with non-empty headers
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
