# PPR Processing Tool

A web-based application that processes Excel files and converts them into a standardized PPR (Print Production Report) format.

## Features

- Excel file upload and processing
- Interactive data preview table
- Row deletion (single or multiple)
- Standardized PPR output format
- Automatic line number generation
- Status and result code generation

## Input Format

The tool expects an Excel file (.xlsx or .xls) with the following structure:
- Column A: Reference/Order Number
- Column C: ISBN

The reference number in the first data row (A2) will be used for the output filename.

## Output Format

Generates a file named `T1.M{reference}.PPR` with the following format:
```
reference,line_number,isbn,AR,OK
```

### Field Formatting

- **Reference**: Taken directly from Column A
- **Line Numbers**: Auto-generated, zero-padded 5-digit numbers (e.g., "00001")
- **ISBN**: Taken from Column C
- **Status**: Fixed value "AR"
- **Result**: Fixed value "OK"

Example output line:
```
7000627211,00001,9780367180133,AR,OK
```

## Usage Instructions

1. Open the application in a web browser
2. Click "Choose File" to select your Excel file
3. Review the processed data in the preview table
4. Use actions available in the preview table:
   - Delete individual rows using the trash icon
   - Select multiple rows using checkboxes and delete them
   - Use "Clear All" to reset the form
5. Click "Download PPR" to generate and download the formatted PPR file
6. The Excel is downloaded from the following report in PACE - Search for the order/reference number
http://192.168.10.251/epace/company:c001/inquiry/UserDefinedInquiry/view/5283?