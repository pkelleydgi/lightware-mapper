# Lightware Excel Processor

A web-based tool to convert Lightware Americas price list Excel files to Q360 format.

## Features

- Client-side processing (no data uploaded to servers)
- Drag and drop file upload
- Automatic data mapping and validation
- Excludes rows with invalid cost data
- Downloads processed file with timestamp

## Data Mapping

The tool maps data from Lightware price lists to Q360 format as follows:

- **Part number** → MASTERNO
- **Product name** → PARTNO
- **Description** → DESCRIPTION
- **PSNI PARTNER COST** → STANDARDCOST
- **MSRP USD** → MSRP
- **Brand**: Always set to "Lightware"
- **TAXABLE**: Always set to "Y"
- **USETAXFLAG**: Always set to "Y"

## Usage

1. Visit the web application
2. Upload your Lightware Excel file (.xlsx format)
3. Click "Process File"
4. Download the converted Q360 format file

## Technical Details

- Built with vanilla JavaScript and the SheetJS library
- Runs entirely in the browser
- Compatible with modern web browsers
- No server-side processing required

## Deployment

This application is designed to be deployed on GitHub Pages as a static website.

## File Structure

```
├── index.html          # Main web application
├── README.md           # This file
└── excel_processor.py  # Python version (for reference)
```
