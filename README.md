# JD Sons Offline Order - Customer Document Generator

This website allows you to upload an Excel file containing customer data and generate formatted Word and PDF documents with customer details.

## ✨ New Features & Improvements

- **🎨 Document Format Customization**: Full control over paper size, margins, fonts, colors, and spacing
- **📄 Enhanced PDF Generation**: Improved multi-page PDF generation with proper formatting
- **📊 Google Sheets Integration**: Connect directly to Google Sheets for real-time data
- **📋 Customer Selection**: Choose specific customers using table selection or row numbers
- **🔔 Better User Feedback**: Improved status messages and error handling
- **💾 Settings Persistence**: Your format preferences are automatically saved

## Features

- **Excel File Upload**: Support for .xlsx and .xls files with drag & drop functionality
- **Google Sheets Integration**: Load data directly from Google Sheets
- **Flexible Customer Selection**: 
  - Table-based selection with checkboxes
  - Row number range selection
- **Document Generation**: Creates properly formatted Word and PDF documents
- **Multi-Page Documents**: Each customer gets their own page in the document
- **Format Customization**: 
  - Paper size presets (A4, A5, A6, Receipt, Business Card, Custom)
  - Adjustable margins, fonts, colors, and spacing
  - Real-time preview of formatting changes
- **Professional Formatting**: Clean, printable format optimized for various paper sizes

## How to Use

1. **Open the Website**: Open `index.html` in your web browser
2. **Choose Data Source**: 
   - Upload an Excel file, or
   - Connect to Google Sheets using a shareable link
3. **Customize Format** (Optional):
   - Adjust paper size, margins, fonts, and colors in the Format Settings
   - Preview changes in real-time
4. **Select Customers**: 
   - Use table selection to pick specific customers, or
   - Use row number ranges for bulk selection
5. **Generate Document**: Click "Generate Word Document" or "Generate PDF Document"
6. **Download**: The document will automatically download or open for printing

## Data Source Requirements

### Excel File Requirements
Your Excel file must contain these columns (case-insensitive):
- **Customer Name** (or "Name", "Customer", "Client")
- **Address** (or "Addr", "Location")
- **Contact Number** (or "Contact", "Phone", "Mobile", "Cell")

### Google Sheets Requirements
1. **Publish to Web**: File → Share → Publish to web → Select CSV format
2. **Share Publicly**: Make sure the sheet is accessible to anyone with the link
3. **Column Structure**: Same as Excel file requirements

### Example Data Structure:
```
| Customer Name | Address        | Contact Number | 
|---------------|----------------|----------------|
| John Doe      | 123 Main St    | 9876543210     |
| Jane Smith    | 456 Oak Ave    | 9876543211     |
```

## Document Format

Each page in the generated document contains:

```
Customer Details:

Customer Name: [Customer Name]
Address: [Customer Address]  
Contact Number: [Customer Contact Number]

─────────────────────────────────────────
From: Jemish (JD Jewellery)
Contact Number: 9773046615
```

## Format Customization

### Paper Size Options
- **Receipt**: 7.75 × 12.5 cm (default)
- **Business Card**: 8.5 × 5.5 cm
- **A6**: 10.5 × 14.8 cm
- **A5**: 14.8 × 21.0 cm
- **A4**: 21.0 × 29.7 cm
- **Custom**: Set your own dimensions

### Styling Options
- **Fonts**: Arial, Times New Roman, Helvetica, Calibri, Georgia
- **Font Sizes**: Separate settings for headers and body text
- **Colors**: Customizable header, text, and company info colors
- **Spacing**: Adjustable line spacing and margins

## Technical Requirements

- Modern web browser (Chrome, Firefox, Safari, Edge)
- Internet connection (for loading external libraries and Google Sheets)
- No additional software installation required

## Browser Compatibility

- ✅ Chrome 60+
- ✅ Firefox 55+
- ✅ Safari 12+
- ✅ Edge 79+

## Files Structure

```
├── index.html          # Main HTML file
├── styles.css          # CSS styling
├── script.js           # JavaScript functionality
└── README.md           # Documentation
```

## Libraries Used

- **SheetJS (xlsx)**: For reading Excel files
- **docx**: For generating Word documents
- **Browser Print API**: For PDF generation

## Troubleshooting

### Common Issues:

1. **"Required columns not found" error**
   - Ensure your data has columns for Customer Name, Address, and Contact Number
   - Column names are case-insensitive and support variations

2. **"Error reading Excel file" error**
   - Verify the file is a valid Excel file (.xlsx or .xls)
   - Try saving the file again and re-uploading

3. **Google Sheets not loading**
   - Make sure the sheet is published to web as CSV
   - Verify the sheet is publicly accessible
   - Check the sharing link is correct

4. **PDF/Word not generating**
   - Check browser popup/download blockers
   - Ensure sufficient disk space
   - Verify customers are selected

5. **Format not applying correctly**
   - Try clicking "Reset to Defaults" and re-apply settings
   - Ensure all format inputs have valid values
   - Check the preview to verify formatting

## Performance Notes

- The application works entirely in your browser - no data is sent to external servers
- All processing happens locally for privacy and security
- Large datasets (1000+ customers) may take longer to process
- Generated documents are optimized for printing

## Recent Updates

### Version 2.0 Improvements:
- ✅ Removed all alert boxes for better user experience
- ✅ Enhanced PDF generation with proper multi-page support
- ✅ Added comprehensive format customization
- ✅ Improved error handling and user feedback
- ✅ Added Google Sheets integration
- ✅ Table-based customer selection
- ✅ Real-time format preview
- ✅ Settings persistence

## Contact Information

Created for: JD Sons Offline Order  
Contact: Jemish (JD Jewellery) - 9773046615 