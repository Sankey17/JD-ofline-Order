# JD Sons Offline Order - Customer Word Document Generator

This website allows you to upload an Excel file containing customer data and generate a formatted Word document with customer details.

## Features

- **Excel File Upload**: Support for .xlsx and .xls files
- **Drag & Drop**: Easy file upload with drag and drop functionality
- **Customer Selection**: Choose specific range of customers (20-25 or any custom range)
- **Word Document Generation**: Creates properly formatted Word documents
- **One Customer Per Page**: Each customer gets their own page in the Word document
- **Professional Formatting**: Clean, printable format with proper spacing

## How to Use

1. **Open the Website**: Open `index.html` in your web browser
2. **Upload Excel File**: 
   - Click the upload area or drag and drop your Excel file
   - The file must contain columns for Customer Name, Address, and Contact Number
3. **Select Customer Range**: 
   - Choose the starting and ending customer numbers
   - Default is set to generate 1-25 customers
4. **Generate Document**: Click "Generate Word Document" button
5. **Download**: The Word document will automatically download

## Excel File Requirements

Your Excel file must contain these columns (case-insensitive):
- **Customer Name** (or "Name", "Customer")
- **Address** (or "Addr")
- **Contact Number** (or "Contact", "Phone", "Mobile")

### Example Excel Structure:
```
| Order Date | Customer Name | Address        | Contact Number | City | ... |
|------------|---------------|----------------|----------------|------|-----|
| 2024-01-01 | John Doe      | 123 Main St    | 9876543210     | NYC  | ... |
| 2024-01-02 | Jane Smith    | 456 Oak Ave    | 9876543211     | LA   | ... |
```

## Generated Word Document Format

Each page in the Word document contains:

```
From: Jemish (JD Jewellery)
Contact Number: 9773046615

─────────────────────────────────────────────────

Customer Details:

Customer Name: [Customer Name]
Address: [Customer Address]
Contact Number: [Customer Contact Number]

─────────────────────────────────────────────────
                JD Sons Offline Order
```

## Technical Requirements

- Modern web browser (Chrome, Firefox, Safari, Edge)
- Internet connection (for loading external libraries)
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
└── README.md           # This file
```

## Libraries Used

- **SheetJS (xlsx)**: For reading Excel files
- **docx**: For generating Word documents

## Troubleshooting

### Common Issues:

1. **"Required columns not found" error**
   - Make sure your Excel file has columns named "Customer Name", "Address", and "Contact Number"
   - Column names are case-insensitive but must contain these words

2. **"Error reading Excel file" error**
   - Ensure the file is a valid Excel file (.xlsx or .xls)
   - Try saving the Excel file again and re-uploading

3. **Document not downloading**
   - Check if your browser is blocking downloads
   - Ensure you have enough disk space
   - Try using a different browser

4. **Website not working**
   - Make sure you have an internet connection
   - Try refreshing the page
   - Check browser console for errors (F12 → Console)

## Support

If you encounter any issues:
1. Check that your Excel file format matches the requirements
2. Try with a smaller number of customers first
3. Make sure all required columns are present in your Excel file

## Notes

- The website works entirely in your browser - no data is sent to any server
- All processing happens locally on your computer
- Generated documents are automatically downloaded to your default download folder
- Each customer appears on a separate page in the Word document
- The format is optimized for printing

## Contact Information

Created for: JD Sons Offline Order
Contact: Jemish (JD Jewellery) - 9773046615 