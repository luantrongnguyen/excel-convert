# Excel Data Filler Web App

A React-based web application for mapping and filling data from Excel files into templates with drag-and-drop functionality.

## Features

- **File Upload**: Upload template and data Excel files
- **Drag & Drop Mapping**: Drag columns from data file to template columns
- **DOB Special Handling**: Special mapping for Date of Birth with month, day, and year
- **Default Values**: Set default values for unmapped columns
- **Export**: Generate Excel files with a limit of 1000 records per file
- **Responsive Design**: Works on desktop and mobile devices

## Installation

1. Navigate to the project directory:
```bash
cd convert-excel
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```

4. Open [http://localhost:3000](http://localhost:3000) to view it in the browser.

## Usage

### 1. Upload Files
- Upload your template Excel file (the structure you want to fill)
- Upload your data Excel file (the source data)

### 2. Map Columns
- Drag columns from the "Data Columns" section to the corresponding "Template Columns"
- For Date of Birth (DOB), you can:
  - Drag one column that contains complete date (MM/DD/YYYY)
  - Or drag three separate columns for month, day, and year
  - If you only drag month and day, the year will default to the current year

### 3. Set Default Values
- For columns that don't have data, set default values in the input fields
- These values will be used for all records where no data is available

### 4. Export
- Click "Export Excel Files" to generate the output
- Files will be limited to 1000 records each
- Multiple files will be created if you have more than 1000 records

## Special Features

### Date of Birth (DOB) Mapping
The app has special handling for Date of Birth columns:

1. **Complete Date**: If you drag a column that contains complete dates (MM/DD/YYYY), it will use that directly
2. **Separate Components**: You can drag three separate columns:
   - Month column
   - Day column  
   - Year column
3. **Partial Date**: If you only provide month and day, the year will default to the current year

### Default Values
- Set default values for any column that doesn't have data
- Useful for filling in missing information like addresses, phone numbers, etc.

## Technical Details

### Dependencies
- React 19.2.0
- TypeScript 4.9.5
- react-beautiful-dnd for drag and drop
- react-dropzone for file uploads
- xlsx for Excel file processing
- file-saver for file downloads

### Browser Support
- Chrome (recommended)
- Firefox
- Safari
- Edge

## File Formats Supported
- .xlsx (Excel 2007+)
- .xls (Excel 97-2003)

## Limitations
- Maximum 1000 records per exported file
- Files are processed in the browser (limited by available memory)
- Large files may take time to process

## Troubleshooting

### Common Issues
1. **File won't upload**: Make sure the file is a valid Excel format (.xlsx or .xls)
2. **Drag and drop not working**: Try refreshing the page and ensure JavaScript is enabled
3. **Export not working**: Check that all required columns are mapped

### Performance Tips
- For large files, consider splitting your data into smaller chunks
- Close other browser tabs to free up memory
- Use Chrome for best performance

## Development

### Project Structure
```
src/
├── App.tsx          # Main application component
├── App.css          # Application styles
├── index.tsx        # Application entry point
└── index.css        # Global styles
```

### Building for Production
```bash
npm run build
```

This builds the app for production to the `build` folder.

## License

This project is licensed under the MIT License.