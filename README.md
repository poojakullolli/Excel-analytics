# Excel Analytics Platform

## Overview

Excel Analytics Platform is a modern, browser-based application that enables users to upload, analyze, and visualize Excel data through interactive charts and graphs. The platform provides a seamless experience for transforming spreadsheet data into meaningful visual insights with no coding required.

## Features

- **Excel File Upload**: Drag & drop or browse to upload .xls and .xlsx files (up to 10MB)
- **Data Preview**: View and analyze your Excel data before creating visualizations
- **Multiple Chart Types**: 
  - Bar Charts
  - Line Charts
  - Scatter Plots
  - Pie Charts
  - 3D Bar Charts (interactive with rotation)
- **Data Limiting**: Choose how many rows to analyze for large datasets
- **AI-Powered Insights**: Automatic generation of key statistics and observations
- **Export Options**: Download charts as PNG or PDF
- **Admin Dashboard**: Monitor usage statistics and manage uploaded files
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## Getting Started

1. **Upload an Excel File**
   - Drag and drop your Excel file into the upload area
   - Or click "browse files" to select a file from your device
   - Alternatively, use the sample data to try the platform

2. **Analyze Your Data**
   - View a preview of your data
   - Limit the number of rows to analyze (for large datasets)
   - Choose columns for X and Y axes
   - Select your preferred chart type

3. **Generate Visualizations**
   - Click "Generate Chart" to create your visualization
   - View automatically generated insights about your data
   - Interact with the chart (especially 3D charts which can be rotated)

4. **Export Your Results**
   - Download charts as PNG images or PDF files
   - Use the charts in presentations or reports

## Admin Features

Access the admin panel to:
- View platform usage statistics
- Manage uploaded files
- View detailed file information
- Delete files from storage

To access the admin panel:
1. Click the "Admin" button in the header
2. Enter the password (default: `Admin123`)

## Technical Details

- **Built With**: HTML5, CSS3, JavaScript (ES6+)
- **Libraries**:
  - Chart.js for 2D visualizations
  - Three.js for 3D charts
  - SheetJS for Excel file parsing
  - jsPDF for PDF generation
- **Storage**: Uses browser's localStorage for data persistence
- **No Server Required**: Runs entirely in the browser

## Browser Compatibility

- Chrome 70+
- Firefox 63+
- Safari 12+
- Edge 79+

## Future Enhancements

- Cloud storage integration
- Additional chart types and customization options
- Collaborative features for team analysis
- Advanced data transformation capabilities
- User account management

## License

MIT License - See LICENSE file for details

---

Developed with ❤️ by Excel Analytics Team
