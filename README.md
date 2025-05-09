# Volume_Fill_Parser
A simple parser to transfer PDF with long filling line data into excel files, and plot some graphs.

Volume Fill Parser is a Streamlit web application designed to extract measurement data from PDF files, perform statistical analysis, and generate comprehensive Excel reports with distribution graphs. This tool is particularly useful for analyzing measurement data from FlexLine M60 systems.

## Features
- Multiple File Processing: Upload and process multiple PDF files simultaneously
- Statistical Analysis: Calculate mean, standard deviation, min, and max values
- Customizable Cutoffs: Set lower and upper cutoff values for statistical analysis
- Data Visualization: Generate histograms with normal distribution curves
- Excel Output: Create detailed Excel reports with raw data and statistical graphs
- User-Friendly Interface: Simple, intuitive web interface with progress tracking

## Requirements
streamlit>=1.0.0
PyPDF2>=3.0.0
openpyxl>=3.0.0
numpy>=1.20.0
matplotlib>=3.4.0
scipy>=1.7.0 

## Data Processing Details

The application processes PDF files containing measurement data with the following workflow:

Extracts text from PDF files (starting from page 4)
Parses data lines containing position measurements
Filters values based on user-defined cutoffs
Calculates statistical metrics for each position
Generates distribution graphs with statistical indicators
Creates an Excel workbook with data sheets and graph sheets

## Output Format
The generated Excel file contains:

One data sheet per PDF file with raw measurement values

One graph sheet per PDF file with:
Distribution histograms with normal curves
Statistical indicators (min, max, mean)
Cutoff value indicators
Detailed statistics tables
Limitations
Designed specifically for PDF files with a particular format (FlexLine M60 reports)
Requires proper PDF text extraction capability
Excel sheet names limited to 31 characters

## Troubleshooting
If no data is extracted, check if the PDF format matches the expected structure
If graphs are not displaying correctly, verify that the data contains valid numerical values
For large files, processing may take longer; please be patient

## Version History
v1.0 (Current)

Initial release
Multiple file processing
Statistical analysis with customizable cutoffs
Distribution graph generation
Excel report creation

Developed by Nicholas Michelarakis
Last Updated: May 2025
