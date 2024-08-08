# HTML Testing Extractor Tool

This project is a GUI-based application designed to extract test data from HTML files and save the results into an Excel file. The tool provides a user-friendly interface to select individual HTML files or entire folders containing HTML files, and it processes these files efficiently using Selenium and OpenPyXL.

## Features

- **Select Files and Folders**: Users can choose to select individual HTML files or entire folders containing HTML files for data extraction.
- **Progress Tracking**: A progress bar tracks the extraction process, providing visual feedback on the current progress.
- **Status Updates**: The application provides real-time status updates for each file processed, including any errors encountered.
- **Error Handling**: Robust error handling ensures that users are notified of any issues with file selection or data extraction.
- **Excel Output**: Extracted data is saved into a well-structured Excel file with formatted results, including overall test results.

## Technologies Used

- **PyQt5**: For building the graphical user interface.
- **Selenium**: For headless web scraping to extract data from HTML tables.
- **OpenPyXL**: For creating and formatting the Excel output file.
- **Python**: The core programming language used for developing the application.
- **Threading**: To maintain UI responsiveness during data processing.

