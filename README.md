## Excel to PDF Conversion Script

This Python script automates the conversion of an Excel worksheet to a PDF file using the `win32com` module, which allows interaction with Microsoft Excel through the COM interface. It begins by importing the necessary library and launching Excel in the background without making it visible to the user. The script then specifies the path for the Excel file to be converted and the destination path where the resulting PDF will be saved. It opens the specified Excel file, selects the first worksheet, and exports it as a PDF to the designated location. After the conversion is complete, the script closes the Excel file and quits the Excel application to free up system resources. Finally, it prints a message confirming that the PDF has been successfully created and saved. This script streamlines the process of converting Excel files to PDFs, eliminating the need for manual intervention.
## Output
![exceltopdf](https://github.com/user-attachments/assets/ed1790c8-2b84-4d96-89a7-8999fc2cd11f)
