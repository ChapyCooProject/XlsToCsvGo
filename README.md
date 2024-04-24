Excel-CSV Conversion (Japanese Edition)
=======================================

Development Environment
-----------------------

   * go 1.21.0
   * Windows 10 64-bit

Introduction
------------

   * This software is designed to convert Excel files to CSV files.
   * The software operates as a standalone executable.
   * It only supports XLSX files and does not support XLS files.
   * It is compatible with password-protected Excel files.
   * It does not support multiple sheets.
   * It is compatible with both 64-bit and 32-bit versions of Windows.
   * No EXCEL or ACCESS components are required.

About the Executable File
-------------------------

   * There are both windowed and console applications.  
     Here are the details for each:

      #### *Windowed Application*
         ExcelToCsvGow.exe        (64-bit application)
         ExcelToCsvGow_x86.exe    (32-bit application)

      #### *Console Application*
         ExcelToCsvGoc.exe        (64-bit application)
         ExcelToCsvGoc_x86.exe    (32-bit application)

Usage Guide (Windowed Application)
----------------------------------

   1. Launch the executable file (ExcelToCsvGow.exe / ExcelToCsvGow_x86.exe).  
   2. A dialog will appear for selecting the Excel file.  
   3. Choose the delimiter.  
   4. If necessary, enter the decryption password.  
   5. Click "Select Excel File..." and choose the Excel file.  
   6. A CSV file with the same name will be created in the same location as the Excel file.  

Usage Guide (Console Application)
---------------------------------

   1. Launch the command prompt and navigate to the directory where the executable file is located.  
   2. Specify the following arguments and launch the executable file:  
         -d: Delimiter for CSV file (comma / tab)  
         -f: Excel file to convert  
         -p: Decryption password  
            ```  
            Example: ExcelToCsvGoc -d=comma -p=hitpass -f="C:\sample\sampleXLSX.xlsx"  
            ```  
         -h: Display help  
            ```  
            Example: ExcelToCsvGoc -h  
            ```  
   3. A CSV file with the same name will be created in the same location as the Excel file.  

Downloading the Executable File
-------------------------------

   * The executable file can be downloaded from the following URL.

      - [Download](https://drive.google.com/drive/folders/1tTVl_PZegv5GzEpaEAdBjBrc91978ibB?usp=sharing)

