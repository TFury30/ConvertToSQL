
---


 ¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦¦    ¦¦ ¦¦    ¦¦ ¦¦¦¦¦¦¦ ¦¦¦¦¦¦  ¦¦¦¦¦¦¦¦     ¦¦¦¦¦¦¦¦  ¦¦¦¦¦¦      ¦¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦       
¦¦      ¦¦    ¦¦ ¦¦¦¦   ¦¦ ¦¦    ¦¦ ¦¦      ¦¦   ¦¦    ¦¦           ¦¦    ¦¦    ¦¦     ¦¦      ¦¦    ¦¦ ¦¦       
¦¦      ¦¦    ¦¦ ¦¦ ¦¦  ¦¦ ¦¦    ¦¦ ¦¦¦¦¦   ¦¦¦¦¦¦     ¦¦           ¦¦    ¦¦    ¦¦     ¦¦¦¦¦¦¦ ¦¦    ¦¦ ¦¦       
¦¦      ¦¦    ¦¦ ¦¦  ¦¦ ¦¦  ¦¦  ¦¦  ¦¦      ¦¦   ¦¦    ¦¦           ¦¦    ¦¦    ¦¦          ¦¦ ¦¦ _  ¦¦ ¦¦       
 ¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦   ¦¦¦¦   ¦¦¦¦   ¦¦¦¦¦¦¦ ¦¦   ¦¦    ¦¦           ¦¦     ¦¦¦¦¦¦      ¦¦¦¦¦¦¦  ¦¦\\¦¦  ¦¦¦¦¦¦¦ 
                                                                                                   \\

   ConvertToSQL.ps1 - A PowerShell script for converting files into SQL INSERT and DELETE commands.
     
     How To Use:
       -file <Path to File>          Path to the file (CSV) to be converted to SQL.
       -folder <Folder>              Path to a directory that contains files to be converted.
       -table <TableName>            Name of the table into which the data is to be inserted.
                                       Defaultwert : $$TableName§§
     
     Example:
       .\ConvertToSQL.ps1 -file    'C:\Daten\Datei.csv' -table 'MeineTabelle'
       .\ConvertToSQL.ps1 -folder  'C:\Daten'
     
     Description:
       This script reads data from files (CSV) and converts them into SQL INSERT and DELETE commands.
       If a directory input is used, all supported files in the directory are processed.




# ConvertToSQL.ps1

## Overview

**ConvertToSQL.ps1** is a PowerShell script designed to convert various file formats, including CSV and XML, into SQL INSERT and DELETE statements. The script provides an easy way to manage database entries by allowing users to generate SQL commands from their data files, facilitating the import process into relational databases.

## Features

- Supports multiple file formats: CSV and XML.
- Generates SQL INSERT statements to add new records.
- Optionally generates DELETE statements to remove existing records before inserting new data.
- Processes individual files or all supported files in a specified directory.
- Logs activities and errors to a designated log file.

## Installation

Before running the script, ensure you have the following prerequisites:

1. **PowerShell**: The script is compatible with PowerShell 5.0 and above.


## Usage

You can run the script by specifying either a file or a directory containing files. If both are provided, the file takes precedence.

### Parameters

- `-file`: Path to the specific file (CSV, XML) to convert into SQL statements. This parameter is optional.
- `-folder`: Path to a directory containing files to convert. This parameter is optional.
- `-table`: Name of the database table to insert data into. If not provided, a placeholder will be used.

### Examples

1. Convert a specific CSV file into SQL statements:
   ```powershell
   .\ConvertToSQL.ps1 -file 'C:\Data\File.csv' -table 'MyTable'
   ```

2. Process all supported files in a directory:
   ```powershell
   .\ConvertToSQL.ps1 -folder 'C:\Data'
   ```


## How It Works

1. **File Loading**: The script reads data from the specified file format. For XML files, it removes the XML header to facilitate processing.
2. **SQL Generation**: It generates SQL INSERT statements for each record in the data file and DELETE statements based on a specified identifier.
3. **Output**: The SQL statements are written to a `.sql` file in the same directory as the input file or in the specified output directory.

## Logging

All activities and errors are logged in a file named `ConvertToSql.log` in the current working directory. This allows users to track the operations performed by the script.

## Notes

- The script prompts the user at the end if they want to generate additional DELETE statements.
- If no table name is provided, a placeholder will be used in the generated SQL statements.
- Ensure you have permission from the system administrator to install any necessary modules.

## Author

This script is maintained by Tobias Fourie.

## License

---
