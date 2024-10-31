# Version 7.1
<#
.SYNOPSIS
    A PowerShell script for converting files (CSV, XLS, XLSX, XML) into SQL INSERT and DELETE statements.

.DESCRIPTION
    This script reads data from various file formats (CSV, XLS, XLSX, XML) and converts this data into SQL INSERT statements.
    Additionally, it provides the option to generate DELETE statements to delete existing records before inserting new data.
    The script can process either a single file or all supported files in a specified directory.

.PARAMETER file
    Path to the file (CSV, XLS, XLSX, XML) to be converted into SQL INSERT and DELETE statements. This parameter is optional. If both `file` and `folder` are specified, `file` takes precedence.

.PARAMETER folder
    Path to a directory containing the files to be converted. The script processes all supported files in the directory recursively. This parameter is optional. If both `file` and `folder` are specified, `file` takes precedence.

.PARAMETER table
    Name of the table into which the data will be inserted. This parameter is optional. If no table name is provided, a placeholder will be used. The table name is used in the generated SQL statements.

.EXAMPLE
    .\ConvertToSQL.ps1 -file 'C:\Data\File.csv' -table 'MyTable'
    - Processes the file 'File.csv' and converts its contents into SQL INSERT and DELETE statements for the table 'MyTable'.

.EXAMPLE
    .\ConvertToSQL.ps1 -folder 'C:\Data'
    - Processes all supported files in the 'Data' directory and converts their contents into SQL INSERT and DELETE statements. The table name is derived from the file name.

.EXAMPLE
    .\ConvertToSQL.ps1 -file 'C:\Data\File.xlsx' -table 'MyTable'
    - Processes the Excel file 'File.xlsx' and converts its contents into SQL INSERT and DELETE statements for the table 'MyTable'.

.NOTES
    - To process Excel files (XLS and XLSX), the PowerShell module 'ImportExcel' is required. Make sure it is installed by running `Install-Module -Name ImportExcel`. Please obtain permission from the system administrator first.
    - For XML files, the XML header may be removed to facilitate processing.
    - The script will ask at the end whether additional DELETE statements should be generated to delete existing records before inserting new data.

#>

param (
    [parameter(Mandatory=$false, HelpMessage="Path to the file (CSV & XML) to be converted to SQL.")]
    [string]$file,

    [parameter(Mandatory=$false, HelpMessage="Path to a directory containing files to be converted.")]
    [string]$folder,

    [parameter(Mandatory=$false, HelpMessage="Name of the table that is in the file.")]
    [string]$table

)

# Global variable to control the DELETE query prompt
$global:deleteQueryAsked = $true

# Path to the log file
$global:logFile = ".\ConvertToSql.log"


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#~ Functions approximately in the order of Main, if possible and sensible ~#
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

#~~~~~~~~#
#~ Main ~#
#~~~~~~~~#

# ======================================================= 120 =========================================================
# Function to convert the file into SQL INSERT and DELETE statements
# ======================================================= 120 =========================================================
function Convert-ToSql {
    param (
        [string]$dataSource,
        [string]$tableName,
        [string]$outputDir
    )

    try {
        $fileExtension = [System.IO.Path]::GetExtension($dataSource).ToLower()

        # Load data from the file based on the file type
        switch ($fileExtension) {
            ".csv" {
                try {
                    $data = Import-Csv -Path $dataSource
                } catch {
                    Log-Message "Error reading the CSV file: $_" -type "ERROR"
                    Exit 1  # Clean exit of the script on error
                }
                break;
            }
            ".xml" {
                Process-XmlFile -xmlFilePath $dataSource
                try {
                    $xml = [xml](Get-Content -Path $dataSource)
                    $data = @()
                    foreach ($node in $xml.DocumentElement.ChildNodes) {
                        $item = @{}
                        foreach ($attribute in $node.Attributes) {
                            $item[$attribute.Name] = $attribute.Value
                        }
                        foreach ($childNode in $node.ChildNodes) {
                            $item[$childNode.Name] = $childNode.InnerText
                        }
                        $data += New-Object PSObject -Property $item
                    }
                } catch {
                    Log-Message "Error reading the XML file: $_" -type "ERROR"
                    Exit 1  # Clean exit of the script on error
                }
                break;
            }
            default {
                Log-Message "Invalid file type. Supported formats are CSV & XML." -type "ERROR"
                Exit 1  # Clean exit of the script on invalid file type
            }
        }

        if ($data.Count -eq 0) {
            Log-Message "The file is empty or could not be loaded." -type "ERROR"
            Exit 1  # Clean exit of the script on empty file
        }
        if ($tableName -eq "") {
            Log-Message "No table name was provided. The table name will be filled with a placeholder." -type "WARNING"
            $tableName = "§§TableName§§"
        }

        $sqlStatements = @()
        $deleteStatements = @()

        foreach ($row in $data) {
            $columns = $data[0].PSObject.Properties.Name -join ", "
            $values = @()

            foreach ($column in $row.PSObject.Properties.Name) {
                $value = $row.$column
                if ($null -eq $value -or $value -eq '') {
                    $values += "NULL"
                } else {
                    $values += if ($value -match "^\d+$") { $value } else { "'$value'" }
                }
            }

            $valuesString = $values -join ", "
            $sql = "INSERT INTO $tableName ($columns) VALUES ($valuesString);"
            $sqlStatements += $sql

            # Create the DELETE statement based on the ID (first column)
            $idValue = $row.$($columns.Split(", ")[0])
            $deleteStatement = "DELETE FROM $tableName WHERE $($columns.Split(", ")[0]) = $idValue;"
            $deleteStatements += $deleteStatement
        }

        $outputFilePath = Join-Path $outputDir "$tableName.sql"

        # Write the DELETE statements followed by the INSERT statements to the output file
        Set-Content -Path $outputFilePath -Value ($deleteStatements -join "`n")
        Add-Content -Path $outputFilePath -Value "`n"  # Leerzeile zwischen DELETE und INSERT
        Add-Content -Path $outputFilePath -Value ($sqlStatements -join "`n")

        Log-Message "The SQL statements were written to the file '$outputFilePath'."
    } catch {
        Log-Message "Error converting to SQL: $_" -type "ERROR"
        Exit 1  # Clean exit of the script on error
    }
}


# ======================================================= 120 =========================================================
# Function to display help
# ======================================================= 120 =========================================================
function Show-Help {
    Log-Message "Displaying help." -type "INFO"

    Write-Host "                                                                                                                 "
    Write-Host "                                                                                                                 "
    Write-Host " ¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦¦    ¦¦ ¦¦    ¦¦ ¦¦¦¦¦¦¦ ¦¦¦¦¦¦  ¦¦¦¦¦¦¦¦     ¦¦¦¦¦¦¦¦  ¦¦¦¦¦¦      ¦¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦       "
    Write-Host "¦¦      ¦¦    ¦¦ ¦¦¦¦   ¦¦ ¦¦    ¦¦ ¦¦      ¦¦   ¦¦    ¦¦           ¦¦    ¦¦    ¦¦     ¦¦      ¦¦    ¦¦ ¦¦       "
    Write-Host "¦¦      ¦¦    ¦¦ ¦¦ ¦¦  ¦¦ ¦¦    ¦¦ ¦¦¦¦¦   ¦¦¦¦¦¦     ¦¦           ¦¦    ¦¦    ¦¦     ¦¦¦¦¦¦¦ ¦¦    ¦¦ ¦¦       "
    Write-Host "¦¦      ¦¦    ¦¦ ¦¦  ¦¦ ¦¦  ¦¦  ¦¦  ¦¦      ¦¦   ¦¦    ¦¦           ¦¦    ¦¦    ¦¦          ¦¦ ¦¦ _  ¦¦ ¦¦       "
    Write-Host " ¦¦¦¦¦¦  ¦¦¦¦¦¦  ¦¦   ¦¦¦¦   ¦¦¦¦   ¦¦¦¦¦¦¦ ¦¦   ¦¦    ¦¦           ¦¦     ¦¦¦¦¦¦      ¦¦¦¦¦¦¦  ¦¦\\¦¦  ¦¦¦¦¦¦¦ "
    Write-Host "                                                                                                   \\           "
    Write-Host "                                                                                                                 "
    Write-Host "                                                                                                                 "
    Write-Host "                                                                                                                 "
    Write-Host "                                                                                                                 "


    Write-Host "ConvertToSQL.ps1 - A PowerShell script for converting files into SQL INSERT and DELETE commands."
    Write-Host ""
    Write-Host "How To Use:"
    Write-Host "  -file <Path to File>          Path to the file (CSV) to be converted to SQL."
    Write-Host "  -folder <Folder>              Path to a directory that contains files to be converted."
    Write-Host "  -table <TableName>            Name of the table into which the data is to be inserted."
    Write-Host "                                  Defaultwert : $$TableName§§"
    Write-Host ""
    Write-Host "Example:"
    Write-Host "  .\ConvertToSQL.ps1 -file    'C:\Daten\Datei.csv' -table 'MeineTabelle'"
    Write-Host "  .\ConvertToSQL.ps1 -folder  'C:\Daten'"
    Write-Host ""
    Write-Host "Description:"
    Write-Host "  This script reads data from files (CSV) and converts them into SQL INSERT and DELETE commands."
    Write-Host "  If a directory input is used, all supported files in the directory are processed."
}


# ======================================================= 120 =========================================================
# Function to log messages
# ======================================================= 120 =========================================================
function Process-XmlFile {
    param (
        [string]$xmlFilePath
    )

    try {
        $xmlContent = Get-Content -Path $xmlFilePath -Raw
        Log-Message "XML-Content of '$xmlFilePath' readed." -type "INFO"
        Set-Content -Path $xmlFilePath -Value $xmlContent

        $xmlLines = Get-Content -Path $xmlFilePath
        if ($xmlLines[0] -match "^<\?xml version='1\.1' encoding='UTF-8'\?>$") {
            $xmlLines = $xmlLines | Select-Object -Skip 1
            Set-Content -Path $xmlFilePath -Value $xmlLines
            Log-Message "XML header removed from '$xmlFilePath'." -type "INFO"
        }
    } catch {
        Log-Message "Error when processing the XML file: $_" -type "ERROR"
        Exit 1  # Clean termination of the script in the event of an error
    }
}


# ======================================================= 120 =========================================================
# Function for processing files in a directory
# ======================================================= 120 =========================================================
function Process-Directory {
    param (
        [string]$directoryPath
    )

    try {
        $files = Get-ChildItem -Path $directoryPath -File -Recurse | Where-Object { $_.Extension -in ".csv", ".xml" }

        foreach ($file in $files) {
            Convert-ToSql -dataSource $file.FullName -tableName $file.BaseName -outputDir $file.DirectoryName
        }
    } catch {
        Log-Message "Error when processing the directory: $_" -type "ERROR"
        Exit 1  # Clean termination of the script in the event of an error
    }
}


#~~~~~~~~~~~~~~~~~~~~#
#~ Helper functions ~#
#~~~~~~~~~~~~~~~~~~~~#
# Interception of user interruptions (z.B. Ctrl+C)
trap {
    Log-Message "The script was interrupted by the user." -type "INFO"
    Exit 1  # Clean termination of the script when the user is interrupted
}

#~~~~~~~~~~~~~~~~~~~~#
#~ Log function ~#
#~~~~~~~~~~~~~~~~~~~~#
# Function for logging messages
function Log-Message {
    param (
        [string]$message,
        [string]$type = "INFO"
    )

    # Format for the log message
    $formattedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$type] $message"

    # Output on the console
    Write-Output $formattedMessage

    # Writing to the log file
    Add-Content -Path $global:logFile -Value $formattedMessage
}

# Function for processing the file in SQL INSERT and DELETE commands

#~~~~~~~~~~~~~~~~#
#~ ScriptCall ~#
#~~~~~~~~~~~~~~~~#
if ($file) {
    Convert-ToSql -dataSource $file -tableName $table -outputDir (Split-Path -Path $file -Parent)
} elseif ($folder) {
    Process-Directory -directoryPath $folder
} else {
    Show-Help
    Exit 0  # Clean termination of the script after help
}
