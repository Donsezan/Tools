param (
    [Switch]$CreateNewFile,
    [String]$ExcelFilePath = ".\SQL_Report.xlsx",
    [String]$SqlFolderPath = ".\SQL"
)

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# Connection string defaults to Master. 
# Scripts in the SQL folder are expected to contain 'USE [DatabaseName]' if a specific DB is required.
$ConnectionString = "Server=localhost;Database=Master;Integrated Security=True;"
# ---------------------------------------------------------------------------

Function Get-SqlDataTable {
    param(
        [string]$Query,
        [string]$ConnString
    )
    
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection $ConnString
        $conn.Open()
        
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        # Increase command timeout for longer queries (default is 30s)
        $cmd.CommandTimeout = 300 
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt = New-Object System.Data.DataTable
        [void]$adapter.Fill($dt)
        
        return $dt
    }
    catch {
        Write-Error "Error executing SQL: $_"
        return $null
    }
    finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
    }
}

# Ensure absolute paths
$SqlFolderPath = Resolve-Path $SqlFolderPath -ErrorAction SilentlyContinue
if (-not $SqlFolderPath) {
    Write-Error "SQL Folder not found: $SqlFolderPath"
    exit
}

$Timestamp = Get-Date -Format "yyyy-MM-dd HH-mm"

# ---------------------------------------------------------------------------
# EXCEL SETUP
# ---------------------------------------------------------------------------
try {
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Could not create Excel Application object. Please ensure Microsoft Excel is installed."
    }
    $excel.Visible = $false # Run in background
    $excel.DisplayAlerts = $false

    if ($CreateNewFile -or (-not (Test-Path $ExcelFilePath))) {
        # Create New File
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $Timestamp
        
        # If creating new, we'll need to determine save path later if not provided or just use ExcelFilePath
        if (-not $ExcelFilePath -or (-not (Test-Path $ExcelFilePath) -and $CreateNewFile)) {
            $ExcelFilePath = ".\SQL_Report_$Timestamp.xlsx"
        }
        $SaveMode = "SaveAs"
    }
    else {
        # Open Existing File
        $ExcelFilePath = Resolve-Path $ExcelFilePath
        $workbook = $excel.Workbooks.Open($ExcelFilePath)
        
        # Add a new worksheet at the end
        $sheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count))
        
        # Ensure unique sheet name (Excel limits sheet names to 31 chars and no duplicates)
        try {
            $sheet.Name = $Timestamp
        }
        catch {
            $sheet.Name = $Timestamp + "_" + (Get-Random -Minimum 100 -Maximum 999)
        }
        $SaveMode = "Save"
    }

    $currentRow = 1

    # ---------------------------------------------------------------------------
    # PROCESS SQL FILES
    # ---------------------------------------------------------------------------
    $sqlFiles = Get-ChildItem -Path $SqlFolderPath -Filter "*.sql" | Sort-Object Name

    foreach ($file in $sqlFiles) {
        Write-Host "Processing $($file.Name)..." -ForegroundColor Cyan
        
        $sqlContent = Get-Content $file.FullName -Raw
        
        # We assume the whole file is one batch (or the user handles GO inside the script properly if they use Invoke-Sqlcmd, 
        # but System.Data.SqlClient usually fails on 'GO'. 
        # Simple fix: Remove 'GO' lines if they exist on their own line, or assume script is pure T-SQL.)
        # For this implementation request, we execute as is provided by Get-Content.
        
        $dataTable = Get-SqlDataTable -Query $sqlContent -ConnString $ConnectionString
        
        if ($dataTable) {
            # 1. Write Script Name
            $sheet.Cells.Item($currentRow, 1) = "Script: $($file.Name)"
            $headerRange = $sheet.Range($sheet.Cells.Item($currentRow, 1), $sheet.Cells.Item($currentRow, 1))
            $headerRange.Font.Bold = $true
            $headerRange.Interior.ColorIndex = 37 # Light Blue style
            
            $currentRow++
            
            # 2. Write Column Headers
            $colIndex = 1
            foreach ($col in $dataTable.Table.Columns) {
                $sheet.Cells.Item($currentRow, $colIndex) = $col.ColumnName
                $colIndex++
            }
            
            $headerRowStart = $currentRow
            $currentRow++
            
            # 3. Write Data
            # Using 2D array for speed
            if ($dataTable.Table.Rows.Count -gt 0) {
                $dataArray = New-Object 'object[,]' $dataTable.Table.Rows.Count, $dataTable.Table.Columns.Count
                for ($r = 0; $r -lt $dataTable.Table.Rows.Count; $r++) {
                    for ($c = 0; $c -lt $dataTable.Table.Columns.Count; $c++) {
                        $dataArray[$r, $c] = $dataTable.Table.Rows[$r][$c].ToString()
                    }
                }
                
                $startCell = $sheet.Cells.Item($currentRow, 1)
                $endCell = $sheet.Cells.Item($currentRow + $dataTable.Table.Rows.Count - 1, $dataTable.Table.Columns.Count)
                $range = $sheet.Range($startCell, $endCell)
                $range.Value2 = $dataArray
                
                $currentRow += $dataTable.Table.Rows.Count
            }

            # 4. Format as Table
            if ($dataTable.Table.Columns.Count -gt 0) {
                $lastRow = $currentRow - 1
                $tableRange = $sheet.Range($sheet.Cells.Item($headerRowStart, 1), $sheet.Cells.Item($lastRow, $dataTable.Table.Columns.Count))
                
                # xlSrcRange = 1, xlYes = 1
                $listObject = $sheet.ListObjects.Add(1, $tableRange, $null, 1)
                
                # Sanitize table name to be Excel compliant (No spaces, start with letter)
                $safeName = "Tbl_" + ($file.BaseName -replace '[^a-zA-Z0-9_]', '') + "_" + (Get-Random -Minimum 1000 -Maximum 9999)
                $listObject.Name = $safeName
                $listObject.TableStyle = "TableStyleMedium2"
            }
            
            # 4. Empty Row Separator
            $currentRow++
        }
        else {
            Write-Host "No data returned or error for $($file.Name)" -ForegroundColor Yellow
        }
    }

    # ---------------------------------------------------------------------------
    # FINAL FORMATTING & SAVE
    # ---------------------------------------------------------------------------
    Write-Host "Applying formatting..." -ForegroundColor Green
    $sheet.UsedRange.Columns.AutoFit()

    if ($SaveMode -eq "SaveAs") {
        # Check if file exists, if do, simple save, else SaveAs
        # Excel SaveAs requires absolute path
        $AbsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExcelFilePath)
        # Delete if exists to overwrite (since we are in CreateNewFile mode)
        if (Test-Path $AbsPath) { Remove-Item $AbsPath -Force }
        $workbook.SaveAs($AbsPath)
        Write-Host "File created at: $AbsPath" -ForegroundColor Green
    }
    else {
        $workbook.Save()
        Write-Host "File updated: $ExcelFilePath" -ForegroundColor Green
    }

}
catch {
    Write-Error "An unexpected error occurred: $_"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
