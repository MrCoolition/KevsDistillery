param(
  [Parameter(Mandatory = $true)]
  [string] $Path
)

$ErrorActionPreference = "Stop"

function New-Step {
  param([string] $Name)
  [ordered]@{
    name = $Name
    status = "not_run"
    error = $null
    rows = @()
  }
}

$result = [ordered]@{
  source_path = $Path
  generated_at = (Get-Date).ToString("o")
  file = $null
  excel_application = New-Step "excel_application"
  vba_components = New-Step "vba_components"
  shape_macro_bindings = New-Step "shape_macro_bindings"
  workbook_connections = New-Step "workbook_connections"
  query_tables = New-Step "query_tables"
  limitations = @()
}

$file = Get-Item -LiteralPath $Path
$result.file = [ordered]@{
  name = $file.Name
  full_name = $file.FullName
  length = $file.Length
  last_write_time = $file.LastWriteTime.ToString("o")
  creation_time = $file.CreationTime.ToString("o")
}

$excel = $null
$workbook = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.AskToUpdateLinks = $false
  $result.excel_application.status = "ok"

  $workbook = $excel.Workbooks.Open($Path, 0, $true)

  try {
    $componentRows = @()
    foreach ($component in $workbook.VBProject.VBComponents) {
      $lineCount = 0
      $code = ""
      try {
        $lineCount = $component.CodeModule.CountOfLines
        if ($lineCount -gt 0) {
          $code = $component.CodeModule.Lines(1, $lineCount)
        }
      } catch {
        $result.limitations += "VBA code module extraction failed for $($component.Name): $($_.Exception.Message)"
      }

      $componentRows += [pscustomobject][ordered]@{
        component_name = $component.Name
        component_type = $component.Type
        designer_id = $component.DesignerID
        line_count = $lineCount
        code = $code
      }
    }
    $result.vba_components.status = "ok"
    $result.vba_components.rows = @($componentRows)
  } catch {
    $result.vba_components.status = "blocked"
    $result.vba_components.error = $_.Exception.Message
    $result.limitations += "VBA project access blocked. In Excel Trust Center, enable trusted access to the VBA project object model for desktop deep export."
  }

  try {
    $shapeRows = @()
    foreach ($worksheet in $workbook.Worksheets) {
      foreach ($shape in $worksheet.Shapes) {
        $onAction = ""
        try {
          $onAction = $shape.OnAction
        } catch {}
        if ($onAction) {
          $shapeRows += [pscustomobject][ordered]@{
            sheet_name = $worksheet.Name
            shape_name = $shape.Name
            shape_type = $shape.Type
            on_action = $onAction
            alternative_text = $shape.AlternativeText
            title = $shape.Title
          }
        }
      }
    }
    $result.shape_macro_bindings.status = "ok"
    $result.shape_macro_bindings.rows = @($shapeRows)
  } catch {
    $result.shape_macro_bindings.status = "blocked"
    $result.shape_macro_bindings.error = $_.Exception.Message
  }

  try {
    $connectionRows = @()
    foreach ($connection in $workbook.Connections) {
      $oleDbConnection = $null
      $odbcConnection = $null
      $textConnection = $null
      try { $oleDbConnection = $connection.OLEDBConnection.Connection } catch {}
      try { $odbcConnection = $connection.ODBCConnection.Connection } catch {}
      try { $textConnection = $connection.TextConnection.Connection } catch {}
      $connectionRows += [pscustomobject][ordered]@{
        name = $connection.Name
        description = $connection.Description
        type = $connection.Type
        refresh_with_refresh_all = $connection.RefreshWithRefreshAll
        ole_db_connection = $oleDbConnection
        odbc_connection = $odbcConnection
        text_connection = $textConnection
      }
    }
    $result.workbook_connections.status = "ok"
    $result.workbook_connections.rows = @($connectionRows)
  } catch {
    $result.workbook_connections.status = "blocked"
    $result.workbook_connections.error = $_.Exception.Message
  }

  try {
    $queryRows = @()
    foreach ($worksheet in $workbook.Worksheets) {
      foreach ($listObject in $worksheet.ListObjects) {
        try {
          if ($listObject.QueryTable -ne $null) {
            $queryRows += [pscustomobject][ordered]@{
              sheet_name = $worksheet.Name
              list_object = $listObject.Name
              query_table_name = $listObject.QueryTable.Name
              command_text = $listObject.QueryTable.CommandText
              connection = $listObject.QueryTable.Connection
              refresh_style = $listObject.QueryTable.RefreshStyle
              refresh_on_file_open = $listObject.QueryTable.RefreshOnFileOpen
            }
          }
        } catch {}
      }
      foreach ($queryTable in $worksheet.QueryTables) {
        $queryRows += [pscustomobject][ordered]@{
          sheet_name = $worksheet.Name
          list_object = ""
          query_table_name = $queryTable.Name
          command_text = $queryTable.CommandText
          connection = $queryTable.Connection
          refresh_style = $queryTable.RefreshStyle
          refresh_on_file_open = $queryTable.RefreshOnFileOpen
        }
      }
    }
    $result.query_tables.status = "ok"
    $result.query_tables.rows = @($queryRows)
  } catch {
    $result.query_tables.status = "blocked"
    $result.query_tables.error = $_.Exception.Message
  }
} catch {
  $result.excel_application.status = "blocked"
  $result.excel_application.error = $_.Exception.Message
  $result.limitations += "Excel.Application collector failed: $($_.Exception.Message)"
} finally {
  if ($workbook -ne $null) {
    try { $workbook.Close($false) } catch {}
  }
  if ($excel -ne $null) {
    try { $excel.Quit() } catch {}
  }
}

$result | ConvertTo-Json -Depth 12
