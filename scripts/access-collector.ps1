param(
  [Parameter(Mandatory = $true)]
  [string] $Path,

  [int64] $LargeFileThresholdBytes = 157286400,

  [ValidateSet("all", "fast", "deep")]
  [string] $Mode = "all",

  [int] $MaxDeepObjects = 300
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

function Convert-DataTable {
  param($Table)
  $rows = @()
  foreach ($row in $Table.Rows) {
    $item = [ordered]@{}
    foreach ($column in $Table.Columns) {
      $value = $row[$column.ColumnName]
      if ($value -is [System.DBNull]) {
        $value = $null
      }
      $item[$column.ColumnName] = $value
    }
    $rows += [pscustomobject]$item
  }
  return $rows
}

function Invoke-Query {
  param(
    [System.Data.OleDb.OleDbConnection] $Connection,
    [string] $Sql
  )
  $command = $Connection.CreateCommand()
  $command.CommandText = $Sql
  $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
  $table = New-Object System.Data.DataTable
  [void]$adapter.Fill($table)
  return Convert-DataTable $table
}

$result = [ordered]@{
  source_path = $Path
  generated_at = (Get-Date).ToString("o")
  file = $null
  provider = $null
  ole_db_schema_tables = New-Step "ole_db_schema_tables"
  ole_db_schema_columns = New-Step "ole_db_schema_columns"
  msys_objects = New-Step "msys_objects"
  msys_relationships = New-Step "msys_relationships"
  dao_querydefs = New-Step "dao_querydefs"
  dao_tabledefs = New-Step "dao_tabledefs"
  dao_fields = New-Step "dao_fields"
  dao_indexes = New-Step "dao_indexes"
  dao_relations = New-Step "dao_relations"
  dao_documents = New-Step "dao_documents"
  access_application = New-Step "access_application"
  limitations = @()
}

$file = Get-Item -LiteralPath $Path
$isLargeFile = $file.Length -gt $LargeFileThresholdBytes
$result.file = [ordered]@{
  name = $file.Name
  full_name = $file.FullName
  length = $file.Length
  last_write_time = $file.LastWriteTime.ToString("o")
  creation_time = $file.CreationTime.ToString("o")
  large_file_threshold_bytes = $LargeFileThresholdBytes
  large_file_mode = $isLargeFile
}

if ($isLargeFile) {
  $result.limitations += "Large file mode enabled. Expensive table row counts are skipped unless ACCESS_COLLECTOR_ROW_COUNTS=1. Trusted desktop deep Access.Application SaveAsText export still runs when ACCESS_COLLECTOR_DEEP=1."
}

Add-Type -AssemblyName System.Data

if ($Mode -ne "deep") {
$providers = @("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0")
$connection = $null

foreach ($provider in $providers) {
  try {
    $connectionString = "Provider=$provider;Data Source=$Path;Mode=Read;Persist Security Info=False;"
    $connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
    $connection.Open()
    $result.provider = $provider
    break
  } catch {
    $connection = $null
    $result.limitations += "Provider $provider failed: $($_.Exception.Message)"
  }
}

if ($connection -ne $null) {
  try {
    $tables = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $result.ole_db_schema_tables.status = "ok"
    $result.ole_db_schema_tables.rows = @(Convert-DataTable $tables)
  } catch {
    $result.ole_db_schema_tables.status = "blocked"
    $result.ole_db_schema_tables.error = $_.Exception.Message
  }

  try {
    $columns = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns, $null)
    $result.ole_db_schema_columns.status = "ok"
    $result.ole_db_schema_columns.rows = @(Convert-DataTable $columns)
  } catch {
    $result.ole_db_schema_columns.status = "blocked"
    $result.ole_db_schema_columns.error = $_.Exception.Message
  }

  try {
    $result.msys_objects.rows = @(Invoke-Query $connection "SELECT Id, Name, Type, Flags, DateCreate, DateUpdate, Database, ForeignName, Connect FROM MSysObjects")
    $result.msys_objects.status = "ok"
  } catch {
    $result.msys_objects.status = "blocked"
    $result.msys_objects.error = $_.Exception.Message
  }

  try {
    $result.msys_relationships.rows = @(Invoke-Query $connection "SELECT * FROM MSysRelationships")
    $result.msys_relationships.status = "ok"
  } catch {
    $result.msys_relationships.status = "blocked"
    $result.msys_relationships.error = $_.Exception.Message
  }

  $connection.Close()
}

$daoCreated = $false
foreach ($daoProgId in @("DAO.DBEngine.120", "DAO.DBEngine.36")) {
  if ($daoCreated) {
    break
  }
  try {
    $engine = New-Object -ComObject $daoProgId
    $db = $engine.OpenDatabase($Path, $false, $true)
    $queryRows = @()
    foreach ($queryDef in $db.QueryDefs) {
      $queryRows += [pscustomobject][ordered]@{
        name = $queryDef.Name
        sql = $queryDef.SQL
        type = $queryDef.Type
        connect = $queryDef.Connect
        returns_records = $queryDef.ReturnsRecords
      }
    }
    $tableRows = @()
    $fieldRows = @()
    $indexRows = @()
    foreach ($tableDef in $db.TableDefs) {
      $fieldCount = $null
      try {
        $fieldCount = $tableDef.Fields.Count
      } catch {
        $fieldCount = "blocked: $($_.Exception.Message)"
      }
      $tableRows += [pscustomobject][ordered]@{
        name = $tableDef.Name
        attributes = $tableDef.Attributes
        connect = $tableDef.Connect
        source_table_name = $tableDef.SourceTableName
        field_count = $fieldCount
        record_count = $null
      }
      $ordinal = 0
      try {
        foreach ($field in $tableDef.Fields) {
          $ordinal += 1
          $fieldRows += [pscustomobject][ordered]@{
            table_name = $tableDef.Name
            column_name = $field.Name
            ordinal_position = $ordinal
            data_type = $field.Type
            size = $field.Size
            required = $field.Required
            allow_zero_length = $field.AllowZeroLength
            default_value = $field.DefaultValue
            validation_rule = $field.ValidationRule
            validation_text = $field.ValidationText
            extraction_status = "ok"
            extraction_error = $null
          }
        }
      } catch {
        $fieldRows += [pscustomobject][ordered]@{
          table_name = $tableDef.Name
          column_name = $null
          ordinal_position = $null
          data_type = $null
          size = $null
          required = $null
          allow_zero_length = $null
          default_value = $null
          validation_rule = $null
          validation_text = $null
          extraction_status = "blocked"
          extraction_error = $_.Exception.Message
        }
      }
      try {
        foreach ($index in $tableDef.Indexes) {
          $indexFields = @()
          foreach ($field in $index.Fields) {
            $indexFields += $field.Name
          }
          $indexRows += [pscustomobject][ordered]@{
            table_name = $tableDef.Name
            index_name = $index.Name
            fields = ($indexFields -join "; ")
            primary = $index.Primary
            unique = $index.Unique
            required = $index.Required
            ignore_nulls = $index.IgnoreNulls
            extraction_status = "ok"
            extraction_error = $null
          }
        }
      } catch {
        $indexRows += [pscustomobject][ordered]@{
          table_name = $tableDef.Name
          index_name = $null
          fields = $null
          primary = $null
          unique = $null
          required = $null
          ignore_nulls = $null
          extraction_status = "blocked"
          extraction_error = $_.Exception.Message
        }
      }
      $rowCountMode = if ($isLargeFile -and $env:ACCESS_COLLECTOR_ROW_COUNTS -ne "1") { "skipped_large_file" } else { "attempted" }
      try {
        if ($rowCountMode -eq "skipped_large_file") {
          throw "row count skipped for large file mode"
        }
        $recordset = $db.OpenRecordset("SELECT Count(*) AS RowCount FROM [$($tableDef.Name.Replace(']', ']]'))]")
        if (-not $recordset.EOF) {
          $tableRows[-1].record_count = $recordset.Fields.Item("RowCount").Value
        }
        $recordset.Close()
      } catch {
        $tableRows[-1].record_count = $rowCountMode
      }
    }
    $relationRows = @()
    foreach ($relation in $db.Relations) {
      $relationFields = @()
      foreach ($field in $relation.Fields) {
        $relationFields += "$($field.Name)->$($field.ForeignName)"
      }
      $relationRows += [pscustomobject][ordered]@{
        name = $relation.Name
        table = $relation.Table
        foreign_table = $relation.ForeignTable
        attributes = $relation.Attributes
        fields = ($relationFields -join "; ")
      }
    }
    $documentRows = @()
    foreach ($container in $db.Containers) {
      foreach ($document in $container.Documents) {
        $documentRows += [pscustomobject][ordered]@{
          container = $container.Name
          name = $document.Name
          owner = $document.Owner
          date_created = $document.DateCreated
          last_updated = $document.LastUpdated
        }
      }
    }
    $db.Close()
    $result.dao_querydefs.status = "ok"
    $result.dao_querydefs.rows = @($queryRows)
    $result.dao_tabledefs.status = "ok"
    $result.dao_tabledefs.rows = @($tableRows)
    $result.dao_fields.status = "ok"
    $result.dao_fields.rows = @($fieldRows)
    $result.dao_indexes.status = "ok"
    $result.dao_indexes.rows = @($indexRows)
    $result.dao_relations.status = "ok"
    $result.dao_relations.rows = @($relationRows)
    $result.dao_documents.status = "ok"
    $result.dao_documents.rows = @($documentRows)
    $daoCreated = $true
  } catch {
    $result.limitations += "DAO collector $daoProgId failed: $($_.Exception.Message)"
    $result.dao_querydefs.status = "blocked"
    $result.dao_querydefs.error = $_.Exception.Message
    $result.dao_tabledefs.status = "blocked"
    $result.dao_tabledefs.error = $_.Exception.Message
    $result.dao_fields.status = "blocked"
    $result.dao_fields.error = $_.Exception.Message
    $result.dao_indexes.status = "blocked"
    $result.dao_indexes.error = $_.Exception.Message
    $result.dao_relations.status = "blocked"
    $result.dao_relations.error = $_.Exception.Message
    $result.dao_documents.status = "blocked"
    $result.dao_documents.error = $_.Exception.Message
  }
}
}

if ($Mode -eq "fast") {
  $result.access_application.status = "skipped"
  $result.access_application.error = "Access.Application SaveAsText export skipped in fast collector phase. Deep export runs in a bounded follow-up phase."
} elseif ($env:ACCESS_COLLECTOR_DEEP -ne "1") {
  $result.access_application.status = "skipped"
  $result.access_application.error = "Access.Application SaveAsText export is disabled by default to keep local and Vercel-compatible discovery bounded. Set ACCESS_COLLECTOR_DEEP=1 for a trusted desktop deep export."
  $result.limitations += $result.access_application.error
} else {
try {
  $access = New-Object -ComObject Access.Application
  $access.Visible = $false
  $access.OpenCurrentDatabase($Path, $false)
  $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("dossier_access_" + [System.Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
  $exports = @()

  $collections = @(
    @{ label = "query"; collection = $access.CurrentData.AllQueries; object_type = 1 },
    @{ label = "macro"; collection = $access.CurrentProject.AllMacros; object_type = 4 },
    @{ label = "form"; collection = $access.CurrentProject.AllForms; object_type = 2 },
    @{ label = "report"; collection = $access.CurrentProject.AllReports; object_type = 3 },
    @{ label = "module"; collection = $access.CurrentProject.AllModules; object_type = 5 }
  )

  foreach ($entry in $collections) {
    $objectCounter = 0
    foreach ($object in $entry.collection) {
      $objectCounter += 1
      if ($objectCounter -gt $MaxDeepObjects) {
        $result.limitations += "Access.Application SaveAsText $($entry.label) export stopped after $MaxDeepObjects objects to keep the trusted desktop deep export bounded."
        break
      }
      $name = $object.Name
      $exportPath = Join-Path $tempRoot (($entry.label + "_" + ($name -replace '[\\/:*?"<>|]', '_')) + ".txt")
      $status = "not_exported"
      $text = ""
      $errorText = $null
      try {
        $access.SaveAsText($entry.object_type, $name, $exportPath)
        $text = Get-Content -LiteralPath $exportPath -Raw -ErrorAction Stop
        $status = "ok"
      } catch {
        $status = "blocked"
        $errorText = $_.Exception.Message
      }
      $exports += [pscustomobject][ordered]@{
        object_type = $entry.label
        name = $name
        save_as_text_status = $status
        save_as_text_error = $errorText
        text = $text
      }
    }
  }

  try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
  $access.CloseCurrentDatabase()
  $access.Quit()
  $result.access_application.status = "ok"
  $result.access_application.rows = @($exports)
} catch {
  $result.access_application.status = "blocked"
  $result.access_application.error = $_.Exception.Message
  $result.limitations += "Access.Application collector failed: $($_.Exception.Message)"
}
}

$result | ConvertTo-Json -Depth 12
