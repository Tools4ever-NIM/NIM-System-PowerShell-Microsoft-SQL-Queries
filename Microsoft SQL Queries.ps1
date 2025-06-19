# Microsoft SQL.ps1 - IDM System PowerShell Script for Microsoft SQL Server via Queries.

$Log_MaskableKeys = @(
    'password'
)


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'connection_header'
                type = 'text'
                text = 'Connection'
				tooltip = 'Connection information for the database'
            }
            @{
                name = 'server'
                type = 'textbox'
                label = 'Server'
                description = 'Name of Microsoft SQL server'
                value = ''
            }
            @{
                name = 'database'
                type = 'textbox'
                label = 'Database'
                description = 'Name of Microsoft SQL database'
                value = ''
            }
            @{
                name = 'use_svc_account_creds'
                type = 'checkbox'
                label = 'Use credentials of service account'
                value = $true
            }
            @{
                name = 'username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                description = 'User account name to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                description = 'User account password to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'timeout'
                type = 'textbox'
                label = 'Query Timeout'
                description = 'Time it takes for the query to timeout'
                value = '1800'
            }
            @{
                name = 'connection_timeout'
                type = 'textbox'
                label = 'Connection Timeout'
                description = 'Time it takes for the connection to timeout'
                value = '3600'
            }
            @{
                name = 'session_header'
                type = 'text'
                text = 'Session Options'
				tooltip = 'Options for system session'
            }
			@{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                tooltip = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                tooltip = ''
                value = 1
            }
            @{
                name = 'table_header'
                type = 'text'
                text = 'Tables'
				tooltip = 'Query to Table mapping'
            }
			@{
                name = 'table_1_header'
                type = 'text'
                text = 'Table 1'
				tooltip = 'Table 1 Config'
            }
            @{
                name = 'table_1_name'
                type = 'textbox'
                label = 'Table 1 Name'
                description = ''
            }
            @{
                name = 'table_1_query'
                type = 'textbox'
                label = 'Table 1 Query'
                description = ''
				disabled = '!table_1_name'
                hidden = '!table_1_name'
            }
			
			@{
				name = 'table_2_header'
				type = 'text'
				text = 'Table 2'
				tooltip = 'Table 2 Config'
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_name'
				type = 'textbox'
				label = 'Table 2 Name'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_query'
				type = 'textbox'
				label = 'Table 2 Query'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}

			@{
				name = 'table_3_header'
				type = 'text'
				text = 'Table 3'
				tooltip = 'Table 3 Config'
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_name'
				type = 'textbox'
				label = 'Table 3 Name'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_query'
				type = 'textbox'
				label = 'Table 3 Query'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}

			@{
				name = 'table_4_header'
				type = 'text'
				text = 'Table 4'
				tooltip = 'Table 4 Config'
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_name'
				type = 'textbox'
				label = 'Table 4 Name'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_query'
				type = 'textbox'
				label = 'Table 4 Query'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}

			@{
				name = 'table_5_header'
				type = 'text'
				text = 'Table 5'
				tooltip = 'Table 5 Config'
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_name'
				type = 'textbox'
				label = 'Table 5 Name'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_query'
				type = 'textbox'
				label = 'Table 5 Query'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}

			@{
				name = 'table_6_header'
				type = 'text'
				text = 'Table 6'
				tooltip = 'Table 6 Config'
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_name'
				type = 'textbox'
				label = 'Table 6 Name'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_query'
				type = 'textbox'
				label = 'Table 6 Query'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}

			@{
				name = 'table_7_header'
				type = 'text'
				text = 'Table 7'
				tooltip = 'Table 7 Config'
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_name'
				type = 'textbox'
				label = 'Table 7 Name'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_query'
				type = 'textbox'
				label = 'Table 7 Query'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}

			@{
				name = 'table_8_header'
				type = 'text'
				text = 'Table 8'
				tooltip = 'Table 8 Config'
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_name'
				type = 'textbox'
				label = 'Table 8 Name'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_query'
				type = 'textbox'
				label = 'Table 8 Query'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}

			@{
				name = 'table_9_header'
				type = 'text'
				text = 'Table 9'
				tooltip = 'Table 9 Config'
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_name'
				type = 'textbox'
				label = 'Table 9 Name'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_query'
				type = 'textbox'
				label = 'Table 9 Query'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}

			@{
				name = 'table_10_header'
				type = 'text'
				text = 'Table 10'
				tooltip = 'Table 10 Config'
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_name'
				type = 'textbox'
				label = 'Table 10 Name'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_query'
				type = 'textbox'
				label = 'Table 10 Query'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}

			@{
				name = 'table_11_header'
				type = 'text'
				text = 'Table 11'
				tooltip = 'Table 11 Config'
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_name'
				type = 'textbox'
				label = 'Table 11 Name'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_query'
				type = 'textbox'
				label = 'Table 11 Query'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}

			@{
				name = 'table_12_header'
				type = 'text'
				text = 'Table 12'
				tooltip = 'Table 12 Config'
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_name'
				type = 'textbox'
				label = 'Table 12 Name'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_query'
				type = 'textbox'
				label = 'Table 12 Query'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}

			@{
				name = 'table_13_header'
				type = 'text'
				text = 'Table 13'
				tooltip = 'Table 13 Config'
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_name'
				type = 'textbox'
				label = 'Table 13 Name'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_query'
				type = 'textbox'
				label = 'Table 13 Query'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}

			@{
				name = 'table_14_header'
				type = 'text'
				text = 'Table 14'
				tooltip = 'Table 14 Config'
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_name'
				type = 'textbox'
				label = 'Table 14 Name'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_query'
				type = 'textbox'
				label = 'Table 14 Query'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}

			@{
				name = 'table_15_header'
				type = 'text'
				text = 'Table 15'
				tooltip = 'Table 15 Config'
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_name'
				type = 'textbox'
				label = 'Table 15 Name'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_query'
				type = 'textbox'
				label = 'Table 15 Query'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}

			@{
				name = 'table_16_header'
				type = 'text'
				text = 'Table 16'
				tooltip = 'Table 16 Config'
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_name'
				type = 'textbox'
				label = 'Table 16 Name'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_query'
				type = 'textbox'
				label = 'Table 16 Query'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}

			@{
				name = 'table_17_header'
				type = 'text'
				text = 'Table 17'
				tooltip = 'Table 17 Config'
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}
			@{
				name = 'table_17_name'
				type = 'textbox'
				label = 'Table 17 Name'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
				
			}
			@{
				name = 'table_17_query'
				type = 'textbox'
				label = 'Table 17 Query'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}

			@{
				name = 'table_18_header'
				type = 'text'
				text = 'Table 18'
				tooltip = 'Table 18 Config'
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_name'
				type = 'textbox'
				label = 'Table 18 Name'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_query'
				type = 'textbox'
				label = 'Table 18 Query'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}

			@{
				name = 'table_19_header'
				type = 'text'
				text = 'Table 19'
				tooltip = 'Table 19 Config'
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_name'
				type = 'textbox'
				label = 'Table 19 Name'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_query'
				type = 'textbox'
				label = 'Table 19 Query'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}

			@{
				name = 'table_20_header'
				type = 'text'
				text = 'Table 20'
				tooltip = 'Table 20 Config'
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_name'
				type = 'textbox'
				label = 'Table 20 Name'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_query'
				type = 'textbox'
				label = 'Table 20 Query'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}

        )
    }

    if ($TestConnection) {
        Open-MsSqlConnection $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log verbose "Done"
}


function Idm-OnUnload {
    Close-MsSqlConnection
}

#
# CRUD functions
#

$ColumnsInfoCache = @{}

$SqlInfoCache = @{}

function Fill-SqlInfoCache {
    param (
        [switch] $Force,
        [string] $Query,
        [string] $Class,
        [string] $Timeout = 3600
    )

    # Refresh cache
    $sql_command = New-MsSqlCommand $Query
    $sql_command.CommandTimeout = $Timeout
    $result = (Invoke-MsSqlCommand $sql_command) | Get-Member -MemberType Properties | Select-Object Name

    Dispose-MsSqlCommand $sql_command

    $objects = New-Object System.Collections.ArrayList
    $object = @{}
    # Process in one pass
    foreach ($row in $result) {
            $object = @{
                full_name = $Class
                type      = 'Query'
                columns   = New-Object System.Collections.ArrayList
            }

        $object.columns.Add(@{
            name           = $row.Name
            is_primary_key = $false
            is_identity    = $false
            is_computed    = $false
            is_nullable    = $true
        }) | Out-Null
    }

    if ($object.full_name -ne $null) {
        $objects.Add($object) | Out-Null
    }
    @($objects)
}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $connection_params = ConvertFrom-Json2 $SystemParams

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # Get all tables and views in database
            #

            Open-MsSqlConnection $SystemParams
            
            #
            # Output list of supported operations per table/view (named Class)
            #
            for (($i = 0), ($j = 0); $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name".length -gt 0)
                {
                    @(
                        [ordered]@{
                            Class = $connection_params."table_$($i)_name"
                            Operation = 'Read'
                            'Source type' = 'Query'
                            'Primary key' = ''
                            'Supported operations' = 'R'
                            'Query' = $connection_params."table_$($i)_query"
                        }
                    )
                    }
            }
        }
        else {
            # Purposely no-operation.
        }

    }
    else {

        if ($GetMeta) {
            #
            # Get meta data
            #

            Open-MsSqlConnection $SystemParams

            @()

        }
        else {
            #
            # Execute function
            #

            Open-MsSqlConnection $SystemParams

            for (($i = 0), ($j = 0); $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name" -eq $class)
                {
                    $class_query = ($connection_params."table_$($i)_query" -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join ' '
                    break
                }
            }
            
            $sql_command = New-MsSqlCommand

            switch ($Operation) {
                'Read' {
                    $sql_command.CommandText = $class_query
                    break
                }
            }

            if ($sql_command.CommandText) {
                $deparam_command = DeParam-MsSqlCommand $sql_command

                LogIO info ($deparam_command -split ' ')[0] -In -Command $deparam_command

                if ($Operation -eq 'Read') {
                    # Streamed output
                    Invoke-MsSqlCommand $sql_command $deparam_command
                }
                else {
                    # Log output
                    $rv = Invoke-MsSqlCommand $sql_command $deparam_command
                    LogIO info ($deparam_command -split ' ')[0] -Out $rv

                    $rv
                }
            }

            Dispose-MsSqlCommand $sql_command

        }

    }

    Log verbose "Done"
}


#
# Helper functions
#

function New-MsSqlCommand {
    param (
        [string] $CommandText
    )

    New-Object System.Data.SqlClient.SqlCommand($CommandText, $Global:MsSqlConnection)
}


function Dispose-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand
    )

    $SqlCommand.Dispose()
}


function AddParam-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand,
        $Param
    )

    $param_name = "@param$($SqlCommand.Parameters.Count)_"
    $SqlCommand.Parameters.AddWithValue($param_name, $Param) | Out-Null

    return $param_name
}


function DeParam-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand
    )

    $deparam_command = $SqlCommand.CommandText

    foreach ($p in $SqlCommand.Parameters) {
        $value_txt = 
            if ($p.Value -eq [System.DBNull]::Value) {
                'NULL'
            }
            else {
                switch ($p.SqlDbType) {
                    { $_ -in @(
                        [System.Data.SqlDbType]::Char
                        [System.Data.SqlDbType]::Date
                        [System.Data.SqlDbType]::DateTime
                        [System.Data.SqlDbType]::DateTime2
                        [System.Data.SqlDbType]::DateTimeOffset
                        [System.Data.SqlDbType]::NChar
                        [System.Data.SqlDbType]::NText
                        [System.Data.SqlDbType]::NVarChar
                        [System.Data.SqlDbType]::Text
                        [System.Data.SqlDbType]::Time
                        [System.Data.SqlDbType]::VarChar
                        [System.Data.SqlDbType]::Xml
                    )} {
                        "'" + $p.Value.ToString().Replace("'", "''") + "'"
                        break
                    }
        
                    default {
                        $p.Value.ToString().Replace("'", "''")
                        break
                    }
                }
            }

        $deparam_command = $deparam_command.Replace($p.ParameterName, $value_txt)
    }

    # Make one single line
    @($deparam_command -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join ' '
}


function Invoke-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand,
        [string] $DeParamCommand
    )

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: ?
    function Invoke-MsSqlCommand-ExecuteReader {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )
        Log verbose "Executing Query: $($SqlCommand.CommandText)"
        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            $hash_table = [ordered]@{}

            foreach ($column_name in $column_names) {
                $hash_table[$column_name] = ""
            }

#           $obj = [PSCustomObject]$hash_table
            $obj = New-Object -TypeName PSObject -Property $hash_table

            # Read data
            if($data_reader.HasRows) {
                while ($data_reader.Read()) {
                    foreach ($column_name in $column_names) {
                        $obj.$column_name = if ($data_reader[$column_name] -is [System.DBNull]) { $null } else { $data_reader[$column_name] }
                    }

                    # Output data
                    $obj
                }
            } else { [PSCustomObject]$hash_table }
        }

        $data_reader.Close()
    }

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: 16.7 s
    function Invoke-MsSqlCommand-ExecuteReader00 {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            # Initialize result
            $hash_table = [ordered]@{}

            for ($i = 0; $i -lt $column_names.Count; $i++) {
                $hash_table[$column_names[$i]] = ''
            }

            $result = New-Object -TypeName PSObject -Property $hash_table

            # Read data
            while ($data_reader.Read()) {
                foreach ($column_name in $column_names) {
                    $result.$column_name = $data_reader[$column_name]
                }

                # Output data
                $result
            }

        }

        $data_reader.Close()
    }

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: 01:11.9 s
    function Invoke-MsSqlCommand-ExecuteReader01 {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_reader = $SqlCommand.ExecuteReader()
        $field_count = $data_reader.FieldCount

        while ($data_reader.Read()) {
            $hash_table = [ordered]@{}
        
            for ($i = 0; $i -lt $field_count; $i++) {
                $hash_table[$data_reader.GetName($i)] = $data_reader.GetValue($i)
            }

            # Output data
            New-Object -TypeName PSObject -Property $hash_table
        }

        $data_reader.Close()
    }

    # Non-streaming (data stored in $data_table)
    # ERAM dbo.Files (426.977 rows) execution time: 15.5 s
    function Invoke-MsSqlCommand-DataAdapter-DataTable {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_adapter = New-Object System.Data.SqlClient.SqlDataAdapter($SqlCommand)
        $data_table   = New-Object System.Data.DataTable
        $data_adapter.Fill($data_table) | Out-Null

        # Output data
        $data_table.Rows

        $data_table.Dispose()
        $data_adapter.Dispose()
    }

    # Non-streaming (data stored in $data_set)
    # ERAM dbo.Files (426.977 rows) execution time: 14.8 s
    function Invoke-MsSqlCommand-DataAdapter-DataSet {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_adapter = New-Object System.Data.SqlClient.SqlDataAdapter($SqlCommand)
        $data_set     = New-Object System.Data.DataSet
        $data_adapter.Fill($data_set) | Out-Null

        # Output data
        $data_set.Tables[0]

        $data_set.Dispose()
        $data_adapter.Dispose()
    }

    if (! $DeParamCommand) {
        $DeParamCommand = DeParam-MsSqlCommand $SqlCommand
    }

    Log debug $DeParamCommand

    try {
        Invoke-MsSqlCommand-ExecuteReader $SqlCommand
    }
    catch {
        Log error "Query Failure: $_"
        throw $_
    }

    Log debug "Done"
}


function Open-MsSqlConnection {
    param (
        [string] $ConnectionParams
    )

    $connection_params = ConvertFrom-Json2 $ConnectionParams

    $cs_builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder

    # Use connection related parameters only
    $cs_builder['Data Source']     = $connection_params.server
    $cs_builder['Initial Catalog'] = $connection_params.database
    $cs_builder["Connect Timeout"] = $connection_params.connection_timeout
    
    if ($connection_params.use_svc_account_creds) {
        $cs_builder['Integrated Security'] = 'SSPI'
    }
    else {
        $cs_builder['User ID']  = $connection_params.username
        $cs_builder['Password'] = $connection_params.password
    }
    
    $connection_string = $cs_builder.ConnectionString

    if ($Global:MsSqlConnection -and $connection_string -ne $Global:MsSqlConnectionString) {
        Log verbose "MsSqlConnection connection parameters changed"
        Close-MsSqlConnection
    }

    if ($Global:MsSqlConnection -and $Global:MsSqlConnection.State -ne 'Open') {
        Log warn "MsSqlConnection State is '$($Global:MsSqlConnection.State)'"
        Close-MsSqlConnection
    }

    if ($Global:MsSqlConnection) {
        Log verbose "Reusing MsSqlConnection"
    }
    else {
        Log verbose "Opening MsSqlConnection '$connection_string'"

        try {
            $connection = New-Object System.Data.SqlClient.SqlConnection($connection_string)
            $connection.Open()

            $Global:MsSqlConnection       = $connection

            $Global:ColumnsInfoCache = @{}
            $Global:SqlInfoCache = @{}
        }
        catch {
            Log error "Connection Failure: $($_)"
            throw $_
        }

        Log verbose "Done"
    }
}


function Close-MsSqlConnection {
    if ($Global:MsSqlConnection) {
        Log verbose "Closing MsSqlConnection"

        try {
            $Global:MsSqlConnection.Close()
            $Global:MsSqlConnection = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log verbose "Done"
    }
}
