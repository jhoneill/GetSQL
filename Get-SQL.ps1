if (-not $Global:DbSessions ) { $Global:DbSessions = @{}  }

function Get-SQL {
    <#
      .Synopsis
        Queries an ODBC, SQLite or SQL Server database
      .Description
        Get-SQL queries SQL databases using either ODBC, the ADO driver for SQLite or the native SQL-Server client.
        Connections to databases are kept open and reused to avoid the need to make connections for every query,
        but the first time the command is run it needs a connection string; this come from $DefaultDBConnection.
        (e.g. set in your Profile) rather than being passed as a parameter: if it is set you can run
        sql "Select * From Customers"
        without any other setup, PowerShell will assume "sql" means "GET-SQL" if there is no other command named SQL.

        Get-SQL -Connection allows a connection to be specified explicitly; -MsSQLserver forces the use of
        the native SQL Server driver, -lite allows the file name of a SQLite Database to be used
        and -Excel or -Access allow a file name to be used without converting it into an ODBC connection string.

        Multiple named sessions may be open concurrently, and the global variable $DbSessions holds objects
        for each until Get-SQL is run with -Close. Note that you can run a query and make and/or close
        the connection in a single command. However, if you pipe the output into a command like
        Select-Object -First 2 then when Get-SQL is stopped by the downstream command it is unable to
        close the connection.

        Get-Sql will also build simple queries, for example
        Get-SQL -Table Authors
        Will run the "Select * from Authors" and a condition can be specified with
        Get-SQL -Table Authors -Where Name -like "*Smith"
        Get-SQL -ShowTables will show the available tables, and Get-SQL -Describe Authors will show the design of the table.

        Argument completers fill in names of ODBC connections, databases, tables, and columns where needed.
      .Parameter SQL
        A SQL statement. If other parameters (such as -Table, or -Where) are provided, it becomes the end of the SQL statement.
        If no statement is provided, or none can be built from the other parameters, Get-SQL returns information about the connection.
      .Parameter Connection
        An ODBC connection string or an Access, Excel, or SQLite file name or the name of a SQL Server
        It can be in the form "DSN=LocallyDefinedDSNName;" or
        "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.1.234;PORT=3306;DATABASE=xstreamline;UID=johnDoe;PWD=password;"
        A default connection string can be set in in $DBConnection so that you can just run "Get-SQL " «SQL Statement» ".
      .Parameter Excel
        Specifies that the string in -Connection is an Excel file path to be converted into a connection string.
      .Parameter Access
        Specifies that the string in -Connection is an Access file path to be converted into a connection string.
      .Parameter Lite
        Specifies the SQLite driver should be used and the string in -Connection may be the path to a SQLite file.
      .Parameter MsSQLserver
        Specifies the SQL Native client should be used and string in -Connection may be the name of a SQL Server.
      .Parameter Session
        Allows a database connection to be Identified by name: this sets the name used in the global variable $DBSessions.
        In addition, an alias is added: for example, if the session is named "F1" you can use the command F1 in place of Get-SQL -Session F1
      .Parameter ForceNew
        If specified, makes a new connection for the default or named session.
        If a connection is already established, -ForceNew is required to change the connection string.
      .Parameter ChangeDB
        For SQL server and ODBC sources which support it (like MySQL) switches to a different database at the same server.
      .Parameter Close
        Closes a database connection. Note this is run in the "end" phase of the command. If Get-SQL is stopped by another command
        in the pipeline (for example Select-object -first ) then it may not close the connection, so although this command can be
        combined with a select query, care is needed to ensure it is not defeated by another command in the same pipeline.
      .Parameter Table
        Specifies a table to select or delete from or to update.
      .Parameter Where
        If specified, applies a SQL WHERE condition to the selected table. -Where specifies the field and the text in -SQL supplies the condition.
      .Parameter GT
        Used with -Where specifies the > operator should be used, with the operand for the condition found in -SQL.
      .Parameter GE
        Used with -Where specifies the >= operator should be used, with the operand for the condition found in -SQL.
      .Parameter EQ
        Used with -Where specifies the = operator should be used, with the operand for the condition found in -SQL.
      .Parameter NE
        Used with -Where specifies the <> operator should be used, with the operand for the condition found in -SQL.
      .Parameter LE
        Used with -Where specifies the <= operator should be used, with the operand for the condition found in -SQL.
      .Parameter LT
        Used with -Where specifies the < operator should be used, with the operand for the condition found in -SQL.
      .Parameter Like
        Used with -Where specifies the Like operator should be used, with the operand for the condition found in -SQL. "*" in -SQL will be replaced with "%".
      .Parameter NotLike
        Used with -Where specifies the Not Like operator should be used, with the operand for the condition found in -SQL. "*" in -SQL will be replaced with "%".
      .Parameter Select
        If Select is omitted, -Table TableName will result in "SELECT * FROM TableName".
        Select specifies field-names (or other text) to use in place of "*".
      .Parameter Distinct
        Specifies that "SELECT DISTINCT ..." should be used in place of "SELECT ...".
      .Parameter OrderBy
        Specifies fields to be used in a SQL ORDER BY clause added at the end of the query.
      .Parameter Delete
        If specified, changes the query from a SELECT to a DELETE. This allows a query to be tested as a SELECT before adding -Delete to the command.
        -Delete requires a WHERE clause and not all ODBC drivers support deletion.
      .Parameter Set
        If specified, changes the query from a Select to a Update -Set Specifies the field(s) to be updated.
        -Set requires a WHERE clause.
      .Parameter Values
        If -Set is specified, -Values contains the new value(s) for the fields being updated.
      .Parameter Insert
        Specifies a table to insert into. The SQL parameter should contain a hash table or PSObject which holding the data to be inserted.
      .Parameter DateFormat
        Allows the format applied to Dates to be inserted to be changed if a service requires does not follow standard conventions.
      .Parameter GridView
        If specified, sends the output to gridview instead of the PowerShell console.
      .Parameter GroupBy
        If specified, adds a group by clause to a select query; in this case the SELECT clause needs to contain fields suitable for grouping.
      .Parameter Describe
        Returns a description of the specified table - note that some ODBC providers don't support this.
      .Parameter ShowTables
        If specified, returns a list of tables in the current database - note that some ODBC providers don't support this.
      .Parameter Paste
        If specified, takes an SQL statement from the clipboard.
        Line breaks and any text before SELECT , UPDATE or DELETE will be removed.
      .Parameter Quiet
        If specified, suppresses printing of the console message saying how many rows were returned.
      .Parameter OutputVariable
         Behaves like the common parameters errorVariable, warningvariable etc.to pass back a table object instead of an array of data rows.
      .Example
        Get-SQL -MsSQLserver -Connection "server=lync3\rtclocal;database=rtcdyn; trusted_connection=true;" -Session Lync
        Creates a new session named "LYNC" to the rtcdyn database on the Rtclocal SQL instance on server Lync
      .Example
        Get-SQL -Session LR -Connection "DSN=LR" -Quiet -SQL $SQL
        Runs the SQL in $SQL - if the Session LR already exists it will be used, otherwise it will be created to the ODBC source "LR"
        Note that a script should always name a its session(s), something else may already have set the defualt session
      .Example
        Get-Sql -showtables *dataitem
        Gives a list of tables on the default connection that end with "dataitem"
      .Example
        Get-SQL -Session f1 -Excel  -Connection C:\Users\James\OneDrive\Public\F1\f1Results.xlsx -showtables
        Creates a new connection named F1 to an Excel file, and shows the tables available.
      .Example
        f1  -Insert "[RACES]" @{RaceName = $raceName, RaceDate = $racedate.ToString("yyyy-MM-dd") }
        Uses the automatically created alias "f1" which was created in the previous example to insert a row of data into the "Races" Table
      .Example
        Get-SQL -Session F1 -Table "[races]"  -Set "[poleDriver]" -Values $PoleDriver -SQL "WHERE RaceDate = $rd" -Confirm:$false
        Updates the races table in the "F1" session, setting the value in the column "PoleDriver" to the contents of
        the variable $PoleDriver, in those rows where the RaceDate = $RD. This time the session is explicitly specified
        (using aliases is OK at the command line but not in scripts especially if the alias is created by a command run in the script)
        Changes normally prompt the user to confirm but here -Confirm:$false  prevents it
      .Example
        "CREATE USER 'johndoe' IDENTIFIED BY 'password'" , "GRANT ALL PRIVILEGES ON *.* TO 'johndoe'@'%' WITH grant option"  | Get-SQL
        Pipes two commands into the default connection, giving a new mySql user full access to all tables in all databases
      .Example
        Get-Sql -paste -gridview
        Runs the query currently in the clipboard against the default existing and outputs to the Gridview
      .Example
        SQL -table catalog_dataitem -select dataStatus -distinct -orderBy dataStatus -gridView
        Builds the query " SELECT DISTINCT dataStatus FROM catalog_dataitem ORDER BY dataStatus",
        runs it against the default existing connection and displays the results in a grid.
      .Example
        [void](Get-sql $sql  -OutputVariable Table)
        PowerShell unpacks Datatable objects into rows; so anything which needs a DataTable object cannot get it with
        $table = Get-Sql $sql
        because $table will contain an Array of DataRow objects, not a single DataTable.
        To get round this Get-SQL has -OutputVariable which behaves like the common parameters errorVariable, warningvariable etc.
        (using the Name of the variable 'Table' not its value '$table' as the parameter value)
        After running the command, the variable in the scope where the command is run contains the DataTable object.
        Usually the datarow objects will not be required, so the output can be cast to a void or piped to Out-Null.
    #>
    [CmdletBinding(DefaultParameterSetName='Describe',SupportsShouldProcess=$true,ConfirmImpact="High")]
    param   (
        [parameter(Position=0, ValueFromPipeLine=$true)]
        $SQL,
        [parameter(Position=1)][ValidateNotNullOrEmpty()]
        [string]$Connection  = $global:DefaultDBConnection ,
        [ValidateNotNullOrEmpty()]
        [string]$Session = "Default",
        [parameter(Position=2)]
        [alias('Use')]
        [string]$ChangeDB,
        [alias('Renew')]
        [switch]$ForceNew  ,
        [parameter(ParameterSetName="Paste")]
        [parameter(ParameterSetName="Describe")]
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [alias('g')][switch]$GridView,
        [parameter(ParameterSetName="Describe")]
        [alias('d')][string]$Describe,
        [parameter(ParameterSetName="ShowTables" , Mandatory=$true)]
        [switch]$ShowTables,
        [parameter(ParameterSetName="Paste"      , Mandatory=$true)]
        [switch]$Paste,
        [parameter(ParameterSetName="UpdateWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(ParameterSetName="DeleteWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(ParameterSetName="SelectWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(ParameterSetName="Update"     , Mandatory=$false)]
        [parameter(ParameterSetName="Delete"     , Mandatory=$false)]
        [parameter(ParameterSetName="Select"     , Mandatory=$false)]
        [alias('from','update')][string]$Table,
        #region Parameters for queries with a WHERE clause
        [parameter(ParameterSetName="UpdateWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(ParameterSetName="DeleteWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(ParameterSetName="SelectWhere", Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]$Where,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$GT,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$GE,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$EQ,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$NE,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$LE,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$LT,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$Like,
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$NotLike,
        #endregion
        #Parameters for SELECT Queries
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [alias('Property')][string[]]$Select,
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$Distinct,
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [string[]]$OrderBy,
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [String[]]$GroupBy,
        #Parameters for Delete queries
        [parameter(ParameterSetName="DeleteWhere", Mandatory=$true)]
        [parameter(ParameterSetName="Delete"     , Mandatory=$true)]
        [switch]$Delete,
        #Parameters for Update queries
        [parameter(ParameterSetName="UpdateWhere", Mandatory=$true)]
        [parameter(ParameterSetName="Update"     , Mandatory=$true)]
        [string[]]$Set,
        [parameter(ParameterSetName="UpdateWhere", Mandatory=$true,Position=1)]
        [parameter(ParameterSetName="Update"     , Mandatory=$true,Position=1)]
        [Object[]]$Values,
        #Parameters for INSERT Queries
        [parameter(ParameterSetName="Insert"     , Mandatory=$true)]
        [alias('into')][string]$Insert,
        [parameter(ParameterSetName="Insert")]
        [parameter(ParameterSetName="Update")]
        [parameter(ParameterSetName="UpdateWhere")]
        [parameter(ParameterSetName="DeleteWhere")]
        [parameter(ParameterSetName="SelectWhere")]
        [String]$DateFormat   = "'\''yyyy'-'MM'-'dd HH':'mm':'ss'\''",
        [parameter(ParameterSetName="Paste")]
        [parameter(ParameterSetName="Describe")]
        [parameter(ParameterSetName="Select")]
        [parameter(ParameterSetName="SelectWhere")]
        [switch]$Quiet,
        [switch]$MsSQLserver,
        [switch]$Lite,
        [switch]$Access,
        [switch]$Excel,
        [String]$OutputVariable,
        [switch]$Close
    )
    begin   {
        #Prepare  session, if needed, and leave it in the global variable DBSessions - a hash table with Name and connection object
        #If the function was invoked with an Alias of "DB" and there is session named "DB" switch to using that session
        if   (("Default" -eq $Session) -and $Global:DbSessions[$MyInvocation.InvocationName]) {$Session = $MyInvocation.InvocationName}

        #if the session doesn't exist or we're told to force a new session, then create and open a session
        if   (  ($ForceNew)  -or  (  -not   $Global:DbSessions[$session]) -and -not $Close) {
            if     ($Lite -and $PSVersionTable.PSVersion.Major -gt 5 -and $IsMacOS ) {
                Add-Type -path (Join-Path $PSScriptRoot "linux-x64\System.Data.SQLite.dll" )
            }
            elseif ($Lite -and $PSVersionTable.PSVersion.Major -gt 5 -and $linux )   {
                Add-Type -path (Join-Path $PSScriptRoot "linux-x64\System.Data.SQLite.dll" )
            }
            elseif ($lite -and -not [System.Environment]::Is64BitProcess) {
                Add-Type -path (Join-Path $PSScriptRoot "win-x86\System.Data.SQLite.dll" )
            }
            elseif ($lite ) {
                Add-Type -path (Join-Path $PSScriptRoot "win-x64\System.Data.SQLite.dll" )
            }
            #Catch -force to refresh instead of replace the current connection (e.g. Server has timed out )
            if (($ForceNew)      -and       $Global:DbSessions[$session] -and -not $PSBoundParameters.ContainsKey('Connection'))  {
                if     ($Global:DbSessions[$session].GetType().name -eq "SqlConnection"    ) {$MsSQLserver = $true}
                elseif ($Global:DbSessions[$session].GetType().name -eq "SQLiteConnection" ) {$Lite = $true}
            }
            #If -MSSQLServer  switch is used assume connection is a server if there is no = sign in the connection string
            if  ($MsSQLserver    -and       $Connection                  -and  $connection -notmatch "=") {
                $Connection = "server=$Connection;trusted_connection=true;timeout=60"
            }
            #If -Lite switch is used assume connection is a file if there is no = sign in the connection string, check it exists and build the connection string
            if  ($Lite           -and       $Connection                  -and  $connection -notmatch "=") {
              if (Test-Path -Path  $Connection)  {
                     $Connection  = "Data Source="+
                                    (Resolve-Path -Path $Connection -ErrorAction SilentlyContinue).Path + ";"
              }
              else { Write-Warning -Message "Can't create database connection: could not find $Connection" ; return}
            }
            #If the -Excel or Access switches are used, then the connection parameter is the path to a file, so check it exists and build the connection string
            if  ($Excel)                  {
              if (Test-Path -Path  $Connection)  {
                  $Connection  = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;ReadOnly=0;Dbq=" +
                                 (Resolve-Path -Path $Connection -ErrorAction SilentlyContinue).Path + ";"
              }
              else { Write-Warning -Message "Can't create database connection: could not find $Connection" ; return}
            }
            if  ($Access)                 {
              if (Test-Path -Path  $Connection)  {
                     $Connection  = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq="+
                                    (Resolve-Path -Path $Connection -ErrorAction SilentlyContinue).Path + ";"
              }
              else { Write-Warning -Message "Can't create database connection: could not find $Connection" ; return}
            }
            if (-not $Connection)         { Write-Warning -Message "A connection was needed but -Connection was not provided."; break}
            Write-Verbose -Message "Connection String is '$connection'"
            #Use different types for SQL server, SQLite and ODBC. They (and the logic) are almost interchangable.
            if  ($MsSQLserver)            { $Global:DbSessions[$Session] = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $Connection }
            elseif ($Lite)                { $Global:DbSessions[$Session] = New-Object -TypeName System.Data.SQLite.SQLiteConnection -ArgumentList $Connection }
            else                          { $Global:DbSessions[$Session] = New-Object -TypeName System.Data.Odbc.OdbcConnection     -ArgumentList $Connection
                                            $Global:DbSessions[$Session].ConnectionTimeout = 30
            }
            #Open our connection. NB, if 32 bit office is installed Excel, Access ETC have 32 bit ODBC drivers which need 32 bit Powershell not 64 bit.
            try                           { $Global:DbSessions[$Session].open() }
            catch                         { Write-Warning -Message "Error opening connection to '$Connection'"
                                            if (($Access -or $Excel) -and [System.Environment]::Is64BitProcess) {
                                                Write-Warning -Message  "This is 64-bit PowerShell, If Office is 32-bit you need to use 32 bit-PowerShell"}
                                            $Global:DbSessions[$Session] = $null
                                            break
            }
            #Create an alias which matches the connection name.
            if  ("Default" -eq $Session)  { $Global:DefaultDBConnection = $Connection }
            else                          { New-Alias -Name $Session -Value Get-SQL -Scope Global -Force}
        }
        if      ($ChangeDB)               { $Global:DbSessions[$Session].ChangeDatabase($ChangeDB) } #This method to change DB won't work with every provider
        if   (  ($Paste) -and (Get-Command -Name 'Get-Clipboard' -ErrorAction SilentlyContinue))  {
            #You could use [windows.clipboard]::GetText() - be warned this may not work in the older releases of the standard shell
            #For older versions of PowerShell I have a Get-Clipboard function which wraps this
            $SQL = (Get-Clipboard) -replace "^.*?(?=select|update|delete)","" -replace "[\n|\r]+"," "
        }
    }
    process {
        #If $table is specified make sure $SQL isn't empty otherwise we won't get to Select * from $Table; also make sure conditions allow for it to be zero!
        if   ($Table -and $null -eq $SQL) { $SQL = "   "}
        if   ($SQL.SQL)                   { $SQL = $SQL.SQL}
        if                    ($Describe) { #support -Describe [tablename] to descibe a table
            if ($Global:DbSessions[$Session].driver -match "SQORA" ) { #Oracle is special ...
                Get-SQL -Session $Session -Quiet -SQL  "select COLUMN_NAME, data_type as TYPE_NAME, data_length AS COLUMN_SIZE " +
                                                        " from user_tab_cols where table_name = '$Describe' order by COLUMN_NAME"
            }
            else  { #Remove any [] around the table name - because that's how .GetSchema() works ...
                $Describe = $Describe -replace "\[(.*)\]",'$1'
                # For some drivers .GetSchema() can get the columns for a single table. But the Excel driver can't, so get all columns and filter.
                if     ($Global:DbSessions[$Session].Driver -match "ACEODBC.DLL" ) {
                        $columns = $Global:DbSessions[$Session].GetSchema("Columns") | Where-Object {$_.TABLE_NAME -eq $Describe }
                }
                elseif ($Global:DbSessions[$Session].gettype().name -match "SqlConnection" ) {#SQL server uses slightly differnet syntax
                        $columns = $Global:DbSessions[$Session].GetSchema("Columns", @("%","%",$Describe,"%"))
                }
                else {  $columns = $Global:DbSessions[$Session].GetSchema("Columns", @("","",$Describe)) }
                if ($GridView) {$columns | Out-GridView -Title "Table $Describe"}
                else           {$columns | Select-Object -Property @{n="COLUMN_NAME";e={if ($_.Column_Name -match "\W") {"[$($_.Column_Name)]"} else {$_.column_Name} }},
                                                                TYPE_NAME, COLUMN_SIZE, IS_NULLABLE
                }
            }
        }
        elseif              ($Showtables) { #ODBC method to get tables won't work with every provider, but nor will executing "show tables". $SQL param becomes a filter
            if   ($Global:DbSessions[$Session].driver -match "SQORA" ) {#Oracle is special ...
                (Get-SQL  -Session $Session -Quiet -SQL  "select OBJECT_NAME from user_objects where object_type  IN ('VIEW','TABLE'); ").object_name |
                    Where-Object {$_ -like "$SQL*"}
            }
            else {$Global:DbSessions[$Session].GetSchema("Tables") | Where-Object {$Global:DbSessions[$Session].DataSource -ne "Access" -or $_.TABLE_TYPE -ne "SYSTEM TABLE"} |
                ForEach-Object {
                    if     ($_.TABLE_NAME -like "$SQL*" -and $_.TABLE_NAME -match "\W") {"[" + $_.TABLE_NAME + "]"}
                    elseif ($_.TABLE_NAME -like "$SQL*")                                {      $_.TABLE_NAME      }
                } | Sort-Object
            }
        }
        elseif           ($null -ne $SQL) { #$SQL holds any SQL which we can't (or don't want to( assemble from the cmdline, a whole statement or final clause
          ForEach        ($s     in $SQL) { #More than one statement/clause can be passed
            if ($Delete -or $Set -and -not $Table) { Write-Warning -Message "You must specifiy a table and where condition to use -Delete or -Set" ; return }
            if                            ($Table) { #If $Table was specified, build a Select, Delete or Update query
                #Support -table [tablename] -Where [ColumnName] -eq 99 and similar syntax.
                # -eq -ne and other operators are *switches*. The operand for = (etc.) is in $SQL so only Operator is allowed. Too complex to enforce this in Param() block!
                $opCount        =  (($Like, $EQ, $NE , $LT , $GT , $GE, $LE, $NotLike) -eq $true).Count
                #Can't have multiple operators, and operator requires -Where to be specified and a value in -SQL (-SQL usually implied in cmdline)
                if ((($opCount) -gt 1) -or  (($opCount -eq 1) -and -not $Where ) -or ($Where -and  "   " -eq  $s )) {
                    Write-Warning -Message  "You can't specify a where condition like that"
                    return
                }
                if  (($opCount) -eq 1) { #If we have an operator, column and value in $s turn $s into the condition (add the column name after)
                    #if the operand for -eq etc is a date format it for SQL
                    if ($s -is [datetime])  {
                        $s = $s.tostring($DateFormat)  #Default format has "'" this works for Excel inserts and SQL server.
                        if ($Global:DbSessions[$Session].Driver -eq "ACEODBC.DLL") {       #For Excel where needs # not quotes as date markers
                            $s = $s -replace "'","#"
                        }
                    } #if the operand for -eq etc is not a number or isn't wrrapped in quotes. Wrap it in quotes and double up the ' character
                    elseif (($s -notmatch "^\d+\.?\d*$") -and ($s -notmatch "^'.*'$"))
                                      {$s = "'" + ($s -replace "(?<!')'(?!')","''") +"'"  }
                    if          ($EQ) {$s =   " =  $s "            }
                    if          ($NE) {$s =   " <> $s "            }
                    if          ($GE) {$s =   " >= $s "            }
                    if          ($LE) {$s =   " <= $s "            }
                    if          ($GT) {$s =   " >  $s "            }
                    if          ($LT) {$s =   " <  $s "            }
                    if       (($Like)  -or ($NotLike) ) {   #for the like operators replace * wildcard with SQL % wildcard
                                       $s = $s -replace "\*","%"   }
                    if        ($Like) {$s =     " like $s "        }
                    if     ($NotLike) {$s = " not like $s "        }
                    #At the end of this $s holds the condition but not the column name
               }
                if           ($Delete) { #Support Delete queries  -Table [tableName] -Delete -where [Column] -eq [Value]
                    #A careless -Delete could wipe out a table - so insist on either -where [columnName] and a condition, or "Where blah blah" in $SQL
                  if ((($Where) -and $s) -or ($s -match "where\s+\w+")) {
                    if ($Where)     {$s = "DELETE FROM $Table WHERE $Where " + $S }
                    else            {$s = "DELETE FROM $Table "              + $S }
                  }
                  else {Write-Warning -Message "You must specifiy a where condition to use -Delete"; return }
               }
                elseif          ($Set) {
                  #Support update ... set queries -Table [tableName] -Set [Columns] -Values [values] -Where [Column] -EQ [Value]
                  #Don't allow set to modify all the rows (same logic as Delete)
                  if ( (  $Where  -and $s) -or ($s -match "where\s+\w+")) {
                    #We have a list of columns in Set and values for them need the same number of each - then build the set clause, wrapping text values in ''
                    if ($Set.Count  -ne  $Values.Count)          {Write-Warning -Message "Must have the same number of columns to set as values to set them to"; return }
                    $setList = ""
                    for  ($i = 0; $i  -lt  $set.count; $i++) {
                        if  (   $Values[$i] -is [datetime])  {
                            if ($Global:DbSessions[$Session].gettype().name -match "SQLiteConnection") {
                                    $Vi = [int]($Values[$i].Subtract([datetime]::UnixEpoch).TotalSeconds)
                            }
                            else {  $Vi = $Values[$i].tostring($DateFormat)}     #Default format has "'" this works for Excel, Access and SQL server.
                            $SetList = $SetList + $Set[$i] + "= " + $vi +"  ,"
                        }
                        #  Wrap text in ' and escape ' char
                        elseif ($Values[$i] -notmatch "^[\d\.]*$") {$SetList = $SetList + $Set[$i] + "='" + ($Values[$i] -replace "'","''") +"' ," }
                        else                                       {$SetList = $SetList + $Set[$i] + "= " +  $Values[$i] +"  ," }
                    }
                    #will have an extra "," at the end.
                    $setList = $setList -replace ",$",""
                    if   ($Where)   {$s = "UPDATE $Table SET $setList WHERE $Where " + $s }
                    else            {$s = "UPDATE $Table SET $setList "              + $s }
                }
                  else                {Write-Warning -Message "You must specifiy a where condition to use -Set"   ; return }
               }
                else                   {#If we're not updating or deleting and -Table was passed we must be selecting ....
                    if   (   $Select) {$SelectClause = ($Select -join ", ") + " FROM $Table " }
                    else              {$SelectClause =                      " * FROM $Table " }
                    if   (    $Where) {$SelectClause = $SelectClause   +      "WHERE $Where " } #note we need to have the "what" part of SQL. but SQL could be @("=10",">73") we'll run 2 queries
                    if   ( $Distinct) {$s = "SELECT DISTINCT "         +  $SelectClause + $s  }
                    else              {$s = "SELECT "                  +  $SelectClause + $s  }
                    if   (  $GroupBy) {$s = $s +    " GROUP BY "       + ($GroupBy -join ", ")}
                    if   (  $OrderBy) {$s = $s +    " ORDER BY "       + ($OrderBy -join ", ")}
                }
            }
            elseif                       ($Insert) {
            #Support -insert [IntoTableName] @{hashtable of fields and values}
                if     ($s -is [Hashtable]) {$index = $s.keys}
                elseif ($s -is [psobject] ) {$index = (Get-Member -InputObject $s -MemberType NoteProperty).Name }
                else                        { Write-Warning -Message "Can't build an Insert statement from $s. Pass a hashtable or a PSObject" ; return}
                $fieldsPart     = "  "
                $valuesPart     = "  "
                foreach ($name in $index) {
                    $fieldsPart = $fieldsPart  + $name + " , "
                    $v          = $s.$name
                    if     ($Global:DbSessions[$Session].gettype().name -match "SQLiteConnection") {
                        if ($v -is [datetime] )  {$v = [int]($v.Subtract([datetime]::UnixEpoch).TotalSeconds)}
                        if ($v -is [Boolean]  )  {$v = [int]$v}
                    }
                    #$DateFormat defaults to the standard date format which SQL dialects support, but it can be overridden for special cases
                    if     ($v -is [datetime] )  {$valuesPart = $valuesPart +         $v.tostring($DateFormat) + " , " }
                    elseif ($v -is [int]    -or
                            $v -is [float]  -or
                            $v -is [boolean]  )  {$valuesPart = $valuesPart +         $v.tostring()         +    " , " }
                    elseif ($v -match "^\d+$" )  {$valuesPart = $valuesPart +         $v.tostring()         +    " , " }
                    else                         {$valuesPart = $valuesPart +  "'" + ($v -replace "'","''") +   "' , " }
                }
                $s  = ("INSERT INTO {0} ({1}) VALUES ({2})" -f $Insert,($fieldsPart -replace ",\s*$",""),($valuesPart -replace ",\s*$",""))
                $s = $s -replace ",\s*,",", null ," -replace "(?<=[(,])\s*''\s*(?=[),])"," null " -replace ",\s*\)",", null)"
            }
            Write-Verbose -Message $s
            #Choose suitable data adapter object based on session type.
            if     ($Global:DbSessions[$Session].gettype().name -match "SqlConnection" )  {
               $da = New-Object    -TypeName System.Data.SqlClient.SqlDataAdapter -ArgumentList (
                        New-Object -TypeName System.Data.SqlClient.SqlCommand     -ArgumentList $s,$Global:DbSessions[$Session] )
            }
            elseif ($Global:DbSessions[$Session].gettype().name -match "SQLiteConnection" ) {
               $da = New-Object    -TypeName System.Data.SQLite.SQLiteDataAdapter -ArgumentList (
                        New-Object -TypeName System.Data.SQLite.SQLiteCommand     -ArgumentList $s,$Global:DbSessions[$Session] )
            }
            else  {
               $da = New-Object    -TypeName System.Data.Odbc.OdbcDataAdapter     -ArgumentList (
                        New-Object -TypeName System.Data.Odbc.OdbcCommand         -ArgumentList $s,$Global:DbSessions[$Session])

            }
            $dt       = New-Object -TypeName System.Data.DataTable
            #And finally we get to execute the SQL Statement.
            try  { if ((-not ($Set -or $Delete -or ($Insert -and $ConfirmPreference -ne "high"))) -or ($PSCmdlet.ShouldProcess("$Session database", $s)) ) {
                   $rows  = $da.fill($dt)
                   if (-not ($Quiet -or $Delete -or $Set -or $Insert)) {Write-host -Object ("" + [int]$rows + " row(s) returned")}
            }}
            catch {
               if($S)               { #if we get an error and -SQL was passed show the final SQL statement.
                  $e=$Global:error[0]
                  throw ( New-Object -TypeName "System.Management.Automation.ErrorRecord" `
                                     -ArgumentList (($e.exception.message -replace "^(.*])\s*","`$1`n") + "`n `n>>>  $S `n `n" ), $e.FullyQualifiedErrorId ,"ParserError" ,$e.TargetObject)
               }
               else                   { throw }
              }
            if   (($GridView) -and (($PSVersionTable.PSVersion.Major -GE 3)-or ($host.name -match "ISE" )) ) {$dt | Out-GridView -Title $s}
            else  {$dt}
            if ($OutputVariable) {Set-Variable -Scope 2 -Name $OutputVariable -Value $dt -Visibility Public}
            }
        }
        elseif              (-not $Close -and -not $Quiet) { #If $SQL, $table, $describe or $showtimes weren't included either we're opening a new connection, or we're checking or closing an existing one.
            $Global:DbSessions[$Session]
        }
    }
    end     {
        if ($Close -and $Global:DbSessions[$Session]) {
            $Global:DbSessions[$Session].close()
            $Global:DbSessions[$Session].dispose()
            $Global:DbSessions.Remove($Session)
            Remove-Item -Path (Join-Path -Path "Alias:\" -ChildPath $Session) -ErrorAction SilentlyContinue
        }
    }
}

function Hide-GetSQL {
<#
    .Synopsis
        Allows a command line with quote marks to passed into Get-SQL, can be used simply as ¬
    .Example
        ¬ select host,user from mysql.user
        Sends the command "select host,user from mysql.user" to the default ODBC session

#>
  Get-Sql -sql ($MyInvocation.line.substring($MyInvocation.OffsetInLine)) | Format-Table -AutoSize
}
Set-Alias -Name ¬ -Value Hide-GetSQL
