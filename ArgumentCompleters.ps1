Function SQLDBSourceCompletion  {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    (get-item 'HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources','HKCU:\software\ODBC\ODBC.INI\ODBC Data Sources' -ErrorAction SilentlyContinue).Property.Where({$_ -notmatch " Files$" -and  $_ -like "$wordToComplete*" })  |
         Sort-Object | ForEach-Object {
                $tooltip =  (Get-ItemProperty -name $_ -path 'HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', 'HKCU:\software\ODBC\ODBC.INI\ODBC Data Sources' -ErrorAction SilentlyContinue).$_
                #New-CompletionResult "DSN=$_" $tooltip
                #completionText, listItemText, resultType, toolTip
                New-Object System.Management.Automation.CompletionResult "DSN=$_", "DSN=$_", ([System.Management.Automation.CompletionResultType]::ParameterValue) , $tooltip
    }
}

Function SQLDBSessionCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
 # $parameters = (Get-Item 'HKCU:\software\ODBC\ODBC.INI\ODBC Data Sources').property | Where-Object {$_ -notmatch " Files$"}
    $Global:DbSessions.Keys | Where-Object { $_ -like "$wordToComplete*" } | Sort-Object | ForEach-Object {
            #$tooltip = "$_"
            #New-CompletionResult $_ $tooltip
            New-Object System.Management.Automation.CompletionResult $_,$_, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
    }
}

Function SQLDBNameCompletion    {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $cmdnameused         = $commandAst.toString() -replace "^(.*?)\s.*$",'$1'
    if ($Global:DbSessions[$cmdnameused]) {
           $session      = $cmdnameused
    }
    else {  $session             = $(if($fakeBoundParameter['Session']) {$fakeBoundParameter['Session']} else {'Default'} ) }
    if ($DbSessions[$session] -is [System.Data.SqlClient.SqlConnection]) {
           $dbList = (Get-SQL -Session $session -SQL "SELECT name FROM sys.databases" -Quiet).name
    }
    else { $dblist = (Get-SQL -Session $session -SQL "show databases" -quiet).database}

     $dblist | Where-Object { $_ -like "$wordToComplete*" } | Sort-Object | ForEach-Object {
           # $tooltip = "$_"
           # New-CompletionResult $_ $tooltip
           New-Object System.Management.Automation.CompletionResult $_,$_, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
    }
}

Function SQLTableNameCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $cmdnameused         = $commandAst.toString() -replace "^(.*?)\s.*$",'$1'
    if ($Global:DbSessions[$cmdnameused]) {
           $session      = $cmdnameused
    }
    else {    $session      = $(if($fakeBoundParameter['Session']) {$fakeBoundParameter['Session']} else {'Default'} ) }
    if (-not $global:DbSessions[$session] -and $fakeBoundParameter['Connection'] ) {
        Get-SQL -Connection $fakeBoundParameter['Connection'] -Session $session | Out-Null
    }
    If ( $global:DbSessions[$session] ) {
        Get-SQL -Session $session -Showtables | Where-Object { $_ -like "*$wordToComplete*" } | Sort-Object | ForEach-Object {
                 $display    = $_ -replace "^\[(.*)\]$",'$1' -replace "^'(.*)'$",'$1'
                 $returnValue = """$_"""
                 New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList $returnValue,
                            $display , ([System.Management.Automation.CompletionResultType]::ParameterValue) ,$display
        }
    }
}

Function SQLFieldNameCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
  # $global:Cast = $commandAst ;
   # $global:fbp = $fakeBoundParameter
        $TableName       = $fakeBoundParameter['Table']
        $cmdnameused     = $commandAst.toString() -replace "^(.*?)\s.*$",'$1'
        if  ($Global:DbSessions[$cmdnameused]) {
              $session   = $cmdnameused
         }
        else {
        $session   = $(if($fakeBoundParameter['Session']) {$fakeBoundParameter['Session']} else {'Default'} ) }
        Get-SQL -Session $session -describe $TableName | Where-Object { $_.column_name -like "*$wordToComplete*" }  | Sort-Object -Property column_name |
            ForEach-Object {
                $display    = $_.COLUMN_NAME -replace "^\[(.*)\]$",'$1' -replace "^'(.*)'$",'$1'
                $returnValue = '"' + $_.COLUMN_NAME + '"'
                New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList $returnValue,
                            $display , ([System.Management.Automation.CompletionResultType]::ParameterValue) ,$display
    }
}


#In PowerShell 3 and 4 Register-ArgumentCompleter is part of TabExpansion ++. From V5 it is part of Powershell.core
if (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Connection' -ScriptBlock $Function:SQLDBSourceCompletion  #-Description 'Selects an ODBC Data Source'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Session'    -ScriptBlock $Function:SQLDBSessionCompletion #-Description 'Selects a session already opend by Get-SQL '
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'changeDB'   -ScriptBlock $Function:SQLDBNameCompletion    #-Description 'Selects an alternate Database available in a session'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Table'      -ScriptBlock $Function:SQLTableNameCompletion #-Description 'Complete Table names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Insert'     -ScriptBlock $Function:SQLTableNameCompletion #-Description 'Complete Table names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Describe'   -ScriptBlock $Function:SQLTableNameCompletion #-Description 'Complete Table names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Where'      -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Set'        -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Set'        -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'Select'     -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'GroupBy'    -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
 Register-ArgumentCompleter -CommandName 'Get-SQL' -ParameterName 'OrderBy'    -ScriptBlock $Function:SQLFieldNameCompletion #-Description 'Complete Field names'
}