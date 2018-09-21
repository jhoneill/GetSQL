
$tableName    = "CallType"
$fieldName1   = "CallType"    #Must be a name used to test wild card
$fieldName2   = "CallTypeId"  #Test for values 2,3,4  
$dbName       = "LcsCDR" 
$sessionName  = "LcsCDR"
$sqlconn      = "bp1xeucc023" 
$End          = [datetime]::Now ; 
$Start        =   $End.AddHours(-1) 
$ArbitrarySQL = "exec dbo.CdrP2PSessionList @_StartTime ='" + $Start.ToString("yyyy-MM-dd HH:mm") + "', @_EndTime  ='" + $End.ToString("yyyy-MM-dd HH:mm") + "'"         

#Import-Module -Name GetSQL -Force        

Describe "Connect to and query SQL Server " { 
    
    BeforeAll {$session = Get-SQL  -MSSqlServer  -Connection $sqlconn -use $dbName -Session $sessionName -ForceNew } 
    
    It "Creates a PowerShell alias, matching the session name '$sessionName'" {
        {Get-Alias -Name $sessionName}                                                         | Should not throw 
        (invoke-command -ScriptBlock ([scriptblock]::Create("$sessionname")) ).database        | should be  $sessionName 
    }
    It "Creates an open session in `$DBSessions, named '$sessionName'" {
        $DbSessions["$sessionName"].State                                  | Should be "Open"
    }
    It "Can select a database using the -USE Alias" {
         $DbSessions["$sessionName"].database                              | Should be $dbName
    }
    It "Can show tables in the database" {
         (Get-SQL -Session $sessionName -ShowTables).count                 | Should beGreaterThan 0        
    }
    It "Can describe the fields in the table [$tableName]" {
         (Get-SQL -Session $sessionName -Describe $tableName).count        | Should beGreaterThan 0        
    }
    It "Can return the [whole] table [$tableName]" {
         (Get-SQL -Session $sessionName -Quiet -Table $tableName ).count   | Should beGreaterThan 0        
    }
    It "Can run abritrary SQL as passed as via the pipe" {
        ($ArbitrarySQL |   Get-SQL -Session $sessionName -Quiet ).Count | should beGreaterThan 0    
    }
    It "Can run abritrary SQL as passed as a parameter" {
        (Get-SQL -Session $sessionName -Quiet $ArbitrarySQL     ).Count | should beGreaterThan 0    
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters" { 
         (Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldName1 -Distinct -OrderBy $fieldName1 -Where $fieldName2 -gt 0 ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters, and values for where condition Piped " { 
         (2,3,4  | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldName1 -Distinct -OrderBy $fieldName1 -Where $fieldName2 -eq   ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters and where condition piped " { 
         ("=2","=3",">=4" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldName1 -Distinct -OrderBy $fieldName1 -Where $fieldName2       ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct and -OrderBy parameters and WHERE... clause piped " { 
        ("Where $fieldName2 =2","Where $fieldName2 =3","Where $fieldName2 >=4" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldName1 -Distinct -OrderBy $fieldName1                           ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with the WHERE... clause piped but no -Select, -Distinct or -OrderBy " { 
        ("Where $fieldName2 =2","Where $fieldName2 =3","Where $fieldName2 >=4" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                              ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with multiple fields in -Select and -OrderBy" { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldName1, $fieldName2 -OrderBy $fieldName1, $fieldName2           ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select holding a 'Top' clause " { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Top 5 *" -OrderBy $fieldName1,$fieldName2                           ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with a different final clause (e.g. 'order by') as a parameter " { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName "order by $fieldName1 "                                                      ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT query with a different final clause piped " { 
        ("order by $fieldName1 " | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                              ).count | should beGreaterThan 0               
    }
    It "Can run a SELECT ... WHERE ... LIKE query with 'naked' syntax and translate * as a wildcard" { 
         (    SQL -Session $sessionName -Quiet -Select CallType,CallTypeId -From CallType -Where CallType -Like audio*                                          ).count | should beGreaterThan 0 
    }
    It "Can run a SELECT query with a date object as a value for where, -GroupBy and both fieldName & aggreate function in -Select " { 
        ( Get-SQL -Session $sessionname -Quiet -Table "Registration" -Select RegistrarId,"Count(*) As total"  `
                -Where "RegisterTime" -GT ([datetime]::Today) -GroupBy "RegistrarId"                                                           ).count  | Should beGreaterThan 0
    }
    It "Can add a row to a table"      {} -Pending
    It "Can Delete a row from a table" {} -Pending 
    It "Can Change a row in a table"   {} -Pending 

    AfterAll {Get-Sql -Session $sessionName -Close }
}
