[CmdletBinding()]
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments","")]
Param()

 Describe "Connect to and query Access Database " {

    BeforeAll {
        $sessionName   = "ACCESS"
        $ACCconn       = ".\Database1.accdb"
        $tableName     = "TestData"
        $ArbitrarySQL  = "SELECT * from $tableName"
        $fieldname1    = "Extension"
        $fieldname2    = "Length"
        $null = Get-SQL  -Access  -Connection $ACCconn -Session $sessionName -ForceNew
    }

    It "Creates a PowerShell alias, matching the session name '$sessionName'" {
        {Get-Alias -Name $sessionName}                                                                | Should -not -throw
        (invoke-command -ScriptBlock ([scriptblock]::Create("$sessionname")) ).database               | Should -Be (Resolve-Path $ACCconn).Path.Trim()
    }
    It "Creates an open session in `$DBSessions, named '$sessionName'" {
        $DbSessions["$sessionName"].State                                                             | Should -Be "Open"
    }
    It "Can show tables in the database" {
         (Get-SQL -Session $sessionName -ShowTables                                           ).Count | Should -BeGreaterThan 0
    }
    It "Can describe the fields in the table $tableName" {
         (Get-SQL -Session $sessionName -Describe $tableName                                  ).Count | Should -BeGreaterThan 0
    }
    It "Can return the [whole] table $tableName" {
         (Get-SQL -Session $sessionName -Quiet -Table $tableName                              ).Count | Should -BeGreaterThan 0
    }
    It "Can return the [whole] table $tableName and capture the data table in a variable " {
        [void](Get-Sql -Session $sessionName -Quiet -Table $tableName  -OutputVariable Table  )
        $table.GetType().fullname                                                                     | Should -Be "System.Data.DataTable"
    }
    It "Can run abritrary SQL as passed as via the pipe" {
        ($ArbitrarySQL |   Get-SQL -Session $sessionName -Quiet                               ).Count | Should -BeGreaterThan 0
    }
    It "Can run abritrary SQL as passed as a parameter" {
        ( Get-SQL -Session $sessionName -Quiet $ArbitrarySQL                                   ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters" {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT 500 ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters, and values for where condition Piped " {
         (500,1000 , 10000 |
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT     ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters and where condition piped " {
         ("> 500","> 1000",">= 10000" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2         ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct and -OrderBy parameters and WHERE... clause piped " {
        ("Where  Length >500 ","Where Length >1000","Where Length >= 10000" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1                            ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with the WHERE... clause piped but no -Select, -Distinct or -OrderBy " {
        ("Where  Length >500 ","Where Length >1000","Where Length >= 10000" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with multiple fields in -Select and -OrderBy" {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Name",$fieldname1 -OrderBy $fieldname1,$fieldname2                   ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with -Select holding a 'Top' clause " {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Top 5 *" -OrderBy $fieldname1,$fieldname2                            ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with a different final clause (e.g. 'order by') as a parameter " {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName "order by $fieldname1 "                                                       ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT query with a different final clause piped " {
        ("order by $fieldname1 " |
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT ... WHERE ... LIKE query with 'naked' syntax and translate * as a wildcard" {
        ( Get-SQL -Session $sessionName -Select Name,Length,LastWriteTime -from TestData  -Where Extension -like ".ps*" -Quiet                 ).Count | Should -BeGreaterThan 0
    }
    It "Can run a SELECT Query with a date object as a parameter, -GroupBy and both fieldName & aggreate function in -Select " {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Where "CreationTime" -LT ([datetime]::Now).AddDays(-3) `
               -select $fieldname1,"Count(*) as total" -GroupBy $fieldname1                                                                    ).Count | Should -BeGreaterThan 0
    }
    It "Can INSERT rows via the pipeline or a parameter"      {
          $dirEntry = Get-Item (Get-Command -name powershell).Source | Select-Object -Property * -ExcludeProperty mod*
          $dirEntry, $dirEntry | Get-sql -Session $sessionName -Insert $tableName
          Get-SQL -Session $sessionName -Insert $tableName $dirEntry
        ( Get-SQL -Session $sessionName -Table  $tableName -Where "PSPath" -EQ $dirEntry.PSPath -Quiet                                         ).Count | Should -BeGreaterThan 0
    }
    It "Can DELETE rows from a table "      {
          $dirEntry = Get-Item (Get-Command -name powershell).Source | Select-Object -Property * -ExcludeProperty mod*
          $dirEntry, $dirEntry | Get-sql -Session $sessionName -Insert $tableName
          Get-SQL -Session $sessionName -Table  $tableName -where "PSPath" -EQ $dirEntry.PSPath -Delete -Confirm:$false
        ( Get-SQL -Session $sessionName -Table  $tableName -where "PSPath" -EQ $dirEntry.PSPath -Quiet                                         ).Count | Should -Be 0
    }
    It "Can SET values in a row in a table"   {
          $old = Get-SQL -Session $sessionName -Table  $tableName -Select "top 1 *" -Quiet
          Get-SQL        -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc `
                            -set "Attributes" -Values "Modified" -Confirm:$false
          $new = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc -Quiet
          $new.PSPath     | Should -Be $old.PSPath
          $new.Attributes | Should -Be "Modified"
          Get-SQL        -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc `
                            -set "Attributes" -Values  $old.Attributes -Confirm:$false
          $end = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc -Quiet
          $end.attributes | Should -Be $old.Attributes
    }

    AfterAll {Get-Sql -Session $sessionName -Close }
 }
