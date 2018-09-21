 $sessionName   = "XL" 
 $xlconn        = ".\dbtest.xlsx " 
 $tableName     = "[TestData$]"
 $ArbitrarySQL  = "SELECT * from $tableName"
 $fieldname1    =   "Extension" 
 $fieldname2    =   "Length" 
 Describe "Connect to and query Excel Spreadsheet " { 
    
    BeforeAll {$session = Get-SQL  -Excel  -Connection $xlconn -Session $sessionName -ForceNew } 
    
    It "Creates a PowerShell alias, matching the session name '$sessionName'" {
        {Get-Alias -Name $sessionName}                                                  | Should not throw 
        (invoke-command -ScriptBlock ([scriptblock]::Create("$sessionname")) ).database | Should be (Resolve-Path $xlconn).Path.Trim()
    }
    It "Creates an open session in `$DBSessions, named '$sessionName'" {
        $DbSessions["$sessionName"].State                                               | Should be "Open"
    }
    It "Can show tables in the database" {
         (Get-SQL -Session $sessionName -ShowTables                          ).Count    | Should beGreaterThan 0        
    }
    It "Can describe the fields in the table $tableName" {
         (Get-SQL -Session $sessionName -Describe $tableName                 ).Count    | Should beGreaterThan 0        
    }
    It "Can return the [whole] table $tableName" {
         (Get-SQL -Session $sessionName -Quiet -Table $tableName             ).Count    | Should beGreaterThan 0        
    }
        It "Can return the [whole] table $tableName and capture the data table in a variable " {
        [void](Get-Sql -Session $sessionName -Quiet -Table $tableName  -OutputVariable Table  )  
        $table.GetType().fullname                                                       | Should be "System.Data.DataTable" 
    }
    It "Can run abritrary SQL as passed as via the pipe" {
        ($ArbitrarySQL |   Get-SQL -Session $sessionName -Quiet              ).Count    | Should beGreaterThan 0    
    }
    It "Can run abritrary SQL as passed as a parameter" {
        (Get-SQL -Session $sessionName -Quiet $ArbitrarySQL                  ).Count    | Should beGreaterThan 0    
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters" { 
         (Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT 500 ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters, and values for where condition Piped " { 
         (500,1000 , 10000 | 
         Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT      ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters and where condition piped " { 
         ("> 500","> 1000",">= 10000" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2         ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select, -Distinct and -OrderBy parameters and WHERE... clause piped " { 
        ("Where  Length >500 ","Where Length >1000","Where Length >= 10000" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1                            ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with the WHERE... clause piped but no -Select, -Distinct or -OrderBy " { 
        ("Where  Length >500 ","Where Length >1000","Where Length >= 10000" | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with multiple fields in -Select and -OrderBy" { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Name",$fieldname1 -OrderBy $fieldname1,$fieldname2                   ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with -Select holding a 'Top' clause " { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Top 5 *" -OrderBy $fieldname1,$fieldname2                            ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with a different final clause (e.g. 'order by') as a parameter " { 
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName "order by $fieldname1 "                                                       ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT query with a different final clause piped " { 
        ("order by $fieldname1 " | 
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should beGreaterThan 0               
    }
    It "Can run a SELECT ... WHERE ... LIKE query with 'naked' syntax and translate * as a wildcard" { 
         (sql -Session $sessionName -Select Name,Length,LastWriteTime -from [TestData$]  -Where Extension -like ".ps*" -Quiet                  ).Count | Should beGreaterThan 0
    }
    It "Can run a SELECT Query with a date object as a parameter, -GroupBy and both fieldName & aggreate function in -Select " { 
      (Get-SQL -Session $sessionName -Quiet -Table $tableName -Where "CreationTime" -LT ([datetime]::Now).AddDays(-3) `
               -select $fieldname1,"Count(*) as total" -GroupBy $fieldname1                                                                    ).Count |  Should beGreaterThan 0        
    }
    It "Can INSERT rows into a table via the pipeline or a parameter"      {
         $dirEntry = Get-Item (Get-Command -name powershell).Source | select * 
         $dirEntry, $dirEntry | Get-sql -Session $sessionName -Insert $tableName  
         Get-sql -Session $sessionName -Insert $tableName $dirEntry  
         (Get-sql -Session $sessionName -table  $tableName -where "PSPath" -eq $dirEntry.PSPath -Quiet                                         ).Count | Should beGreaterThan 0    
    } 
    # Excel Driver doens not support "Can Delete a row from a table" 
    It "Can SET new values in a row in a table"   {
         $old = Get-SQL -Session $sessionName -Table  $tableName -Select "top 1 *" -Quiet 
         Get-SQL        -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc `
                            -set "Attributes" -Values "Modified" -Confirm:$false 
         $new = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc -Quiet 
         $new.PSPath     | Should be $old.PSPath
         $new.Attributes | Should be "Modified"
          Get-SQL        -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc `
                            -set "Attributes" -Values  $old.Attributes -Confirm:$false 
         $end = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Format(LastWriteTimeUtc)" -eq $old.LastWriteTimeUtc -Quiet 
         $end.attributes | Should be $old.Attributes                                                                
    } 

    AfterAll {Get-Sql -Session $sessionName -Close }
 }
