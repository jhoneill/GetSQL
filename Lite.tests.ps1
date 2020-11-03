
 Describe "Connect to and query Excel Spreadsheet " {

    BeforeAll {
        $sessionName   = "lite"
        $liteconn      = ".\TestData.sqlite"
        $tableName     = "F1Results"
        $ArbitrarySQL  = "SELECT * from $tableName"
        $fieldname1    = "Driver"
        $fieldname2    = "Points"
        $session = Get-SQL  -Lite -Connection $liteConn -Session $sessionName -ForceNew
    }

    It "Creates a PowerShell alias, matching the session name '$sessionName'" {
        {Get-Alias -Name $sessionName}                                                  | Should -not -throw
        (invoke-command -ScriptBlock ([scriptblock]::Create("$sessionname")) ).database | Should -be 'main'
    }
    It "Creates an open session in `$DBSessions, named '$sessionName'" {
        $DbSessions["$sessionName"].State                                               | Should -be "Open"
    }
    It "Can show tables in the database" {
         (Get-SQL -Session $sessionName -ShowTables                          ).Count    | Should -beGreaterThan 0
    }
    It "Can describe the fields in the table $tableName" {
         (Get-SQL -Session $sessionName -Describe $tableName                 ).Count    | Should -beGreaterThan 0
    }
    It "Can return the [whole] table $tableName" {
         (Get-SQL -Session $sessionName -Quiet -Table $tableName             ).Count    | Should -beGreaterThan 0
    }
        It "Can return the [whole] table $tableName and capture the data table in a variable " {
        [void](Get-Sql -Session $sessionName -Quiet -Table $tableName  -OutputVariable Table  )
        $table.GetType().fullname                                                       | Should -be "System.Data.DataTable"
    }
    It "Can run abritrary SQL as passed as via the pipe" {
        ($ArbitrarySQL |   Get-SQL -Session $sessionName -Quiet              ).Count    | Should -beGreaterThan 0
    }
    It "Can run abritrary SQL as passed as a parameter" {
        (Get-SQL -Session $sessionName -Quiet $ArbitrarySQL                  ).Count    | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters" {
         (Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT 20 ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters, and values for where condition Piped " {
         (5,10 , 20 |
         Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2 -GT      ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct, -OrderBy and -Where parameters and where condition piped " {
         ("> 5","> 10",">= 20" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1 -Where $fieldname2         ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with -Select, -Distinct and -OrderBy parameters and WHERE... clause piped " {
        ("Where  Points >5 ","Where Points >10","Where Points >= 20" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName -Select $fieldname1 -Distinct -OrderBy $fieldname1                            ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with the WHERE... clause piped but no -Select, -Distinct or -OrderBy " {
        ("Where  Points >5 ","Where Points >10","Where Points >= 20" |
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with multiple fields in -Select and -OrderBy" {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "Race",$fieldname1 -OrderBy $fieldname2,"GridPosition"                ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with -Select holding a date formula" { #SQlite doesn't support "Top"
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName -Select "datetime(date, 'unixepoch') as RaceDate","*" -OrderBy $fieldname1    ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with a different final clause (e.g. 'order by') as a parameter " {
        ( Get-SQL -Session $sessionName -Quiet -Table $tableName "order by $fieldname1 "                                                       ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT query with a different final clause piped " {
        ("order by $fieldname1 " |
          Get-SQL -Session $sessionName -Quiet -Table $tableName                                                                               ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT ... WHERE ... LIKE query with 'naked' syntax and translate * as a wildcard" {
         (sql -Session $sessionName -Select Race,GridPosition,Points -from  "F1Results" -Where Driver -like "Lewis*" -Quiet                    ).Count | Should -beGreaterThan 0
    }
    It "Can run a SELECT Query with -GroupBy and both fieldName & aggreate function in -Select " {
        (  Get-SQL -Session $sessionName -Quiet -Table $tableName  -select $fieldname1,"Count(*) as total" -GroupBy $fieldname1                ).Count | Should -beGreaterThan 0
    }
    It "Can INSERT rows into a table via the pipeline or a parameter and translate dates"      {
        $raceResult = @{Race="Portugese"; Date=([datetime]"2020-10-25"); Driver="Lewis Hamilton"; Team="Mercedes";FinishPosition=1;GridPosition=1;Points=26}
        $raceResult | Get-sql -Session $sessionName -Insert $tableName
        Get-sql -Session $sessionName -Insert $tableName $raceResult
       (Get-sql -Session $sessionName -table  $tableName -where "date" -eq $raceResult.Date.Subtract([datetime]::UnixEpoch).totalseconds -Quiet).Count | Should -be 2
    }
    # Excel Driver doens not support "Can Delete a row from a table"
    It "Can SET new values in a row in a table"   {
        Get-SQL  -Session $sessionName -Table  $tableName -WHERE "Race" -eq "Portugese"`
                            -set "Race" -Values "Portugal" -Confirm:$false
         $new = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Race" -eq "Portugal" -Quiet
         $New.count      | Should -be 2
         $new[0].Points  | Should -be 26
    }

    It "Can Delete rows from a table"   {
        Get-SQL  -Session $sessionName -Table $tableName -WHERE "Race" -eq "Portugal" -Delete -Confirm:$false
        $new = Get-SQL -Session $sessionName -Table  $tableName -WHERE "Race" -eq "Portugal" -Quiet

         $new  | Should -BeNullOrEmpty
    }



    AfterAll {Get-Sql -Session $sessionName -Close }
 }
