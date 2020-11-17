filter Test-SQlite {
<#
.synopsis
    By piping files through this command those which NOT SQlite database will be discarded
.example
    dir '~\AppData\Local\google\chrome\User Data\Default\' -File | Test-SQlite
    Gives a directory listing of SQlite files in the user data for Google Chrome
#>
    [char[]]$c = " " *16
    $o = $_.OpenText()
    [void]$o.read($c,0,16)
    $o.Close()
    if ([string]::new($c) -match "SQLite") {$_}
}