Install-Package -ProviderName nuget -Name System.Data.SQLite.Core -Scope CurrentUser -Verbose -Force
$dir = Get-ChildItem ~\AppData\Local\PackageManagement\NuGet\Packages\Stub.System.Data.SQLite.Core.NetStandard* -Directory |
            Select-Object -Last 1
Get-ChildItem "$dir\runtimes" -Directory  | ForEach-Object {
    $dest = mkdir $_.name
    Get-ChildItem  $_ -Recurse -Include *.dll | Copy-Item -Destination $dest  -Verbose
    Copy-Item $dir\lib\netstandard2.0\System.Data.SQLite.dll -Destination $dest -Verbose
}

