
#THIS ONE WORKED

write-host -nonewline "Continue? (Y/N) "
$response = read-host
if ( $response -ne "Y" ) { exit }

Foreach ($path in Get-Content C:\Users\mmorales\TestingPaths.txt) {
    [PSCustomObject]@{
         Path   = $path
         Exists = Test-Path $path
         
        }
    }
