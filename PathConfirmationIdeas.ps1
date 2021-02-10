{
    ForEach ($path in Get-Content C:\Users\mmorales\TestingPaths.txt) {
        [PSCustomObject]@{
            path = $path
            Exists = Test-Path $path
        }
    }
}

Foreach ($path in Get-Content C:\Users\mmorales\TestingPaths.txt) {
    $path = $path.Split('"')
    [PSCustomObject]@{
        path   = $path[1]
        Exists = Test-Path $path[1]
    }
}

#THIS ONE WORKED

Foreach ($path in Get-Content .\test.txt) {
    [PSCustomObject]@{
         Path   = $path
         Exists = Test-Path $path
         
    }
 }