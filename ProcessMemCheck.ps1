$processToMeasure = 'McShield'
$results = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]
Get-ADObject -LDAPFilter "(objectClass=computer)" | select-object -ExpandProperty name | ForEach-Object {
    $currComp = $_
    #$currComp
    Try { 
        Test-Connection $_ -Count 1 -ErrorAction Stop > $null
        Try {
            $totalSystemMemory = (Get-WmiObject -ComputerName $currComp -ClassName 'Cim_PhysicalMemory' | 
                                  Measure-Object -Property Capacity -Sum).Sum
            $result = get-process -name $processToMeasure | select-object -ExpandProperty workingset
            $resultMB = [math]::Round(($result/1MB), 2)
            $resultPercentSystemTotal = ($result/$totalSystemMemory) * 100
            $results.Add([PSCustomObject]@{'Hostname'=$currComp ; "MemUsage" = $resultMB ; 'PercentSystemTotal' = $resultPercentSystemTotal})
        }
        Catch { $errors.Add("Error reaching ($currComp): $_.Exception.Message") }
    }
    Catch{ 
         $errors.Add("$($currComp): $_.Exception.Message")
    }
}
Write-Host "`n$processToMeasure Memory Usage Results"
$results | Sort-Object $_.PercentSystemTotal | Select-Object Hostname, @{n='Memory Usage (MB)' ; e= { $_.MemUsage }},
                                                                       @{n='Percentage of Total System Memory' ; e= { $_.PercentSystemTotal }} -OutVariable export
$timeStamp = $(Get-Date -Format MMddyyyy_HHmmss)
$Export | Export-Csv -Path ".\$processToMeasure Memory Usage Check Output-$timestamp.csv"

Write-Host "`nErrors:`n"
$errors

$errors | Export-Csv -Path ".\$processToMeasure Memory Usage Check Error Log-$timestamp.csv"

# Resources
# https://stackoverflow.com/questions/1984186/what-is-private-bytes-virtual-bytes-working-set
# https://stackoverflow.com/questions/14726143/get-memory-usage-as-reported-by-windows-8-task-manager
# https://stackoverflow.com/questions/26552223/get-process-with-total-memory-usage
# https://stackoverflow.com/questions/7954781/whats-the-difference-between-working-set-and-commit-size#:~:text=1%20Answer&text=From%20here%2C%20the%20working%20set%20is%3A&text=So%20you%20can%20think%20of,other%20than%20the%20page%20file).