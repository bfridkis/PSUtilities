Import-Module ActiveDirectory

$processToMeasure = $utilizationThreshold = $results = $cleared = $errors = $null

([string]$args).split('-') | ForEach-Object { 
                                if ($_.Split(' ')[0] -eq "ProcessToMeasure") { $processToMeasure = $_.Split(' ')[1] }
                                if ($_.Split(' ')[0] -eq "UtilizationThreshold") { $utilizationThreshold = $_.Split(' ')[1] }
                              }

if (!$processToMeasure -and !$utilizationThreshold) { $runAsScript = $true } else { $runAsScript = $false }
if (!$processToMeasure) { $processToMeasure = Read-Host -prompt "`nName of Process to Measure " }
if (!$utilizationThreshold) { $utilizationThreshold = Read-Host -prompt "Utilization Threshold (% of total system memory to flag via email - leave blank for none) " }

$flaggedNodesOutputFilePath = "E:\ProcessMemCheck\$($processToMeasure) MemCheck Exceedance\$($processToMeasure) Flagged Nodes.txt"

Get-ChildItem -Path "E:\ProcessMemCheck\$($processToMeasure) MemCheck*\" -Recurse -Force -ErrorAction Silent | 
Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt (Get-Date).AddDays(-14) -and $_.Name -ne "$($processToMeasure) Flagged Nodes.txt"} | 
Remove-Item -Force

Write-Host "`nRunning, please wait..."

$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

$results = New-Object System.Collections.Generic.List[System.Object]
$cleared = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]

if (!(Test-Path -Path "E:\ProcessMemCheck\$($processToMeasure) MemCheck Exceedance")) {New-Item "E:\ProcessMemCheck\$($processToMeasure) MemCheck Exceedance" -ItemType directory -force >$null}
if (!(Test-Path -Path "E:\ProcessMemCheck\$($processToMeasure) MemCheck Output")) {New-Item "E:\ProcessMemCheck\$($processToMeasure) MemCheck Output" -ItemType directory -force >$null}
if (!(Test-Path -Path $flaggedNodesOutputFilePath)) {
    $flaggedNodesOutputFile = New-Item $flaggedNodesOutputFilePath -ItemType file -force
    $flaggedNodesOutputFile.attributes = "Hidden"
}

#Get-ADObject -LDAPFilter "(objectClass=computer)" | where-object { $_.name -like "*DEV*" } | select-object -expandproperty name -outvariable comps > $null
Get-ADObject -LDAPFilter "(objectClass=computer)" | select-object -expandproperty name -outvariable comps > $null
$comps | ForEach-Object {
    $currComp = $_
    Try { 
        Test-Connection $_ -Count 1 -ErrorAction Stop > $null
        Try {
            $totalSystemMemory = (Get-WmiObject -ComputerName $currComp -ClassName 'Cim_PhysicalMemory' -ErrorAction Stop | 
                                  Measure-Object -Property Capacity -Sum).Sum
            Try {
                $result = get-process -name $processToMeasure -ComputerName $currComp -ErrorAction Stop | 
                          select-object -ExpandProperty workingset
                $resultMB = [math]::Round(($result/1MB), 2)
                $resultPercentSystemTotal = [math]::Round(($result/$totalSystemMemory) * 100, 2)
                $results.Add([PSCustomObject]@{'Hostname'=$currComp ; 
                                               'MemUsage' = $resultMB ; 
                                               'PercentSystemTotal' = $resultPercentSystemTotal ;
                                               'TotalSystemMemory' = $($totalSystemMemory / 1GB)
                                               })
                if($utilizationThreshold) {
                    if(!(Select-String -Path $flaggedNodesOutputFilePath -Pattern $currComp) -and $resultPercentSystemTotal -gt $utilizationThreshold) { 
                        $_ | Add-Content -Path $flaggedNodesOutputFilePath 
                    }
                    elseif ((Select-String -Path $flaggedNodesOutputFilePath -Pattern "$currComp -emailed") -and $resultPercentSystemTotal -lt $utilizationThreshold) { 
                        $flaggedNodesOutputFileUpdate = Get-Content -Path $flaggedNodesOutputFilePath | Where-Object {$_ -ne "$currComp -emailed"}
                        $flaggedNodesOutputFileUpdate | Set-Content $flaggedNodesOutputFilePath

                        $cleared.Add([PSCustomObject]@{'Hostname'=$currComp ; 
                                                       'MemUsage' = $resultMB ; 
                                                       'PercentSystemTotal' = $resultPercentSystemTotal ;
                                                       'TotalSystemMemory' = $($totalSystemMemory / 1GB)
                                                      })
                    }
                }
            }
            Catch { $errors.Add([PSCustomObject]@{'Hostname'=$currComp ; 'Exception' = $_.Exception.Message}) }
        }
        Catch { $errors.Add([PSCustomObject]@{'Hostname'=$currComp ; 'Exception' = $_.Exception.Message}) }
    }
    Catch{ 
         $errors.Add([PSCustomObject]@{'Hostname'=$currComp ; 'Exception' = $_.Exception.Message})
    }
}

$elapsedTime = $stopWatch.Elapsed.TotalSeconds

$timeStamp = Get-Date -Format MMddyyyy_HHmmss

$results | Sort-Object PercentSystemTotal -Descending | Select-Object Hostname, @{n='Memory Usage (MB)' ; e= { "{0:N2}" -f $_.MemUsage }},
                                                                                @{n='Percentage of Total System Memory' ; e= { "{0:N2}" -f $_.PercentSystemTotal }},
                                                                                @{n='Total System Memory (GB)' ; e= {$_.TotalSystemMemory}} -OutVariable Export >$null

$outputFile = "E:\ProcessMemCheck\$($processToMeasure) MemCheck Output\$($processToMeasure) MemCheck Output-$timeStamp.csv"
$outputString = "$processToMeasure Memory Usage Results"
Add-Content -Path $outputFile -Value $outputString
$Export | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $outputFile

$outputString = "`r`n** Errors **"
Add-Content -Path $outputFile -Value $outputString
$errors | Select-Object @{ n = 'Hostname' ; e = {$_.Hostname}},
                        @{ n = 'Exceptions Generated' ; e = {$_.Exception}} |
          Sort-Object Hostname | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $outputFile

Write-Host "`n$processToMeasure Memory Usage Results"
$export | Format-Table Hostname, @{n='Memory Usage (MB)' ; e= { "{0:N2}" -f $_."Memory Usage (MB)" } ; a="right"},
                                 @{n='Percentage of Total System Memory' ; e= { "{0:N2}" -f $_."Percentage of Total System Memory" } ; a="right"},
                                 @{n='Total System Memory (GB)' ; e= {$_."Total System Memory (GB)"}}
Write-Host "Errors:"
$errors | Sort-Object Hostname | Format-Table

write-output "`nExecution Complete. $(if ($elapsedTime -gt 60) { "$([math]::Round($elapsedTime/60, 2)) minutes" } else { "$([math]::Round($elapsedTime, 2)) seconds" })."

if($utilizationThreshold) {

    $recipients = @("ben.j.fridkis@p66.com", "john.gusewelle@p66.com")
    $Cc = @("paul.l.bird@p66.com")

    $export.Where({$_."Percentage of Total System Memory" -lt $utilizationThreshold}, "Until") | Select-Object -OutVariable utilizationThresholdExceeded > $null

    if($utilizationThresholdExceeded.Count -gt 0) {
        $ExceedanceOutputFile = "E:\ProcessMemCheck\$($processToMeasure) MemCheck Exceedance\$($processToMeasure) MemCheck Exceedance Report-$timeStamp.csv"
        $outputString = "$processToMeasure Memory Usage Exceedance Report (> $utilizationThreshold% Total System Memory)"
        Add-Content -Path $ExceedanceOutputFile -Value $outputString
        $utilizationThresholdExceeded | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $ExceedanceOutputFile

        $unprocessedFlagCount = (Get-Content $flaggedNodesOutputFilePath | Where-Object { ($_ -notlike "*-emailed" -and $_ -ne "") } | Measure-Object).Count
        if($unprocessedFlagCount -gt 0) {
            Send-MailMessage -From PCNSMS04@wmgpcn.local -To $recipients -Cc $Cc -Subject "$processToMeasure Memory Utilization Threshold Exceedance Report-$timestamp"  `
            #Send-MailMessage -From PCNSMS04@wmgpcn.local -To ben.j.fridkis@p66.com -Subject "$processToMeasure Memory Utilization Threshold Exceedance Report-$timestamp"  `
                             -Attachments $ExceedanceOutputFile -SmtpServer 164.123.219.98
            $flaggedNodesOutputFileUpdate = $null
            Get-Content $flaggedNodesOutputFilePath | ForEach-Object { if($_ -like "*-emailed") {$flaggedNodesOutputFileUpdate += "$_`r`n" } elseif($_ -ne "") {$flaggedNodesOutputFileUpdate += "$_ -emailed`r`n" } }
            $flaggedNodesOutputFileUpdate | Set-Content -Path $flaggedNodesOutputFilePath -force
        }
    }
    if($cleared.Count -gt 0) {
        $outputString = "$processToMeasure Memory Usage Exceedances CLEARED (now < $utilizationThreshold% Total System Memory)"
        $tempFileForClearedResults = New-Item "E:\ProcessMemCheck\$($processToMeasure) MemCheck Exceedance\$($processToMeasure) MemCheck Exceedance CLEARED-$timeStamp.csv" -ItemType file -force
        Add-Content -Path $tempFileForClearedResults -Value $outputString
        $cleared | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $tempFileForClearedResults
        Send-MailMessage -From PCNSMS04@wmgpcn.local -To $recipients -Cc $Cc -Subject "$processToMeasure Memory Utilization Threshold Exceedances CLEARED-$timestamp"  `
        #Send-MailMessage -From PCNSMS04@wmgpcn.local -To ben.j.fridkis@p66.com -Subject "$processToMeasure Memory Utilization Threshold Exceedances CLEARED-$timestamp"  `
                         -Attachments $tempFileForClearedResults -SmtpServer 164.123.219.98
        remove-item $tempFileForClearedResults
    }
}

if($runAsScript) {
    Write-Host "`n"
    Pause
}

# Resources
# https://stackoverflow.com/questions/1984186/what-is-private-bytes-virtual-bytes-working-set
# https://stackoverflow.com/questions/14726143/get-memory-usage-as-reported-by-windows-8-task-manager
# https://stackoverflow.com/questions/26552223/get-process-with-total-memory-usage
# https://stackoverflow.com/questions/7954781/whats-the-difference-between-working-set-and-commit-size#:~:text=1%20Answer&text=From%20here%2C%20the%20working%20set%20is%3A&text=So%20you%20can%20think%20of,other%20than%20the%20page%20file).
# https://stackoverflow.com/questions/4857792/powershell-reference-a-property-that-contains-a-space
# https://stackoverflow.com/questions/48425562/powershell-add-content-should-create-path-but-throws-exception-could-not-find-a
# https://stackoverflow.com/questions/41871147/find-specific-string-in-textfile-powershell/41871326
# https://stackoverflow.com/questions/226596/powershell-array-initialization
# https://stackoverflow.com/questions/10241816/powershell-send-mailmessage-email-to-multiple-recipients
# https://social.technet.microsoft.com/Forums/en-US/33665b03-d383-41ed-a836-fd83c217b3f1/making-files-hidden-with-powershell?forum=ITCG
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.1