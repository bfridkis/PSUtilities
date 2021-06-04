$user = "admbfridkis"
$nodesWithUserLoggedOn = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]
Get-ADObject -LDAPFilter "(objectClass=computer)" | select-object -ExpandProperty name | ForEach-Object {
    $currComp = $_
    #$currComp
    Try { 
        Test-Connection $_ -Count 1 -ErrorAction Stop > $null
        if($currComp -notlike "*PCNVS*") { 
            Try {
                $results = $(qwinsta "/server:$currComp")
                if ($results -like "*$user*") { $nodesWithUserLoggedOn.Add($currComp) }
            }
            Catch { $errors.Add("Error reaching ($currComp): $_.Exception.Message") }
        }
    }
    Catch{ 
         $errors.Add("$($currComp): $_.Exception.Message")
    }
}
Write-Host "`nNodes with User Logged On:`n$nodesWithUserLoggedOn"

Write-Host "`nErrors:`n"
$errors