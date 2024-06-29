Import-Module ActiveDirectory

$validUserFound = $false

while (-not $validUserFound) {
    $samAccountName = Read-Host -Prompt "Please enter the Username you want to edit"

    try {
        $user = Get-ADUser -Identity $samAccountName -Properties sAMAccountName -Server edw.or.at
        if ($null -ne $user) {
            Write-Host "You chose to edit: $($user.sAMAccountName)" -ForegroundColor Green
            $validUserFound = $true
        } else {
            Write-Host "User does not exist." -ForegroundColor Red
        }
    } catch {
        Write-Host "User does not exist." -ForegroundColor Red
    }
}

$newFirstName = Read-Host -Prompt "Please enter the new First Name"
$newLastName = Read-Host -Prompt "Please enter the new Last Name"

if ([string]::IsNullOrEmpty($newFirstName) -or [string]::IsNullOrEmpty($newLastName)) {
    Write-Host "First Name or Last Name cannot be empty." -ForegroundColor Red
    exit
}

$newFirstName = $newFirstName.Substring(0,1).ToUpper() + $newFirstName.Substring(1).ToLower()
$newLastName = $newLastName.Substring(0,1).ToUpper() + $newLastName.Substring(1).ToLower()

$newLastNamePart = if ($newLastName.Length -ge 6) { $newLastName.Substring(0, 6) } else { $newLastName }
$newFirstNamePart2 = if ($newFirstName.Length -ge 2) { $newFirstName.Substring(0, 2) } else { $newFirstName }
$newFirstNamePart = if ($newFirstName.Length -ge 1) { $newFirstName.Substring(0, 1) } else { $newFirstName }

$newSAMAccountName = "$newLastNamePart$newFirstNamePart2"
$newDisplayName = "$newFirstName $newLastName"

try {
    Set-ADUser -Identity $user -GivenName $newFirstName -Surname $newLastName -SamAccountName $newSAMAccountName -displayName $newDisplayName -Server edw.or.at
    
    $newEmail = "$newFirstNamePart.$newLastName@edw.or.at"
    #$newMailNickname = "$newFirstNamePart.$newLastName"
    
    Set-ADUser -Identity $user -EmailAddress $newEmail<#-OtherAttributes @{mailNickname=$newMailNickname}#> -Server edw.or.at

    Rename-ADObject -Identity $user.DistinguishedName -NewName $newSAMAccountName
    
    Write-Host "`b"
    Write-Host "User details updated successfully." -ForegroundColor Yellow
    Write-Host "`b"
    Write-Host "-----------------------------------------"
    Write-Host "New Username: $newSAMAccountName" -ForegroundColor Green
    Write-Host "New Displayed Name: $newDisplayName" -ForegroundColor Green
    Write-Host "New Email Address: $newEmail" -ForegroundColor Green
    Write-Host "`b"
} catch {
    Write-Host "Failed to update user details." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
