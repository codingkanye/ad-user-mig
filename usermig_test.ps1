Import-Module ActiveDirectory

$validUserFound = $false

Write-Host "     ! Question !" -ForegroundColor Yellow
$edworkk = Read-Host -Prompt "Do you want to edit for EDW or @KK?"

if ($edworkk -eq "edw") {

    while (-not $validUserFound) {
        Write-Host "You chose to edit for EDW!" -ForegroundColor Green
        $samAccountName = Read-Host -Prompt "Please enter the Username you want to edit"

        try {
            $user = Get-ADUser -Identity $samAccountName -Properties sAMAccountName, proxyAddresses, company, department, description -Server edw.or.at
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
        Set-ADUser -Identity $user -GivenName $newFirstName -Surname $newLastName -SamAccountName $newSAMAccountName -DisplayName $newDisplayName -Server edw.or.at
        
        $newEmail = "$newFirstNamePart.$newLastName@edw.or.at"
        $newMailNickname = "$newFirstNamePart.$newLastName"
        $newUPN = "$newMailNickname@edw.or.at"

        Set-ADUser -Identity $user -EmailAddress $newEmail -Server edw.or.at
        Set-ADUser -Identity $user -Replace @{mailNickname=$newMailNickname; UserPrincipalName=$newUPN} -Server edw.or.at

        $existingProxyAddresses = @()
        if ($user.proxyAddresses) {
            $existingProxyAddresses = $user.proxyAddresses
        }

        $updatedProxyAddresses = @()
        foreach ($address in $existingProxyAddresses) {
            if ($address -like "SMTP:*") {
                $updatedProxyAddresses += $address.ToLower().Replace("smtp:", "smtp:")
            } else {
                $updatedProxyAddresses += $address
            }
        }
        $updatedProxyAddresses += "SMTP:$newUPN"

        Set-ADUser -Identity $user -Replace @{proxyAddresses=$updatedProxyAddresses} -Server edw.or.at

        Rename-ADObject -Identity $user.DistinguishedName -NewName $newSAMAccountName
        
        Write-Host "`b"
        Write-Host "-----------------------------------------"
        Write-Host "`b"
        Write-Host "User details updated successfully." -ForegroundColor Yellow
        Write-Host "`b"
        Write-Host "New Username: $newSAMAccountName" -ForegroundColor Green
        Write-Host "`b"
        Write-Host "New Displayed Name: $newDisplayName" -ForegroundColor Green
        Write-Host "`b"
        Write-Host "New Mail Nickname: $newMailNickname" -ForegroundColor Green
        Write-Host "`b"
        Write-Host "New Email Address: $newEmail" -ForegroundColor Green
        Write-Host "`b"
        Write-Host "New UPN: $newUPN" -ForegroundColor Green
        Write-Host "`b"
        Write-Host "Updated Proxy Addresses: $($updatedProxyAddresses -join ', ')" -ForegroundColor Green
        Write-Host "-----------------------------------------"
        Write-Host "`b"

        $referenceSamAccountName = Read-Host -Prompt "Please enter the Username of the reference user to copy attributes from"

        try {
            $referenceUser = Get-ADUser -Identity $referenceSamAccountName -Properties company, department, description, extensionAttribute15  -Server edw.or.at
            if ($null -ne $referenceUser) {
                $userAttributes = @{
                    company = $referenceUser.company
                    department = $referenceUser.department
                    description = $referenceUser.description
                    extensionAttribute15 = $referenceUser.extensionAttribute15
                }
                
                Set-ADUser -Identity $newSAMAccountName -Replace $userAttributes -Server edw.or.at
        
                Write-Host "`b"
                Write-Host "-----------------------------------------"
                Write-Host "`b"
                Write-Host "Attributes copied from reference user" -ForegroundColor Yellow
                Write-Host "`b"
                Write-Host "Company: $($referenceUser.company)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Department: $($referenceUser.department)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Description: $($referenceUser.description)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Responsible Cost Center: $($referenceUser.extensionAttribute15)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "-----------------------------------------"
                Write-Host "`b"
            } else {
                Write-Host "Reference user does not exist." -ForegroundColor Red
            }
        } catch {
            Write-Host "Failed to copy attributes from reference user." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
        

    } catch {
        Write-Host "Failed to update user details." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}


    <# Write-Host "     ! Question !" -ForegroundColor Yellow
    $changeCompany = Read-Host -Prompt "Do you want to change the company as well? (yes/no)"

    $newCompany = $user.company
    if ($changeCompany -eq "yes") {
        $newCompany = Read-Host -Prompt "Please enter the new company name"
        Set-ADUser -Identity $newSAMAccountName -Company $newCompany -Server edw.or.at
        Write-Host "`b"
        Write-Host "-----------------------------------------"
        Write-Host "`b"
        Write-Host "New Company: $newCompany" -ForegroundColor Green
        Write-Host "`b"
    } else {
        Write-Host "`b"
        Write-Host "-----------------------------------------"
        Write-Host "You did not change the Company! The old Value stays." -ForegroundColor Red
    }

    Write-Host "-----------------------------------------"
    Write-Host "     ! Question !" -ForegroundColor Yellow
    $changeDepartment = Read-Host -Prompt "Do you want to change the Department as well? (yes/no)"

    $newDepartment = $user.department
    if ($changeDepartment -eq "yes") {
        $newDepartment = Read-Host -Prompt "Please enter the new Department name"
        Set-ADUser -Identity $newSAMAccountName -Department $newDepartment -Server edw.or.at
        Write-Host "`b"
        Write-Host "-----------------------------------------"
        Write-Host "`b"
        Write-Host "New Department: $newDepartment" -ForegroundColor Green
        Write-Host "`b"
    } else {
        Write-Host "`b"
        Write-Host "-----------------------------------------"
        Write-Host "You did not change the Department! The old Value stays." -ForegroundColor Red
    }

    $combinedCompanyDepartment = "$newCompany / $newDepartment"
    Set-ADUser -Identity $newSAMAccountName -Description $combinedCompanyDepartment -Server edw.or.at
    Write-Host "-----------------------------------------"
} catch {
    Write-Host "Failed to update user details." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
   }

} else {
   Write-Host "You chose to edit for @KK!" -ForegroundColor Green
}

 #>

<# new lin #>

if ($edworkk -eq "kk") {

   while (-not $validUserFound) {
       $samAccountName = Read-Host -Prompt "Please enter the Username you want to edit"
   
       try {
           $user = Get-ADUser -Identity $samAccountName -Properties sAMAccountName, proxyAddresses, company, department, description -Server edw.or.at
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
   <#$newFirstNamePart = if ($newFirstName.Length -ge 1) { $newFirstName.Substring(0, 1) } else { $newFirstName }#>
   
   $newSAMAccountName = "p9"+"$newLastNamePart$newFirstNamePart2"
   $newDisplayName = "$newFirstName $newLastName"
   
   try {
       Set-ADUser -Identity $user -GivenName $newFirstName -Surname $newLastName -SamAccountName $newSAMAccountName -DisplayName $newDisplayName -Server edw.or.at
       
       $newEmail = "$newFirstName.$newLastName@katholischekirche.at"
       $newMailNickname = "$newFirstName.$newLastName"
       $newUPN = "$newMailNickname@katholischekirche.at"
   
       Set-ADUser -Identity $user -EmailAddress $newEmail -Server edw.or.at
       Set-ADUser -Identity $user -Replace @{mailNickname=$newMailNickname; UserPrincipalName=$newUPN} -Server edw.or.at
   
       $existingProxyAddresses = @()
       if ($user.proxyAddresses) {
           $existingProxyAddresses = $user.proxyAddresses
       }
   
       $updatedProxyAddresses = @()
       foreach ($address in $existingProxyAddresses) {
           if ($address -like "SMTP:*") {
               $updatedProxyAddresses += $address.ToLower().Replace("smtp:", "smtp:")
           } else {
               $updatedProxyAddresses += $address
           }
       }
       $updatedProxyAddresses += "SMTP:$newUPN"
   
       Set-ADUser -Identity $user -Replace @{proxyAddresses=$updatedProxyAddresses} -Server edw.or.at
   
       Rename-ADObject -Identity $user.DistinguishedName -NewName $newSAMAccountName
       
       Write-Host "`b"
       Write-Host "-----------------------------------------"
       Write-Host "`b"
       Write-Host "User details updated successfully." -ForegroundColor Yellow
       Write-Host "`b"
       Write-Host "-----------------------------------------"
       Write-Host "New Username: $newSAMAccountName" -ForegroundColor Green
       Write-Host "`b"
       Write-Host "New Displayed Name: $newDisplayName" -ForegroundColor Green
       Write-Host "`b"
       Write-Host "New Mail Nickname: $newMailNickname" -ForegroundColor Green
       Write-Host "`b"
       Write-Host "New Email Address: $newEmail" -ForegroundColor Green
       Write-Host "`b"
       Write-Host "New UPN: $newUPN" -ForegroundColor Green
       Write-Host "`b"
       Write-Host "Updated Proxy Addresses: $($updatedProxyAddresses -join ', ')" -ForegroundColor Green
       Write-Host "-----------------------------------------"
       Write-Host "`b"

        # Ask for a reference user to copy attributes from
        $referenceSamAccountName = Read-Host -Prompt "Please enter the Username of the reference user to copy attributes from"

        try {
            $referenceUser = Get-ADUser -Identity $referenceSamAccountName -Properties company, department, description, extensionAttribute15 -Server edw.or.at
            if ($null -ne $referenceUser) {
                Set-ADUser -Identity $user -Replace @{
                    company = $referenceUser.company
                    department = $referenceUser.department
                    description = $referenceUser.description
                    extensionAttribute15 = $referenceUser.extensionAttribute15
                } -Server edw.or.at

                Write-Host "`b"
                Write-Host "-----------------------------------------"
                Write-Host "`b"
                Write-Host "Attributes copied from reference user:" -ForegroundColor Yellow
                Write-Host "`b"
                Write-Host "Company: $($referenceUser.company)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Department: $($referenceUser.department)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Description: $($referenceUser.description)" -ForegroundColor Green
                Write-Host "`b"
                Write-Host "Extension Attribute 15: $($referenceUser.extensionAttribute15)" -ForegroundColor Green
                Write-Host "-----------------------------------------"
                Write-Host "`b"
            } else {
                Write-Host "Reference user does not exist." -ForegroundColor Red
            }

        } catch {
            Write-Host "Failed to copy attributes from reference user." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }

   }
       catch {
        Write-Host "Failed to update user details." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
       }
   }
   
    <#   Write-Host "     ! Question !" -ForegroundColor Yellow
       $changeCompany = Read-Host -Prompt "Do you want to change the company as well? (yes/no)"
   
       $newCompany = $user.company
       if ($changeCompany -eq "yes") {
           $newCompany = Read-Host -Prompt "Please enter the new company name"
           Set-ADUser -Identity $newSAMAccountName -Company $newCompany -Server edw.or.at
           Write-Host "`b"
           Write-Host "-----------------------------------------"
           Write-Host "`b"
           Write-Host "New Company: $newCompany" -ForegroundColor Green
           Write-Host "`b"
       } else {
           Write-Host "`b"
           Write-Host "-----------------------------------------"
           Write-Host "You did not change the Company! The old Value stays." -ForegroundColor Red
       }
   
       Write-Host "-----------------------------------------"
       Write-Host "     ! Question !" -ForegroundColor Yellow
       $changeDepartment = Read-Host -Prompt "Do you want to change the Department as well? (yes/no)"
   
       $newDepartment = $user.department
       if ($changeDepartment -eq "yes") {
           $newDepartment = Read-Host -Prompt "Please enter the new Department name"
           Set-ADUser -Identity $newSAMAccountName -Department $newDepartment -Server edw.or.at
           Write-Host "`b"
           Write-Host "-----------------------------------------"
           Write-Host "`b"
           Write-Host "New Department: $newDepartment" -ForegroundColor Green
           Write-Host "`b"
       } else {
           Write-Host "`b"
           Write-Host "-----------------------------------------"
           Write-Host "You did not change the Department! The old Value stays." -ForegroundColor Red
       }
   
       $combinedCompanyDepartment = "$newCompany / $newDepartment"
       Set-ADUser -Identity $newSAMAccountName -Description $combinedCompanyDepartment -Server edw.or.at
       Write-Host "-----------------------------------------"
   } catch {
       Write-Host "Failed to update user details." -ForegroundColor Red
       Write-Host $_.Exception.Message -ForegroundColor Red
      }
   
   } else {
      Write-Host "No valid command!" -ForegroundColor Red
}   
      #>
