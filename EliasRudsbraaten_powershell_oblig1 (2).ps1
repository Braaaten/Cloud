#Oppdaterer moduler til nyeste versjon og kobler opp til exchange online
#Find-Module -Name ExchangeOnlineManagement | Install-Module

#Connect-ExchangeOnline
Connect-ExchangeOnline -UserPrincipalName EliasBraaaten@y0c4t.onmicrosoft.com -ShowProgress $true
Connect-MgGraph -NoWelcome
       

function passwordGenerator {
    #Setter passord lengden til 16, men lengde BURDE være mer flukserende for å gjøre det vanskeligere å bruteforce
    param ( 
        [int]$Length = 16 
    )
    #Velger hvilke ord/tall/etc som passordet skal inneholdet
    $CharacterSet = ("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()-_.=}][{}]").ToCharArray()
    #Lager et tilfeldig passord med lengde 16
    $Password = Get-Random -Count $Length -InputObject $CharacterSet
    #Setter det sammen til en hel string.
    return (-join $Password)
}

# Denne funksjonen itterer radene i csv ($userInfo) og lagrer det i variables
# Den vil deretter kalle på "New-MgUser" og tildele csv dataen til enkelt ansatt for hver
# itterering.
function addUser {

    $userInfo = Import-Csv -Path .\userData.csv
     foreach ($user in $userInfo) {
        

        $newPassword = passwordGenerator
        $passwordProfile = @{
            password = $newPassword
            forceChangePasswordNextSignIn = $true
            forceChangePasswordNextSignInWithMfa = $true
        }

        $displayName = ($user.firstName + ' ' +  $user.lastName)
        $mailName = ($user.firstName + $user.lastName)
        New-MgUser -DisplayName $displayName `
            -GivenName $user.firstName `
            -SurName $user.lastName `
            -UserPrincipalName $user.$userPrincipalName `
            -Department $user.department `
            -MailNickName $mailName `
            -PasswordProfile $passwordProfile `
            -AccountEnabled `
            -MobilePhone $user.phonenumber `
            -Mail $user.email `
    }
}

#Lager enkeltbruker
function addUserStatic {

    $firstName = Read-Host -Prompt "Fornavn:"
    $lastName = Read-Host -Prompt "Etternavn"
    $department = Read-Host -Prompt "Hvilken avdeling jobber den ansatte i? "
    $phoneNumber = Read-Host -Prompt "Mobilnummer: "
    $email = Read-Host -Prompt "Personlig email: "

    $displayNamestat = ($firstName + ' ' +  $lastName)
    $userPrincipalNamestat= ($firstName + $lastName + "@y0c4t.onmicrosoft.com")
    $mailNamestat = ($firstName + $lastName)

    $newPassword = passwordGenerator
    $passwordProfile = @{
        password = $newPassword
        forceChangePasswordNextSignIn = $true
        forceChangePasswordNextSignInWithMfa = $true
    }

try {
    New-MgUser -DisplayName $displayNamestat `
    -GivenName $firstName `
    -SurName $lastName `
    -UserPrincipalName $userPrincipalNamestat `
    -Department $department `
    -MailNickName $mailNamestat `
    -PasswordProfile $passwordProfile `
    -AccountEnabled `
    -MobilePhone $phoneNumber `
    -Mail $email `
    
    return
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
}

#Legger alle brukere i hver sin gruppe basert på avdeling
function addToGroupDynamic {
    # Henter alle ansatte
    $allUsers = Get-MgUser -All 
    
    foreach ($user in $allUsers) { #For hver ansatt
        #Henter ansatt sitt "department"
        try {
        $userDepartment = (Get-MgUser -UserId $user.Id -Property Department).Department 
        } catch {
            Write-Output "Bruker har ikke blitt tildelt avdeling." 
            return
        }

        #Ser om gruppe navn og department navn er =
        $matchFound = Get-MgGroup -All | Where-Object { $_.DisplayName -eq $userDepartment } 

        #Sjekker om gruppen har department
        if ($matchFound.Count -gt 0) {
            foreach ($group in $matchFound)  { #for hver match vi har
                $groupId = $group.Id #lagrer gruppe ID
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $user.Id #Plasserer ansatt i gruppen.
                Write-Output "Added user $($user.DisplayName) to group $($group.DisplayName)" 
            }
        } else {
            Write-Output "Ingen gruppe funnet for avdeling: $userDepartment for bruker: $($user.DisplayName)"
        }
        
    }
}

function createGroup {
    $displayName = Read-Host -Prompt "Velg gruppe navn"
    $description = Read-Host -Prompt "Gi en beskrivelse av gruppens hensikt"
    $mailEnabled = Read-Host -Prompt "Skal gruppen ha aktivert epost? (true/false)"
    $securityEnabledd = Read-Host -Prompt "Er gruppen security-enabled? (true/false)"

    try {
        if ($mailEnabled -eq "false" -and $securityEnabledd -eq "true" ) { #danner securit group
            $newGroup = New-MgGroup -DisplayName $displayName -Description $description -MailEnabled:$False  -MailNickName ($displayName + ".Mail") -SecurityEnabled
        } else { #danner mail group
            New-MgGroup -DisplayName "hei" -Description "kai" -MailEnabled  -MailNickName ("kai" + ".Mail") -SecurityEnabled:$False -GroupTypes "Unified" 
        }

        Write-Host "Ny gruppe lagt til:"
        Write-Host "Gruppe navn: $($newGroup.DisplayName)"
        Write-Host "Beskrivelse: $($newGroup.Description)"
        Write-Host "Mail aktivert: $($newGroup.MailEnabled)"
        Write-Host "Security aktivert: $($newGroup.SecurityEnabled)"
    }
    
    catch {
        Write-Host "Dannelse av ny gruppe feilet: $($_.Exception.Message)"
        return 
    }
   
}

#Brukere og administratorer har forskjellige krav for passord. 
function setPasswordPolicy {
    $userInfo = Import-Csv -Path .\userData.csv

    foreach ($user in $userInfo) {
        if ($user.department  -ne "IT-administrator") {
        Update-MgUser -UserId $user.userPrincipalName -PasswordPolicies "DisableStrongPassword"
        }
        else {
            Write-Output "Cannot change password policy for IT-administrator: $($user.DisplayName)"
        }
    }
}

# Del user for nå sletter ALLE brukere i AD, forsimplet for testing.
function delUsers {
    $allUsers = Get-MgUser -All

    foreach ($user in $allUsers) {
        Remove-MgUser -UserId $user.Id -Confirm:$false
        Write-Host "Deleted user: $($user.UserPrincipalName)"
    }
}

#Dersom man vil at alt skal gjøres i "one go"
function completeSetUp {
    addUser                 #Lage nye brukere
    addToGroupDynamic       #Legger brukere i gruppe
    setPasswordPolicy       #setter ny passord policy til ansatte med mindre rettigheter
}
# Funksjon for å se noe data fra alle brukere
function seeUserInfo {
    Get-MgUser -All -Property DisplayName, Department,Username | Select-Object DisplayName, Department,Username | Format-Table -AutoSize
}

#Update-Module Microsoft.Graph.Identity.DirectoryManagement
#Update-Module Microsoft.Graph.Users.Actions
#Gir ansatt statisk ny license, se https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
function giveLicense {
    
    $groupChoosen = Read-Host "Which group do you want to give a lincense? (ID)"

    Get-MgGroup -All -Property Id,Displayname

    $EmsSku = Read-Host "Which license do you want to give? (f.eks: c42b9cae-ea4f-4ab7-9717-81576235ccac)"    


    $value = @{
        AddLicenses = @(
        @{
        skuId = $EmsSku
        }
            )
        RemoveLicenses = @(

        )
    }

    Set-MgGroupLicense -GroupId $groupChoosen -BodyParameter $value #setter licence
    (Get-MgGroup -All -Property DisplayName,"AssignedLicenses" | Select-Object -ExpandProperty AssignedLicenses).SkuId


}
    do	{ #Interface
        Write-Output  '
        -----------Menu-----------
        End session             (0)
        Create users dynamically(1)
        Create user statically  (2) 
        Delete users            (3)
        User infor              (4) 
        Add user to group       (5)
        Change password policy  (6)
        Opprett ny gruppe       (7)
        Tildel license          (8)
        --------------------------
        '
        

          # Her kan man endre på infrastrukturen.
        $rspns = Read-Host
        Switch($rspns) {
            0 {break}
            1 {addUserDynamic}
            2 {addUserStatic}
            3 {delUsers}
            4 {seeUserInfo}
            5 {addToGroupDynamic}
            6 {setPasswordPolicy}
            7 {createGroup}
            8 {giveLicense}

            default { Write-Host 'Ikke gyldig'}
        }
    } while ($rspns -ne "0")
    