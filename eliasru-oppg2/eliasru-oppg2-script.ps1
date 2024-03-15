# NB!NB!NB!NB! Jeg bruker deler av min forrige oblig. NB!NB!NB!NB!
# Fil "userData.csv" brukes for å opprette brukere i M365
#-----------------------------------------------------------------
# Endringer som har blitt gjort mtp oblig1 kritikk;
# Lagt til input-validering og sikring av gyldig/forventet data
#Kobler til 
Connect-ExchangeOnline -UserPrincipalName EliasBraaaten@y0c4t.onmicrosoft.com -ShowProgress $true
Connect-MgGraph -NoWelcome -Scopes "Sites.FullControl.All"
$userInfo = Import-Csv -Path .\userData.csv

$graphSitesModule = Get-Module -Name Microsoft.Graph.Sites -ListAvailable
$teamsModule = Get-Module -Name MicrosoftTeams -ListAvailable
$Sharepointmodule = Get-Module Microsoft.Online.SharePoint.PowerShell


if ($teamsModule -and $graphSitesModule -and $Sharepointmodule) {

} else {
    Install-Module -Name MicrosoftTeams -Force -Scope CurrentUser
    Install-Module -Name Microsoft.Graph.Sites -Force -Scope CurrentUser
    Install-Module -Name Microsoft.Online.SharePoint.Powershell
}


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



#Funksjon for å opprette ny bruker, tatt fra min forrige oblig.
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
            -UserPrincipalName $user.userPrincipalName `
            -Department $user.department `
            -MailNickName $mailName `
            -PasswordProfile $passwordProfile `
            -AccountEnabled `
            -MobilePhone $user.phonenumber `
            -Mail $user.email `
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

#------------------NYTT INNHOLD TIL OBLIG 2 UNDER HER -----------------------------------#
# Oppgaven inkluderer følgende funksjoner;
# newTeamStatic - Lager teamskanal med userinput.
# addUserTeamStatic - Legger ansatt til Teamsgruppen
# newTeamChannel - Oppretter en ny kanal
# bookRoom - Lar deg lage et nytt rom og lar alle booke det
# newSharePointPage - Lar deg opprette en ny sharepoint side

if ($teamsModule) {

} else {
    Install-Module -Name MicrosoftTeams -Force -Scope CurrentUser
}

#Lager teams statisk ettersom det ikke er logisk å lage automatiske teams. 
function newTeamStatic {

    $readTeamsNavn = Read-Host -Prompt "Skriv navn på Teamskanalen. "
    $readTeamsDesc = Read-Host -Prompt "Gi en kort beskrivelse av teamsgruppen:" 
    do {

    $readTeamsVis  = Read-host -Prompt "Skal den være private/public?"

    } while (-not ($readTeamsVis -notcontains @('private', 'public')))
    
    New-Team -DisplayName $readTeamsNavn -Description $readTeamsDesc -Visibility $readTeamsVis

}
#Her kan vi legge til enkeltbrukere i en teams kanal
function addUserTeamStatic {
    do {
        $readName = Read-Host -Prompt "UserPrincipalName for brukeren du vil legge til: "
        
        $userExists = $userInfo | Where-Object { $_.userPrincipalName -eq $readName }
    
        if (-not $userExists) {
            Write-Host "Bruker $readName eksisterer ikke. Prøv igjen"
        }
    } while (-not $userExists)
    

    do {
        Get-MgGroup | Select-Object Id, DisplayName
        $allGroups = Get-MgGroup | Select-Object Id, DisplayName
        $readGroup = Read-Host "GruppeId som brukeren skal legges til: "
    
        if (-not ($groupExists = $allGroups.Id -contains $readGroup)) { #Ser om gruppen eksisterer
            Write-Host "Gruppe-id $readGroup eksisterer ikke. Prøv igjen"
        } else {
            Add-TeamUser -GroupId $readGroup -User $readName 

            Write-Host -Prompt "Bruker $readName har blitt lagt til i gruppen $readGroupid !"
        }
    
    } while (-not $groupExists)
    
}

function newTeamChannel {


    Get-MgGroup | Select-Object Id, DisplayName
    $allGroups = Get-MgGroup | Select-Object Id, DisplayName
    $readGroupid     = Read-Host -Prompt "Hvilken teamsgruppe vil du lage kanal i?: "
    $readDisplayName = Read-Host -Prompt "Hva skal navnet på kanalen være?: "
    $readDesc        = Read-Host -Prompt "Gi en kort beskrivelse av teamskanalen: " 
    $matchingGroup = $allGroups | Where-Object { $_.Id -eq $readGroupId }

    if (($matchingGroup).Id -contains $readGroupid) { #Ser om gruppen eksisterer
        New-TeamChannel -GroupId $readGroupid -DisplayName $readDisplayName -Description $ReadDesc

        Write-Host -Prompt "Kanal $readDisplayName har blitt lagt til i gruppen $readGroupid !"
    } else {
        Write-Host "Gruppe-id $readGroup eksisterer ikke. Prøv igjen"

    }


}





function bookRoom {
    $roomName = Read-Host "Room navn: "
    $roomAlias = Read-Host "Room alias: "
    New-Mailbox -Name $roomName -Alias $roomAlias  -Room #Lager nytt rom

    Set-MailboxFolderPermission -Identity $roomName -User Default -AccessRights Reviewer #Gir alle mulighet til å booke

}

# Denne funksjon lar deg lage en ny sharepoint side
# Når man lager ny teamsgruppe vil dette bli dannet automatisk, men vi lager en statisk :D

function newSharePointpage {
    Connect-SPOService #krever innlogging ettersom MgGraph ikke har lik tjeneste

    do { #Så lenge karakterer eksisterer
        $siteURL = Read-Host "Hva skal url navnet være?: "
        $finalURL = "https://y0c4t.sharepoint.com/sites/" + $siteURL 
        Write-Host "Checking siteURL: $siteURL"  # Debug line

    } while (-not ($siteURL -match "^[a-zA-Z0-9]+$")) 

    do { #Så lenge karakterer eksisterer
        $siteTitle = Read-Host "Hva skal navnet på siden være?: "
    } while (-not ($siteTitle -match '^[A-Za-z0-9_-]+$'))

        $siteOwner = Read-Host "Hvem skal være eier av siden? (USP): "

    do {    #Så lenge det inneholder kun tall, større eller lik 0 og mindre eller lik 1000 
        $siteStorage = Read-Host "Hvor mye lagring skal tillates siden? (0-1000(MB)): "
    } while ($siteStorage -notmatch "^[0-9]+$" -and ($siteStorage -ge 0 -and $siteStorage -le 1000))

    try {
        New-SPOSite -Url $finalURL -Owner $siteOwner -StorageQuota $siteStorage -Title $siteTitle
    }
    catch {
        Write-Error "Det oppstod en feil under opprettelse av siden: $($_.Exception.Message)"
    }

}

#Get-MgSite -All fungerer ikke, vi må derfor spesifisere siden.
#Derfor ikke mulig å sjekke om userinput er valid.
function searchForSPsite {
    $nameSite = Read-Host "Hva er navnet på siden du vil se? "
    Get-MgSite -Search $nameSite
}


do	{ #Interface
    Write-Output  '
    -----------Menu-----------
    End session             (0)
    Create users dynamically(1)
    Add user to group       (2)
    Opprett ny gruppe       (3)
    Lag nytt teams          (4)
    Legg bruker i teams     (5)
    Lag ny kanal til teams  (6)
    Book ett rom            (7)
    Opprett ny sharepoint   (8)
    Se Sharepoint side      (9)

    --------------------------
    '
    

      # Her kan man endre på infrastrukturen.
    $rspns = Read-Host
    Switch($rspns) {
        0 {break} 
        1 {addUser}
        2 {addToGroupDynamic}
        3 {createGroup}
        4 {newTeamStatic}
        5 {addUserTeamStatic}
        6 {newTeamChannel}
        7 {BookRoom}
        8 {newSharePointpage}
        9 {searchForSPsite}
        default { Write-Host 'Ikke gyldig'}
    }
} while ($rspns -ne "0")
