Import-Module ActiveDirectory

# Check arg number
if ($args.Count -ne 1){
    Write-Host "Error: incorrect number of arguments" -ForegroundColor Red -BackgroundColor Black
    Write-Host "Usage: .\gen-ad-users.ps1 <AD users .csv file>" -ForegroundColor Red -BackgroundColor Black
    exit
}

# Variables
$expectedHeaders = @("Username","Name","Firstname","Lastname","Group","TempPassword")
$actualHeaders = $null
$DN = "OU=Marketing,DC=domain100,DC=com"
$csvFilePath = $args[0]
$domain = $null
$lineNumber = 2

# Check CSV file exists
if (!(Test-Path -Path $csvFilePath)){
    Write-Host "Error: Input file '$csvFilePath' does not exist in current directory" -ForegroundColor Red -BackgroundColor Black
    exit
}

# Check connection to Domain
try{
    $domain = "@"+[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
}catch{
    Write-Host "Error: Couldn't establish connection to the domain" -ForegroundColor Red -BackgroundColor Black
    exit
}

# Read in CSV file
$Users = Import-Csv -Delimiter "," -Path $csvFilePath

# Error checking for headers performed in surrounding python script

# Main
foreach ($User in $Users) {
    try{
        $SAM = $User.Username
        $Displayname = $User.Name
        $Firstname = $User.Firstname
        $Lastname = $User.Lastname
        $assignedSecurityGroup = $User.Group
        $UPN = $User.Username + $domain
        $Password = (ConvertTo-SecureString $User.TempPassword -AsPlainText -Force)
        $EmailAddress = $User.Username + "@" + "mtu.edu"
        New-ADUser -Name "$Displayname" -DisplayName "$Displayname" -SamAccountName "$SAM" -UserPrincipalName "$UPN" -GivenName "$Firstname" -Surname "$Lastname" -EmailAddress "$EmailAddress" -AccountPassword $Password  -Enabled $true -Path "$DN" -ChangePasswordAtLogon $true -PasswordNeverExpires $false
        Add-ADGroupMember -Identity $assignedSecurityGroup -Members $SAM
        Write-Host "Created user: $UPN" -ForegroundColor Green
    }
# Handle Duplicate User Creation
    catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] {
        Write-Host "In $csvFilePath at row $lineNumber" -ForegroundColor Red -BackgroundColor Black
        Write-Host "User '$UPN' already exists. Please choose a different username.`n+                         ~~~~~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Red -BackgroundColor Black
    }
# Other Errors
    catch{
        Write-Host "Failed to create user '$UPN' $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
    }finally{
        $lineNumber += 1
    }
}


