
function New-MaintenenceAcct{
    
    $Password = Get-SecCredentials "$env:ProgramData\OTS\data\1.dat"
    if ($Password -eq $null)
    {
        throw "Not able to retrive password"
    }
    $user = @{
        Name        = 'Maintenence'
        Password    = $Password
        FullName    = 'Maintenence Account'
        Description = 'Created to run office repair'
    }
    # $Result = (Get-LocalUser -Name $user.Name -ErrorAction SilentlyContinue -ErrorVariable err)
    if ((Get-LocalUser -Name $user.Name -ErrorAction SilentlyContinue -ErrorVariable GetUserErr) -eq $null){
        New-LocalUser @user
        Add-LocalGroupMember -Group "Administrators" -Member $user.Name
    }
    else {
        throw "User $($user.Name) already exists. $GetUserErr"
    }
}

