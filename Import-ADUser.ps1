Function Import-ADUser
{
  param
  (
      $userListFile = "UserList.csv"
  )

  $users = Import-Csv $userListFile

  foreach ($u in $users)
  {
    $user = @{ 
      "Name" = [string]"$($u.Name)";
      "DisplayName" = [string]"$($u.Name)";
      "GivenName" = [string]"$($u.GivenName)";
      "Surname" = [string]"$($u.Surname)";
      "SamAccountName" = [string]"$($u.SamAccountName)";
      "Description" = [string]"$($u.Description)";
      "EmailAddress" = [string]"$($u.EmailAddress)";
      "UserPrincipalName" = [string]("$($u.SamAccountName)@$((Get-ADDomain).DNSRoot)");
      "Path" = [string]"$($u.Path)";
      "AccountPassword" = ConvertTo-SecureString "$($u.AccountPassword)" -AsPlainText -Force;
      "ChangePasswordAtLogon" = $false;
      "Enabled" = $true;
      "PasswordNeverExpires" = $true
    }

    try
    {
      $user = Get-ADUser $user.SamAccountName -ErrorAction SilentlyContinue
      Write-Error "-----> IMPORT ERROR: User $($user.SamAccountName) already exists." -Category InvalidArgument
    }
    catch
    {
      $user = New-ADUser @user
    }
  }
}
