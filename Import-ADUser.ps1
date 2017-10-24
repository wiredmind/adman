# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.

function Import-ADUser
{
  <#
  .SYNOPSIS
    Creates new Active Directory users in batch.
    
  .DESCRIPTION  
    Creates a custom user object from CSV files and a new Active Directory users in batch.
    It reads account information from provided CSV file and expects the following columns:
    | Name     | GivenName | Surname | SamAccountName | Description     | EmailAddress     | Path                         |
    |:---------|:----------|:--------|:---------------|:----------------|:-----------------|:-----------------------------|
    | Jane Doe | Jane      | Doe     | jdoe           | Account Manager | jdoe@example.com | "CN=Users,DC=example,DC=com" |    

  .NOTES
    Authors: Marcin Wisniowski (@wiredmind)
    License: ALv2  
  
  .PARAMETER Path
    Specify the CSV file to import account data from. By default it will look for `UserList.csv`
    in current working directory.
    
  .EXAMPLE
    Import-ADUser -Path .\UserList.csv
  #>
  param
  (
    [Parameter(
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)
    ]
    [string]$Path = "UserList.csv"
  )
  
  $Users = Import-Csv $Path

  foreach ($u in $Users)
  {
    try
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
      $user = New-ADUser @user
    }
    catch
    {
      Write-Information "-----> IMPORT Error: User $($user.SamAccountName)" -InformationAction Continue
      Write-Error $_.Exception.Message
      continue
    }
  }
}
