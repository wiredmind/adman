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

Function Import-ADUser
{
  <#
  .Synopsis
  Creates new Active Directory users in batch.
  .Description
  Creates new Active Directory users in batch, reading account information from
  a csv file and passing to New-ADUser.
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
