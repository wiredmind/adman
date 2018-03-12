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
    Creates new Active Directory user from an object.
    
  .DESCRIPTION  
    Creates new Active Directory user from an object.

  .NOTES
    Authors: Marcin Wisniowski (@wiredmind)
    License: ALv2  
  
  .PARAMETER InputObject
    Any kind of object contining following properties
      Name - user account name, e.g., Jane Doe
      GivenName - user given name, e.g., Jane
      Surname - user account last name, e.g., Doe
      SamAccountName - user account SamAccountName, e.g., jdoe
      Description - user account description, e.g., Managing Director of Operations
      EmailAddress - user account email address, e.g., jdoe@example.com
      Path - user account Organizational Unit path, e.g., "CN=Users,DC=example,DC=com"
    
  .EXAMPLE
    Import-Csv .\UserList.csv | Import-ADUser
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(
      Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)
    ]
    [PSCustomObject[]]$InputObject
  )

  BEGIN {}
  PROCESS 
  {
    foreach ($obj in $InputObject)
    {
        try
        {
          $properties = @{ 
            "Name" = [string]"$($obj.Name)";
            "DisplayName" = [string]"$($obj.Name)";
            "GivenName" = [string]"$($obj.GivenName)";
            "Surname" = [string]"$($obj.Surname)";
            "SamAccountName" = [string]"$($obj.SamAccountName)";
            "Description" = [string]"$($obj.Description)";
            "EmailAddress" = [string]"$($obj.EmailAddress)";
            "UserPrincipalName" = [string]("$($obj.SamAccountName)@$((Get-ADDomain).DNSRoot)");
            "Path" = [string]"$($obj.Path)";
            "AccountPassword" = ConvertTo-SecureString "$($obj.AccountPassword)" -AsPlainText -Force;
            "ChangePasswordAtLogon" = $false;
            "Enabled" = $true;
            "PasswordNeverExpires" = $true
          }
          $user = New-ADUser @properties -ErrorAction Stop
          Write-Output $user
        }
        catch
        {
          Write-Verbose "-----> IMPORT Error: User $($obj.SamAccountName)" -OutVariable Con
          Write-Error $_.Exception.Message
          continue
        }
    }
  }
  END {}
}
