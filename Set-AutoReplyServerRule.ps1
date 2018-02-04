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

<#
.SYNOPSIS
    Creates Microsoft Exchange server side auto-reply rule.
    
.DESCRIPTION  
    Creates a custom Microsoft Exchange auto-reply server-side message rule in batch.
    
    The script requires Microsoft Exchange Web Services Managed API 2.2:
    https://www.microsoft.com/en-in/download/details.aspx?id=42951

    It reads account information from provided CSV file and expects the following columns:
    
    First_Name
    Last_Name
    Phone_Number
    Email_Address
    Username
    Password

.NOTES
    Authors: Marcin Wisniowski (@wiredmind)
    License: ALv2  
  
.PARAMETER Path
    Specify the CSV file to import data from.
    
.EXAMPLE
    .\Set-AutoReplyServerRule.ps1 -Path .\UserList.csv
#>
param
(
[Parameter(
    Mandatory=$true,
    Position=0,
    ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true)
]
[string]$Path
)

Import-Module "$Env:ProgramFiles\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$userList = Import-Csv $Path

# Auto-reply email HTML template
$htmlBodyString = @"
<!doctype html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<head>
  <title></title>
  <!--[if !mso]><!-- -->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <!--<![endif]-->
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style type="text/css">
    #outlook a {{
      padding: 0
    }}

    .ReadMsgBody {{
      width: 100%
    }}

    .ExternalClass {{
      width: 100%
    }}

    .ExternalClass * {{
      line-height: 100%
    }}

    body {{
      margin: 0;
      padding: 0;
      -webkit-text-size-adjust: 100%;
      -ms-text-size-adjust: 100%
    }}

    table,
    td {{
      border-collapse: collapse;
      mso-table-lspace: 0;
      mso-table-rspace: 0
    }}

    img {{
      border: 0;
      height: auto;
      line-height: 100%;
      outline: 0;
      text-decoration: none;
      -ms-interpolation-mode: bicubic
    }}

    p {{
      display: block;
      margin: 13px 0
    }}
  </style>
  <!--[if !mso]><!-->
  <style type="text/css">
    @media only screen and (max-width:480px) {{
      @-ms-viewport {{
        width: 320px
      }}
      @viewport {{
        width: 320px
      }}
    }}
  </style>
  <!--<![endif]-->
  <!--[if mso]>
<xml>
  <o:OfficeDocumentSettings>
    <o:AllowPNG/>
    <o:PixelsPerInch>96</o:PixelsPerInch>
  </o:OfficeDocumentSettings>
</xml>
<![endif]-->
  <!--[if lte mso 11]>
<style type="text/css">
  .outlook-group-fix {{
    width:100% !important;
  }}
</style>
<![endif]-->
  <style type="text/css">
    @media only screen and (min-width:480px) {{
      .mj-column-per-100 {{
        width: 100%!important
      }}
    }}
  </style>
</head>

<body>
  <div class="mj-container">
    <!--[if mso | IE]>
      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="600" align="center" style="width:600px;">
        <tr>
          <td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;">
      <![endif]-->
    <div style="margin:0 auto;max-width:600px">
      <table role="presentation" cellpadding="0" cellspacing="0" style="font-size:0;width:100%" align="center" border="0">
        <tbody>
          <tr>
            <td style="text-align:center;vertical-align:top;direction:ltr;font-size:0;padding:20px 0">
              <!--[if mso | IE]>
      <table role="presentation" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td style="vertical-align:top;width:600px;">
      <![endif]-->
              <div class="mj-column-per-100 outlook-group-fix" style="vertical-align:top;display:inline-block;direction:ltr;font-size:13px;text-align:left;width:100%">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
                  <tbody>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px" align="center">
                        <table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-spacing:0" align="center" border="0">
                          <tbody>
                            <tr>
                              <td style="width:100px"><img alt="HRC Financial Group Logo" height="auto" src="http://hrcfinancialgroup.com/sites/all/themes/hrc/logo.png" style="border:none;border-radius:0;display:block;font-size:13px;outline:0;text-decoration:none;width:100%;height:auto" width="100"></td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px">
                        <p style="font-size:1px;margin:0 auto;border-top:4px solid #516572;width:100%"></p>
                        <!--[if mso | IE]><table role="presentation" align="center" border="0" cellpadding="0" cellspacing="0" style="font-size:1px;margin:0px auto;border-top:4px solid #516572;width:100%;" width="600"><tr><td style="height:0;line-height:0;"> </td></tr></table><![endif]-->
                      </td>
                    </tr>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px" align="left">
                        <div style="cursor:auto;color:#333444;font-family:helvetica;font-size:14px;line-height:22px;text-align:left">Thank you for contacting HRC. If your email is regarding Center Coast Capital Advisors ("Center Coast") or any fund or SMA managed by Center Coast, please contact <b>{0} {1}</b> at Brookfield Investment Management Inc.:</div>
                      </td>
                    </tr>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px" align="left">
                        <div style="cursor:auto;color:#333444;font-family:helvetica;font-size:14px;line-height:22px;text-align:left"><b>{0} {1}</b><br>T: <a href="tel:{2}" style="color:#516572">{2}</a><br>E: <a href="mailto:{3}" style="color:#516572">{3}</a></div>
                      </td>
                    </tr>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px" align="left">
                        a<div style="cursor:auto;color:#333444;font-family:helvetica;font-size:14px;line-height:22px;text-align:left">If your email is not regarding Center Coast, another HRC representative will be contacting you shortly.<br></div>
                      </td>
                    </tr>
                    <tr>
                      <td style="word-wrap:break-word;font-size:0;padding:10px 25px" align="left">
                        <div style="cursor:auto;color:#333444;font-family:helvetica;font-size:14px;line-height:22px;text-align:left">Regards,<br><span style="font-variant:small-caps">HRC Financial Group</span></div>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]>
      </td></tr></table>
      <![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]>
      </td></tr></table>
      <![endif]-->
  </div>
</body>

</html>
"@

foreach ($user in $userList) {

    $psCred = New-Object System.Management.Automation.PSCredential(
        $user.Username, 
        (ConvertTo-SecureString $user.Password -AsPlainText -Force))
    $cred = New-Object System.Net.NetworkCredential(
        $psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
    
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
    $service.Credentials = $cred
    $service.AutodiscoverUrl($user.Username, {$true})
    
    $templateEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
    $templateEmail.ItemClass = "IPM.Note.Rules.ReplyTemplate.Microsoft"
    $templateEmail.IsAssociated = $true
    # Auto-reply email Subject
    $templateEmail.Subject = "AUTOMATIC REPLY—PLEASE READ"
    $templateEmail.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody(
        $htmlBodyString -f $user.First_Name, $user.Last_Name, $user.Phone_Number, $user.Email_Address)
    
    # PidTagReplyTemplateId Canonical Property
    # https://technet.microsoft.com/en-us/library/cc815476(v=office.15).aspx
    $PidTagReplyTemplate = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
        0x65C2,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    
    $templateEmail.SetExtendedProperty($PidTagReplyTemplate, [System.Guid]::NewGuid().ToByteArray())
    $templateEmail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

    $inboxRule = New-Object Microsoft.Exchange.WebServices.Data.Rule
    # Inbox rule display name
    $inboxRule.DisplayName = "Termination Auto Reply"
    $inboxRule.IsEnabled = $true
    $inboxRule.Conditions.SentToOrCcMe = $true
    $inboxRule.Actions.ServerReplyWithMessage = $templateEmail.Id

    $createRule = New-Object Microsoft.Exchange.WebServices.Data.CreateRuleOperation[] 1    
    $createRule[0] = $inboxRule

    $service.UpdateInboxRules($createRule, $true)    

}