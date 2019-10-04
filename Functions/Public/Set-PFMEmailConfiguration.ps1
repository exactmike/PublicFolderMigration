function Set-PFMEmailConfiguration
{
    <#
    .SYNOPSIS
    Configures the PFM Module Email Settings for Email Reports when using the -SendEmail Parameter of Get-*Report* functions.
    .DESCRIPTION
    Configures the PFM Module Email Settings for Email Reports when using the -SendEmail Parameter of Get-*Report* functions.  Required before using the -SendEmail parameter of these functions
    .PARAMETER To
    When SendEmail is used, this sets the recipients of the email report.
    .PARAMETER From
    When SendEmail is used, this sets the sender of the email report.
    .PARAMETER SmtpServer
    When SendEmail is used, this is the SMTP Server to send the report through.
    .PARAMETER Subject
    Sets the subject of the email report.  If not set, the default subject is "Public Folder Replication Report from Exchange Organization ______" where ______ is the Identity property of Get-OrganizationConfig.
    .PARAMETER IncludeAttachment
    When SendEmail is used, specifying this switch will set the email report to not include the output files as attachment(s) to the email.
    .EXAMPLE
    Set-PFMEmailConfiguration -To ReportConsumer@contoso.com -From USCLTEX10PF01.us.clt.contoso.com@contoso.com -SmtpServer relay.contoso.com -Subject 'Public Folder Report' -Attachments

    Gets public folder tree data from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "Creates in-memory object only."
    )]
    [CmdletBinding(
        ConfirmImpact = 'None'
    )]
    param (
        [parameter(Mandatory)]
        [ValidateScript( { $_ | TestEmailAddress })]
        [string[]]$To
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestEmailAddress -EmailAddress $_ })]
        [string]$From
        ,
        [parameter(Mandatory)]
        [ValidateScript( { TestTCPConnection -port 25 -ComputerName $_ })]
        [string]$SMTPServer
        ,
        [parameter()]
        [string]$Subject
        ,
        [parameter()]
        [switch]$BodyAsHTML
<#         ,
        [parameter()]
        [switch]$IncludeAttachment #>
    )

    $script:EmailConfiguration = @{
        SMTPServer        = $SMTPServer
        To                = $To
        From              = $From
        BodyAsHTML      = $BodyAsHTML
    }
    if ([string]::IsNullOrEmpty($Subject))
    {
        $script:EmailConfiguration.Subject = 'Public Folder Environment and Replication Status Report'
    }

}