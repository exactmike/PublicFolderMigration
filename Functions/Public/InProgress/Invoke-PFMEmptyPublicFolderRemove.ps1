<#
    .SYNOPSIS
    Processes Empty Public Folder Information Objects for possible removal
    .DESCRIPTION
    Accepts one or more EntryIDs or other Public Folder Unique Identifiers and processes them for removal if they meet the required validations.
    .PARAMETER PublicFolderMailboxServer
    This parameter specifies the Exchange server from which to retrieve folder information to generate the Public Folder Information Objects.
    .PARAMETER Identity
    This parameter specifies the identity(ies) of the public folder(s) to be validated for and processed for removal
    .PARAMETER Passthru
    Controls whether the public folder validation objects are returned to the PowerShell pipeline for further processing.
    .PARAMETER OutputFolderPath
    Mandatory parameter for the already existing directory location where you want public folder replication and stats reports to be placed.  Operational log files will also go to this location.
    .PARAMETER OutputFormat
    Mandatory parameter used to specify whether you want csv, json, xml, clixml or any combination of these.
    .EXAMPLE
    Connect-PFMExchange -ExchangeOnPremisesServer USCLTEX10PF01.us.clt.contoso.com -credential $cred
    Get-PFMPublicFolderTree -OutputFolderPath c:\PFReports -OutputFormats csv,json,xml -PublicFolderMailboxServer USCLTEX10PF01

    If public folders are on Exchange 2010, the ExchangeOnPremisesServer must be an Exchange 2010 server.
    Gets public folder tree data from USCLTEX10PF01.us.clt.contoso.com and exports it to csv, json, and xml formats in c:\PFReports
    #>