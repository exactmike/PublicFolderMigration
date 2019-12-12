Function NewPermissionExportObject
{

    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        $TargetPublicFolder
        ,
        [parameter()]
        [AllowNull()]
        $TargetMailPublicFolder
        ,
        [parameter(Mandatory)]
        [string]$TrusteeIdentity
        ,
        [parameter()]
        [AllowNull()]
        $TrusteeRecipientObject
        ,
        [parameter(Mandatory)]
        [ValidateSet('FullAccess', 'SendOnBehalf', 'SendAs', 'None', 'ClientPermission')]
        $PermissionType
        ,
        [parameter()]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$AccessRights
        ,
        [parameter()]
        [ValidateSet('Direct', 'GroupMembership', 'None', 'Undetermined')]
        [string]$AssignmentType = 'Direct'
        ,
        $TrusteeGroupObjectGUID
        ,
        $ParentPermissionIdentity
        ,
        [string]$SourceExchangeOrganization = $ExchangeOrganization
        ,
        [boolean]$IsInherited = $False
        ,
        [switch]$none

    )#End Param
    $Script:PermissionIdentity++
    $PermissionExportObject =
    [pscustomobject]@{
        PermissionIdentity          = $Script:PermissionIdentity
        ParentPermissionIdentity    = $ParentPermissionIdentity
        SourceExchangeOrganization  = $SourceExchangeOrganization
        TargetEntryID               = $TargetPublicFolder.EntryID
        TargetPublicFolderPath      = $TargetPublicFolder.Identity
        TargetObjectGUID            = ''
        TargetObjectExchangeGUID    = ''
        TargetDistinguishedName     = ''
        TargetPrimarySMTPAddress    = ''
        TargetRecipientType         = ''
        TargetRecipientTypeDetails  = ''
        PermissionType              = $PermissionType
        AccessRights                = $AccessRights
        AssignmentType              = $AssignmentType
        TrusteeGroupObjectGUID      = $TrusteeGroupObjectGUID
        TrusteeIdentity             = $TrusteeIdentity
        IsInherited                 = $IsInherited
        TrusteeObjectGUID           = ''
        TrusteeExchangeGUID         = ''
        TrusteeDistinguishedName    = if ($None) { 'none' } else { '' }
        TrusteePrimarySMTPAddress   = if ($None) { 'none' } else { '' }
        TrusteeRecipientType        = ''
        TrusteeRecipientTypeDetails = ''
    }
    if ($null -ne $TargetMailPublicFolder)
    {
        $PermissionExportObject.TargetObjectGUID = $TargetMailPublicFolder.Guid.Guid
        $PermissionExportObject.TargetDistinguishedName = $TargetMailPublicFolder.DistinguishedName
        $PermissionExportObject.TargetPrimarySMTPAddress = $TargetMailPublicFolder.PrimarySmtpAddress  #.ToString()
        $PermissionExportObject.TargetRecipientType = $TargetMailPublicFolder.RecipientType
        $PermissionExportObject.TargetRecipientTypeDetails = $TargetMailPublicFolder.RecipientTypeDetails
    }
    if ($null -ne $TrusteeRecipientObject)
    {
        $PermissionExportObject.TrusteeObjectGUID = $TrusteeRecipientObject.guid.Guid
        $PermissionExportObject.TrusteeExchangeGUID = $TrusteeRecipientObject.ExchangeGuid.Guid
        $PermissionExportObject.TrusteeDistinguishedName = $TrusteeRecipientObject.DistinguishedName
        $PermissionExportObject.TrusteePrimarySMTPAddress = $TrusteeRecipientObject.PrimarySmtpAddress  #.ToString()
        $PermissionExportObject.TrusteeRecipientType = $TrusteeRecipientObject.RecipientType
        $PermissionExportObject.TrusteeRecipientTypeDetails = $TrusteeRecipientObject.RecipientTypeDetails
    }
    $PermissionExportObject

}
