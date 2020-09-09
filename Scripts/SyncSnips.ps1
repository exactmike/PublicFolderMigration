$EXOOrganizationConfig = Invoke-Command -Session $EXOSession -ScriptBlock { Get-OrganizationConfig } # Examine PublicFoldersLockedForMigration
$EXPOrganizationConfig = Invoke-Command -Session $EXPSession -ScriptBlock { Get-OrganizationConfig } # Examine PublicFoldersLockedForMigration
$EXOAcceptedDomains = Invoke-Command -Session $EXOSession -ScriptBlock { Get-AcceptedDomain }
$EXPAcceptedDomains = Invoke-Command -Session $EXPSession -ScriptBlock { Get-AcceptedDomain }
$EXOPublicFolders = Invoke-Command -Session $EXOSession -ScriptBlock { Get-PublicFolder -Recurse -ResultSize Unlimited }
$EXPPublicFolders = Invoke-Command -Session $EXPSession -ScriptBlock { Get-PublicFolder -recurse -ResultSize Unlimited }
$EX