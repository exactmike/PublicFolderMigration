$Script:PublicFolderPropertyList = @(
    @{n = 'EntryID'; e = { $_.EntryID.tostring() } }
    'Name'
    @{n = 'Identity'; e = { $_.Identity.tostring() } }
    @{n = 'MapiIdentity'; e = { $_.MapiIdentity.tostring() } }
    'ParentPath'
    'HasSubFolders'
    @{n = 'ReplicasString'; e = { $_.Replicas -join ';' } }
    'Replicas'
    @{n = 'ReplicaCount'; e = { $_.Replicas.count } }
    'UseDatabaseReplicationSchedule'
    @{n = 'ReplicationScheduleString'; e = { $_.ReplicationSchedule -join ';' } }
    'ReplicationSchedule'
    'PerUserReadStateEnabled'
    'FolderType'
    'MailEnabled'
    'HiddenFromAddressListsEnabled'
    'MaxItemSize'
    'UseDatabaseQuotaDefaults'
    'IssueWarningQuota'
    'ProhibitPostQuota'
    'UseDatabaseRetentionDefaults'
    'RetainDeletedItemsFor'
    'UseDatabaseAgeDefaults'
    'AgeLimit'
    'HasRules'
    'HasModerator'
    'IsValid'
)
