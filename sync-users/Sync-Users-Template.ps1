[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty]
    [string]
    $CompanyId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty]
    [string]
    $Psk,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty]
    [string]
    $UserRegistry,

    # Add this parameter to preview the detected changes without applying them.
    [switch]
    $WhatIf,

    # Add this parameter to run the script in silent/forced mode, so it proceeds without prompting you to confirm each step.
    [switch]
    $Confirm
)

Write-Host "LogMeIn Central - Synchronize Users from outside source - SCRIPT TEMPLATE"

# This script creates, moves, and deletes LogMeIn Central users by matching the 'Email' property
# of the provided user list to the list of LogMeIn Central users.

# The main goal of this script is that you set the desired LogMeIn Central permissions on 
# User groups and let the users inherit those permissions by putting them into the right
# group.
# This way, you don't have to represent any LogMeIn Central permission in your user registry.

# If you do not find this solution flexible enough, you may set permissions
# programatically on individual users.

# See https://support.logmein.com/central/help/logmein-central-developer-center?highlight=Update+user+settings
# An example API call:
        # $reqBody = @{
        #     email  = "john.doe@example.com"
        #     groupId = -1 # Default group ID: Don't put the user in any particular group
        #     permissions = @{ 
        #         grantAll = $true 
        #         central = @{
        #         enableCentral = $true 
        #         reports = $false
        #         alertManagement = $false 
        #         configurationManagement = $true 
        #         computerGroupManagement = $true 
        #         viewInventoryData = $false 
        #         inventoryManagement = $false 
        #         one2manyManagement = $false 
        #         one2manyRun = $false 
        #         windowsUpdateManagement = $true 
        #         applicationUpdateManagement = $true 
        #         antivirusManagement = $false 
        #         remoteExecution = $true 
        #         remoteExecutionCreateAndRun = $false 
        #         } 
        #         management = @{ 
        #         userManagement = $true
        #         loginPolicyManagement = $true 
        #         saveLoginCredentials = $false 
        #         createDesktopShortcut = $false 
        #         deployment = $false 
        #         adhocSupport = $true
        #         accountSecurity = $false 
        #         }
        #         interface = "advanced"
        #         groupAndComputerPermissions = @{ 
        #         allowFullRemoteControl = $true 
        #         computerPermission = "specified" 
        #         permittedGroupIds = @(1001, 1002, 1003) # Assuming these are Computer Group IDs
        #         permittedHostIds = @(2001, 2002, 2003, 2004) # Assuming these are Computer IDs
        #         }
        #         network = @{
        #         accessNetworks = $false 
        #         networkAndClientManagement = $false 
        #         editClientDefaults = $false 
        #         editNetworkDefaults = $false
        #         }
        #         enforceTfa = $false 
        #     }
        # } | ConvertTo-Json -Compress

        # $addUsersResponse = Invoke-WebRequest `
        #     -ContentType "application/json"  `
        #     -Method Put `
        #     -Uri "$CentralHost/public-api/v3/users/details/set" `
        #     -Headers $Headers `
        #     -UseBasicParsing `
        #     -Body $reqBody


# We use the Invoke-WebRequest commandlet to issue an HTTP request to LogMeIn Central.
# For more information, see https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest?view=powershell-7.5

# This script makes calls to the LogMeIn Central Public API.
# For each API call, there is a comment "LogMeIn Central API: <url>" linking
# to the documentation of the relevant API endpoint.


# Stop the script execution when an error encountered.
$ErrorActionPreference = "Stop"


# Split up the $Collection into arrays of $CunkSize lengths (the last one might be shorter)
function Chunk {
    param (
        $Collection,

        [int]
        $ChunkSize
    )

    $idx = 0
    $chunk = [System.Collections.Generic.List[psobject]]::new()
    $res = [System.Collections.Generic.List[psobject]]::new()

    foreach ($item in $Collection) {
        $chunk.Add($item)
        $idx += 1

        if (($idx % $ChunkSize) -eq 0) {
            $res.Add($chunk)
            $chunk = [System.Collections.Generic.List[psobject]]::new()
        }
    }

    if ($chunk.Count -gt 0) {
        $res.Add($chunk)
    }

    return ,$res
}

function Get-Confirmation {
    param (
        [string]
        $Prompt,

        $Users,

        [scriptblock]
        $Format,

        [bool]
        $WhatIf,

        [bool]
        $Confirm
    )

    Write-Host $Prompt
    if ($Users.Count -eq 0) {
        Write-Host "N/A"
        return $false
    }

    $Users | ForEach-Object { &$Format -U $_ } | Write-Host

    if ($WhatIf) {
        return $false
    }

    if ($Confirm) {
        return $true
    }

    $PossibleAnswers = @("Y", "S", "A")

    do {
        $answer = Read-Host -Prompt "`nDo you want to continue? ([Y]es / [S]kip this step / [A]bort and exit)"
    } while (-not ($PossibleAnswers -contains $answer))

    switch ($answer) {
        "Y" { return $true }
        "S" { return $false }
        Default { throw "Aborted by user" }
    }
}

function Read-UserRegistry {
    param (
        $Source
    )

    # This is a placeholder for your business logic.
    # You may want to make API calls to connect to a third-party Identity service,
    # or read the relevant user information from Active Directory.
    # It is up to you what you want to make the "source of thruth" of your user registry.

    # The only criteria is that users should have 'Email' and 'Group' properties.

    # Here is a simple logic that reads a JSON file's content.
    # An example of the JSON file
    # [
    #   { "Email": "john.doe@example.com", "Group": "" },
    #   { "Email": "jane.doe@example.com", "Group": "Admins" },
    #   ...
    # ]
    $Users = Get-Content $UserRegistry -Raw | ConvertFrom-Json
    return ,$Users
}

function Get-GroupId {
    param (
        $CentralUserGroups,

        [string]
        $GroupName
    )

    if ([string]::IsNullOrEmpty($GroupName)) {
        # When defining a null or empty string as the group's name, users will be invited to the Default group.
        return -1
    }

    $TargetGroup = $CentralUserGroups |
        Where-Object { $_.groupName -eq $GroupName } |
        Select-Object -First 1

    return $TargetGroup.GroupId
}


# Get the current state of your user directory.
$ExternalUsers = Read-UserRegistry -Source $UserRegistry

# Preparing the Public API authentication method.
# We need to send the "Authorize" HTTP header with each API call.
# See https://support.logmein.com/central/help/logmein-central-developer-center?highlight=authentication
$credentials = $CompanyId + ":" + $PSK;
$credentialsBytes = [System.Text.ASCIIEncoding]::ASCII.GetBytes($credentials)
$credentialsBase64 = [System.Convert]::ToBase64String($credentialsBytes)
$AuthHeader = "Basic $credentialsBase64"
$Headers = @{
    Authorization = $authHeader
}

# This is the base address of the LogMeIn Central web servers.
# All Public API endpoints are hosted under this address.
$CentralHost = "https://secure.logmein.com"


# Get the current list of user groups from LogMeIn Central
# LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=Get+a+List+of+User+groups
$CentralUserGroups = Invoke-WebRequest "$CentralHost/public-api/v3/user-groups" `
    -Headers $Headers `
    -UseBasicParsing
$CentralUserGroups = ConvertFrom-Json $CentralUserGroups

# Create a flat user list for easier lookup.
$CentralUsers = [System.Collections.Generic.List[psobject]]::new()
foreach ($group in $CentralUserGroups) {
    foreach ($user in $group.userList) {
        $CentralUsers.Add([PSCustomObject]@{
            "GroupId" = $group.groupId
            "GroupName" = $group.groupName
            "NewGroupName" = $null
            "Email"   = $user.email
        })
    }
}

# ============================================================================
# Add missing users
# ============================================================================
$UsersToBeAdded = [System.Collections.Generic.List[psobject]]::new()
foreach ($externalUser in $ExternalUsers) {
    $found = $false
    foreach ($centralUser in $CentralUsers) {
        if ($centralUser.Email -eq $externalUser.Email) {
            $found = $true
            break
        }
    }

    if ($found -eq $false) {
        $UsersToBeAdded.Add($externalUser)
    }
}

$doAddUsers = Get-Confirmation -Users $UsersToBeAdded `
    -Prompt "`nUsers to be ADDED to LogMeIn Central:" `
    -Format { $_.Email } `
    -WhatIf $WhatIf -Confirm $Confirm

if ($doAddUsers) {
    # Invite users in batches, always into the Default group. They will be moved into their
    # respective user group at a later stage
    $userBatches = Chunk -Collection $UsersToBeAdded -ChunkSize 100
    foreach ($userBatch in $userBatches) {
        $emails = @($userBatch | ForEach-Object { $_.Email })

        $reqBody = @{
            emails  = $emails
            groupId = -1 # Default group ID: Don't put the user in any particular group
            permissions = @{
                grantAll = $false
            }
        } | ConvertTo-Json -Compress

        # LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=invite+users
        $addUsersResponse = Invoke-WebRequest `
            -ContentType "application/json"  `
            -Method Post `
            -Uri "$CentralHost/public-api/v3/users/invitation" `
            -Headers $Headers `
            -UseBasicParsing `
            -Body $reqBody

        # Check whether certain users were not invited for some reason
        # If so, list the relevant email addresses and exit with an error
        $addUsersResponse = $addUsersResponse.Content | ConvertFrom-Json
        if ($addUsersResponse.NotInvitedEmails.lengths -gt 0) {
            Write-Error "Failed to add all users", ($addUsersResponse.NotInvitedEmails | Format-Table)
            throw "Failed to add all users"
        }

        # Add the newly invited users to the local list so that we can put them
        # into the correct user group at a later step
        foreach($User in $userBatch) {
            $CentralUsers.Add([PSCustomObject]@{
                "GroupId" = $CentralUserGroups[0].GroupId # The first group is the Default group
                "GroupName" = $CentralUserGroups[0].GroupName # The Default group's name
                "NewGroupName" = $null
                "Email"   = $User.Email
            })
        }
    }
}


# ============================================================================
# Delete extrenious users
# ============================================================================
$UsersToBeDeleted = [System.Collections.Generic.List[psobject]]::new()
foreach ($centralUser in $CentralUsers) {
    $found = $false
    foreach ($externalUser in $ExternalUsers) {
        if ($centralUser.Email -eq $externalUser.Email) {
            $found = $true
            break
        }
    }

    if ($found -eq $false) {
        $UsersToBeDeleted.Add($centralUser)
    }
}

$doDeleteUsers = Get-Confirmation -Users $UsersToBeDeleted `
    -Prompt "`nUsers to be DELETED from LogMeIn Central:" `
    -Format { $_.Email } `
    -WhatIf $WhatIf -Confirm $Confirm

if ($doDeleteUsers) {
    $UserBatches = Chunk -Collection $UsersToBeDeleted -ChunkSize 50
    foreach ($UserBatch in $UserBatches) {
        $emails = @($UserBatch | ForEach-Object { $_.Email })
        $reqBody = ConvertTo-Json $emails -Compress

        # Delete the users with this API call.
        # Alternatively, you can disable users instead of deleting them.
        # However, you would have to take the 'enabled' status into account when comparing user lists.

        # LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=delete+users
        $response = Invoke-WebRequest `
            -ContentType "application/json"  `
            -Method Delete `
            -Uri "$CentralHost/public-api/v3/users" `
            -Headers $Headers `
            -UseBasicParsing `
            -Body $reqBody

        # Delete the users from the local list so they are not taken into account in the subsequent steps
        foreach($User in $UserBatch) {
            $CentralUsers.Remove($User)
        }
    }
}


# ============================================================================
# Move users into the appropriate user groups and create groups if necessary.
# ============================================================================
$UsersToBeMoved = [System.Collections.Generic.List[psobject]]::new()
foreach ($externalUser in $ExternalUsers) {
    # Handle emptry string or null as the Default group
    if ([System.String]::IsNullOrEmpty($externalUser.Group)) {
        $externalUser.Group = $CentralUserGroups[0].GroupName
    }

    foreach($centralUser in $CentralUsers) {
        if ($externalUser.Email -eq $centralUser.Email) {
            if ($externalUser.Group -ne $centralUser.GroupName) {
                $centralUser.NewGroupName = $externalUser.Group
                $UsersToBeMoved.Add($centralUser)
                break
            }
        }
    }
}

$doMoveUsers = Get-Confirmation -Users $UsersToBeMoved `
    -Prompt "`nUsers to be MOVED into another user group:" `
    -Format { "$($_.Email): `"$($_.GroupName)`" -> `"$($_.NewGroupName)`"" } `
    -WhatIf $WhatIf -Confirm $Confirm

if ($doMoveUsers) {

    # First, we check whether each target user group exists in LogMeIn Central
    # We then create the missing groups if necessary
    $CentralUserGroupNames = @($CentralUserGroups | ForEach-Object { $_.GroupName })
    $MissingGroupNames = @($UsersToBeMoved |
        ForEach-Object { $_.NewGroupName } |
        Select-Object -Unique |
        Where-Object { $_ -notin $CentralUserGroupNames })

    if ($MissingGroupNames.Count -gt 0) {
        foreach($missingGroupName in $MissingGroupNames) {
            Write-Host "Create missing user group `"$missingGroupName`""

            $reqBody = @{ name = $missingGroupName } | ConvertTo-Json -Compress

            # LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=create+user+group
            $response = Invoke-WebRequest `
                -ContentType "application/json"  `
                -Method Post `
                -Uri "$CentralHost/public-api/v3/user-groups" `
                -Headers $Headers `
                -UseBasicParsing `
                -Body $reqBody
        }

        # Get a fresh list of users and groups to get the GroupIDs of the newly created groups
        # LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=Get+a+List+of+User+groups
        $CentralUserGroups = Invoke-WebRequest "$CentralHost/public-api/v3/user-groups" `
            -Headers $Headers `
            -UseBasicParsing
        $CentralUserGroups = ConvertFrom-Json $CentralUserGroups
    }


    $targetGroups = $UsersToBeMoved | Group-Object -Property "NewGroupName"
    foreach($targetGroup in $targetGroups) {
        $centralTargetGroupId = Get-GroupId -CentralUserGroups $CentralUserGroups -GroupName $targetGroup.Name

        $emails = @($targetGroup.Group | ForEach-Object { $_.Email })
        $reqBody = ConvertTo-Json $emails -Compress

        # LogMeIn Central API: https://support.logmein.com/central/help/logmein-central-developer-center?highlight=change+group+membership+of+multiple+users
        $response = Invoke-WebRequest `
            -ContentType "application/json"  `
            -Method Post `
            -Uri "$CentralHost/public-api/v3/user-groups/move-users/$($centralTargetGroupId)" `
            -Headers $Headers `
            -UseBasicParsing `
            -Body $reqBody
    }
}

