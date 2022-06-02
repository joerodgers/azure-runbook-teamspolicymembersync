#requires -modules "Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "MicrosoftTeams" 

function Split-Array
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [object[]]$Array,
        [Parameter(Mandatory=$true)]
        [int]$Size
    )

    begin
    {
        $maxChunks = [Math]::Ceiling( ($Array.Count/$Size) )

        $chunks = New-Object object[] $maxChunks
    }
    process
    {

        Write-Verbose "$(Get-Date) - Total Chunks: $maxChunks"

        for($i = 0; $i -lt $maxChunks; $i++ )
        {
            $sourceIndex = $i * $Size
            $length = [Math]::Min( ($Array.Length - $sourceIndex), $Size)

            $chunk = New-Object object[] $length

            [Array]::Copy($Array, $sourceIndex, $chunk, 0 <# destinationIndex #>, $length)

            $chunks[$i] = $chunk
        }
    }
    end
    {
        return ,$chunks
    }    
}

function Get-TeamsAppPermissionPolicyMember
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$TeamsAppPermissionPolicyName
    )

    begin
    {
    }
    process
    {
        Write-Verbose "$(Get-Date) - Querying memembers of TeamsAppPermissionPolicy '$TeamsAppPermissionPolicy'"

        Get-CsOnlineUser | Where-Object -Property TeamsAppPermissionPolicy -eq $TeamsAppPermissionPolicyName

        Write-Verbose "$(Get-Date) - Querying memembers of TeamsAppPermissionPolicy '$TeamsAppPermissionPolicy' completed"
    }
    end
    {
    }    
}

function Get-GroupMember
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$GroupId
    )

    begin
    {
    }
    process
    {
        $totalMembers = 0

        $GroupId | Foreach-Object {  

            Write-Verbose "$(Get-Date) - Querying memembers of group '$GroupId'"

            $members = @(Get-MgGroupMember -GroupId $_ -All) 
            $members

            $totalMembers += $members.Count

            Write-Verbose "$(Get-Date) - Querying memembers of group '$GroupId' completed. Member count: $($members.Count)"
        }

        Write-Verbose "$(Get-Date) - Total group members: $totalMembers"
    }
    end
    {
    }    
}

function Get-TeamsAppPermissionPolicyMembersToAdd
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$PolicyMemberObjectIds,

        [Parameter(Mandatory=$true)]
        [string[]]$GroupMemberObjectIds
    )

    begin
    {
    }
    process
    {
        $objectsIds = @(Compare-Object -ReferenceObject $PolicyMemberObjectIds -DifferenceObject $GroupMemberObjectIds | Where-Object -Property SideIndicator -eq "=>" | Select-Object -ExpandProperty InputObject)
        $objectsIds

        Write-Verbose "$(Get-Date) - Teams App Permission Policy Members To Add: $($objectsIds.Count)"
    }
    end
    {
    }    
}

function Get-TeamsAppPermissionPolicyMembersToRemove
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$PolicyMemberObjectIds,

        [Parameter(Mandatory=$true)]
        [string[]]$GroupMemberObjectIds
    )

    begin
    {
    }
    process
    {
        $objectsIds = @(Compare-Object -ReferenceObject $PolicyMemberObjectIds -DifferenceObject $GroupMemberObjectIds | Where-Object -Property SideIndicator -eq "<=" | Select-Object -ExpandProperty InputObject)
        $objectsIds

        Write-Verbose "$(Get-Date) - Teams App Permission Policy Members To Remove: $($objectsIds.Count)"
    }
    end
    {
    }    
}

function Wait-BatchPolicyAssignmentOperationComplete
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateRange(1,50)]
        [int]$MaxConcurrentBatches
    )

    begin
    {
    }
    process
    {
        Write-Verbose "$(Get-Date) - Teams App Permission Policy Members To Remove: $($objectsIds.Count)"
    
        $batches = @(Get-CsBatchPolicyAssignmentOperation | Where-Object -Property OverallStatus -ne "Completed")
       
        while( $null -ne $batches -and $batches.Count -ge $MaxConcurrentBatches )
        {
            Write-Warning "$(Get-Date) - $($batches.Count) batches executing, waiting for $($batches.Count - $($MaxConcurrentBatches-1)) batch operations to complete before scheduling next batch operation."
            Start-Sleep -Seconds 5

            $batches = @(Get-CsBatchPolicyAssignmentOperation | Where-Object -Property OverallStatus -ne "Completed")
        }
    }
    end
    {
    }    
}

function Sync-TeamsAppPermissionPolicyMembers
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$TeamsAppPermissionPolicy,

        [Parameter(Mandatory=$false)]
        [AllowEmptyCollection()]
        [string[]]$ObjectIdsToAdd,

        [Parameter(Mandatory=$false)]
        [AllowEmptyCollection()]
        [string[]]$ObjectIdsToRemove,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,5000)]
        [int]$BatchSize = 5000,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,50)]
        [int]$MaxConcurrentBatches = 50
    )

    begin
    {
    }
    process
    {
        if( $PSBoundParameters.ContainsKey("ObjectIdsToRemove") -and $ObjectIdsToRemove.Count -gt 0 )
        {
            Write-Verbose "$(Get-Date) - Processing $($ObjectIdsToRemove.Count) removal objects"

            $counter = 1

            $chunks = Split-Array -Array $ObjectIdsToRemove -Size $BatchSize

            foreach( $chunk in $chunks )
            {
                Wait-BatchPolicyAssignmentOperationComplete -MaxConcurrentBatches $MaxConcurrentBatches

                $operationName = "Teams App Permission Policy Removal - Chunk $counter of $($chunks.Count)"
                Write-Verbose "$(Get-Date) - Creating New CsBatchPolicyAssignmentOperation: $operationName"

                $null = New-CsBatchPolicyAssignmentOperation -PolicyType TeamsAppPermissionPolicy -PolicyName $null -Identity $chunk -OperationName $operationName
                $counter++
            }
        }
        
        if( $PSBoundParameters.ContainsKey("ObjectIdsToAdd") -and $ObjectIdsToAdd.Count -gt 0 )
        {
            Write-Verbose "$(Get-Date) - Processing $($ObjectIdsToAdd.Count) addition objects"
            $counter = 1

            $chunks = Split-Array -Array $ObjectIdsToAdd -Size $BatchSize

            foreach( $chunk in $chunks )
            {
                Wait-BatchPolicyAssignmentOperationComplete -MaxConcurrentBatches $MaxConcurrentBatches

                $operationName = "Teams App Permission Policy Addition - Chunk $counter of $($chunks.Count)"
                Write-Verbose "$(Get-Date) - Creating New CsBatchPolicyAssignmentOperation: $operationName"

                $null = New-CsBatchPolicyAssignmentOperation -PolicyType TeamsAppPermissionPolicy -PolicyName $teamsAppPermissionPolicyName -Identity $chunk -OperationName $operationName
                $counter++
            }
        }
    }
    end
    {
    }    
}

$clientId   = Get-AutomationVariable     -Name 'CLIENTID'
$tenantId   = Get-AutomationVariable     -Name 'TENANTID'
$policyName = Get-AutomationVariable     -Name 'POLICYNAME'
$groupIds   = (Get-AutomationVariable    -Name 'GROUPID') -split ";"
$certficate = Get-AutomationCertificate  -Name 'PFX'
$credential = Get-AutomationPSCredential -Name 'teams-admin'

Write-Output "Configuration Variables"
Write-Output "ClientId:    '$clientId'"
Write-Output "Certificate: '$($certficate.Thumbprint)'"
Write-Output "TenantId:    '$tenantId'"
Write-Output "Policy:      '$policyName'"
Write-Output "Groups:      '$($groupIds -join ',')'"
Write-Output "UserName:    '$($credential.UserName)'"

# connections

    Write-Output "Connecting to Microsoft Graph"

    # requries microsoft.graph.authentication module v1.2.0 to expose the -Certificate parameter
    Connect-MGGraph `
        -ClientId    $clientId `
        -Certificate $certficate `
        -TenantId    $tenantId `
        -ErrorAction Stop


    Write-Output "Connecting to Microsoft Teams"

    Connect-MicrosoftTeams `
        -Credential ([PScredential]::new( $credential.UserName, $credential.Password ))`
        -ErrorAction Stop | Out-Null


# get all group members

    Write-Output "Reading Azure AD Group Members"
    $groupMemberObjectIds = @(Get-GroupMember -GroupId $groupIds | Select-Object -ExpandProperty Id -Unique)
    Write-Output "Read $($groupMemberObjectIds.Count) Group Members"

# get all users in the Teams App Permission Policy

    Write-Output "Reading Teams App Permission Policy '$policyName' members"
    $policyMemberObjectIds = @(Get-TeamsAppPermissionPolicyMember -TeamsAppPermissionPolicyName $policyName | Select-Object -ExpandProperty ObjectId | Select-Object -ExpandProperty Guid)
    Write-Output "Read $($policyMemberObjectIds.Count) Teams App Permission Policy members"

# get all users to remove

    Write-Output "Computing members to remove from Teams App Permission policy"
    $memberObjectIdsToRemove = Get-TeamsAppPermissionPolicyMembersToRemove -GroupMemberObjectIds $groupMemberObjectIds -PolicyMemberObjectIds $policyMemberObjectIds -Verbose
    Write-Output "Computed $($memberObjectIdsToRemove.Count) users to remove"

# get all users to add

    Write-Output "Computing members to Add to Teams App Permission policy"
    $memberObjectIdsToAdd = Get-TeamsAppPermissionPolicyMembersToAdd -GroupMemberObjectIds $groupMemberObjectIds -PolicyMemberObjectIds $policyMemberObjectIds -Verbose
    Write-Output "Computed $($memberObjectIdsToAdd.Count) users to add"

# execute sequential sync

    Write-Output "Syncing members with Teams App Permission policy"
    Sync-TeamsAppPermissionPolicyMembers -TeamsAppPermissionPolicy $policyName -ObjectIdsToAdd $memberObjectIdsToAdd -ObjectIdsToRemove $memberObjectIdsToRemove -BatchSize 2000 -MaxConcurrentBatches 50 -Verbose

# disconnect from cloud services

    Disconnect-MgGraph -ErrorAction Ignore
    Disconnect-MicrosoftTeams -ErrorAction Ignore