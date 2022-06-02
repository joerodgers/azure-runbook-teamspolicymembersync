
param
(
    [Parameter(Mandatory=$false)]
    [string]
    $ResourceGroup = "RG_TEAMS_CALLCENTERPOLICYSYNC_EASTUS_PROD",

    [Parameter(Mandatory=$false)]
    [string]
    $TemplatePath = "main.bicep",

    [parameter(Mandatory=$true)]
    [Guid]
    $TenantId,

    [parameter(Mandatory=$true)]
    [Guid]
    $SubscriptionId,

    [parameter(Mandatory=$true)]
    [Guid]
    $ClientId,

    [parameter(Mandatory=$true)]
    [string]
    $CertificatePath,

    [parameter(Mandatory=$true)]
    [SecureString]
    $CertificatePassword,

    [parameter(Mandatory=$true)]
    [PSCredential]
    $TeamsCredential,

    [parameter(Mandatory=$true)]
    [string]
    $TeamsPolicy,

    [parameter(Mandatory=$true)]
    [Guid[]]
    $GroupId
)

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol   = [System.Net.SecurityProtocolType]::Tls12   

# ensure connection

    $ctx = Get-AzContext

    #if( $ctx.Tenant.Id -ne $TenantId.ToString() -or $ctx.Subscription.SubscriptionId -ne $SubscriptionId.ToString() )
    if( $ctx.Subscription.SubscriptionId -ne $SubscriptionId.ToString() )
    {
        Write-Host "[$(Get-Date)] - Prompting for Azure credentials"
        Login-AzAccount -Tenant $TenantId -WarningAction SilentlyContinue

        $ctx = Get-AzContext
    }

    $subscription = Select-AzSubscription -Subscription $SubscriptionId -WarningAction SilentlyContinue

    Write-Host "[$(Get-Date)] - Connected as: $($ctx.Account.Id)"
    Write-Host "[$(Get-Date)] - Subscription: $($subscription.Subscription.Name) $($subscription.Subscription.Id)"

# build parameters

    $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2( $CertificatePath, $CertificatePassword, "Exportable,PersistKeySet,MachineKeySet")

    $parameters = @{
        certificateBase64Value = [System.Convert]::ToBase64String( $pfx.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx) )
        certificateThumbprint  = $pfx.Thumbprint
        clientId               = $ClientId
        tenantId               = $TenantId
        username               = $TeamsCredential.UserName
        password               = $TeamsCredential.Password
        policyName             = $TeamsPolicy
        groupId                = ($GroupId -join ";")
    }

# start deployment

    Write-Host "[$(Get-Date)] - Starting deployment:"
    Write-Host "[$(Get-Date)] - `tTemplate Path: $($templatePath)"

    $deployment = New-AzResourceGroupDeployment `
                        -ResourceGroupName       $ResourceGroup `
                        -TemplateFile            $TemplatePath `
                        -TemplateParameterObject $parameters `
                        -ErrorAction Stop

    $deployment
