@description('Deployment location.  Default location will be the location of the resource group.')
param location string = toLower(resourceGroup().location)

@description('Azure AD Tenant Id')
param tenantId string = tenant().tenantId

@description('Azure AD Application Principal Client/Application Id')
param clientId string

@description('Certificate Thumbprint')
param certificateThumbprint string

@description('Certificate value, base64 encoded')
param certificateBase64Value string

@description('Teams admin user name')
param username string

@description('Teams admin password.')
@secure()
param password string

@description('Teams policy name.')
param policyName string

@description('Azure AD Group ObjectId.  Separate multiple values with a semi-colon.')
param groupId string

var suffix = toLower(uniqueString(resourceGroup().id))

module automation_account './automation.bicep' = {
  name: 'automation-${suffix}'
  params: {
    name: 'automation-${suffix}'
    location: location
    sku: 'Basic'
    runbooks: [
    ]
    modules: [
      {
        name:    'Microsoft.Graph.Groups'
        version: 'latest'
        uri:     'https://www.powershellgallery.com/api/v2/package'
      }
      {
        name:    'MicrosoftTeams'
        version: 'latest'
        uri:     'https://www.powershellgallery.com/api/v2/package'
      }
    ]
    variables:[
      {
        name: 'CLIENTID'
        value: clientId
        isEncrypted: 'false'
      }
      {
        name: 'TENANTID'
        value: tenantId
        isEncrypted: 'false'
      }
      {
        name: 'POLICYNAME'
        value: policyName
        isEncrypted: 'false'
      }
      {
        name: 'GROUPID'
        value: groupId
        isEncrypted: 'false'
      }
    ] 
    certificates: [
      {
        name: 'PFX'
        thumbprint: certificateThumbprint
        base64value: certificateBase64Value
      }
    ]
    credentials: [
      {
        name: 'teams-admin'
        username: username
        password: password
      }
    ]
  }
}
