@description('Location of the automation account')
param location string = resourceGroup().location

@description('Automation account name')
param name string

@description('Automation account sku')
@allowed([
  'Free'
  'Basic'
])
param sku string = 'Basic'

@description('Modules to import into automation account')
@metadata({
  name: 'Module name'
  version: 'Module version or specify latest to get the latest version'
  uri: 'Module package uri, e.g. https://www.powershellgallery.com/api/v2/package'
})
param modules array = []

@description('Runbooks to import into automation account')
@metadata({
  runbookName: 'Runbook name'
  runbookUri:  'Runbook URI'
  runbookType: 'Runbook type: Graph, Graph PowerShell, Graph PowerShellWorkflow, PowerShell, PowerShell7, PowerShell Workflow, Script'
  logProgress: 'Enable progress logs'
  logVerbose:  'Enable verbose logs'
})
param runbooks array = []

@description('Certificates to import into automation account')
@metadata({
  name:    'Certificate Name'
  thumbprint: 'Certificate Thumbprint'
  base64Value: 'Base64 Value'
})
param certificates array = []

@description('Variables to import into automation account')
@metadata({
  name:  'Variable Name'
  value: 'Variable Value'
  isEncrypted: 'Stored Encrypted'
})
param variables array = []

@description('Credentials to import into automation account')
@metadata({
  username: 'User name'
  passwrod: 'Password'
})
param credentials array = []


resource automation_account 'Microsoft.Automation/automationAccounts@2020-01-13-preview' = {
  name: name
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    sku: {
      name: sku
    }
    encryption: {
      keySource: 'Microsoft.Automation'
    }
    publicNetworkAccess: true
  }
}

resource automation_account_modules 'Microsoft.Automation/automationAccounts/modules@2020-01-13-preview' = [for module in modules: {
  parent: automation_account
  name: module.name
  properties: {
    contentLink: {
      uri:     module.version == 'latest' ? '${module.uri}/${module.name}' : '${module.uri}/${module.name}/${module.version}'
      version: module.version == 'latest' ? null : module.version
    }
  }
}]

resource automation_runbooks 'Microsoft.Automation/automationAccounts/runbooks@2019-06-01' = [for runbook in runbooks: {
  parent: automation_account
  name: runbook.runbookName
  location: location
  properties: {
    runbookType: runbook.runbookType
    logProgress: runbook.logProgress
    logVerbose: runbook.logVerbose
    publishContentLink: {
      uri: runbook.runbookUri
    }
  }
}]

resource automation_certificates 'Microsoft.Automation/automationAccounts/certificates@2019-06-01' = [for certificate in certificates: {
  parent: automation_account
  name: certificate.name
  properties: {
    isExportable: false
    thumbprint: certificate.thumbprint
    base64Value: certificate.base64Value
  }
}]

resource automation_variables 'Microsoft.Automation/automationAccounts/variables@2020-01-13-preview' = [for variable in variables: {
  parent: automation_account
  name: variable.name
  properties: {
    value: '"${variable.value}"'
    isEncrypted: bool(variable.isEncrypted)
    description: variable.name
  }
}]

resource automation_credentials 'Microsoft.Automation/automationAccounts/credentials@2020-01-13-preview' = [for credential in credentials: {
  parent: automation_account
  name: credential.name
  properties: {
    userName: credential.username
    password: credential.password
  }
}]


