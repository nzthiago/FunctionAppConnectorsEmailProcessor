param name string
param location string
param tags object = {}
param connectionName string = ''
param connectorName string = ''
param teamsConnectionName string = ''
param functionAppPrincipalId string = ''
param tenantId string = tenant().tenantId

resource aiGateway 'Microsoft.Web/aigateways@2026-03-01-preview' = {
  name: name
  location: location
  tags: tags
}

resource aiGatewayConnection 'Microsoft.Web/aigateways/connections@2026-03-01-preview' = if (!empty(connectionName)) {
  parent: aiGateway
  name: connectionName
  properties: {
    connectorName: connectorName
  }
}

resource teamsConnection 'Microsoft.Web/aigateways/connections@2026-03-01-preview' = if (!empty(teamsConnectionName)) {
  parent: aiGateway
  name: teamsConnectionName
  properties: {
    connectorName: 'teams'
  }
}

resource teamsConnectionAccessPolicy 'Microsoft.Web/aigateways/connections/accessPolicies@2026-03-01-preview' = if (!empty(teamsConnectionName) && !empty(functionAppPrincipalId)) {
  parent: teamsConnection
  name: 'functionapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: functionAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

@description('The resource ID of the AI Gateway.')
output resourceId string = aiGateway.id

@description('The name of the AI Gateway.')
output name string = aiGateway.name

@description('The resource ID of the AI Gateway Connection.')
output connectionResourceId string = !empty(connectionName) ? aiGatewayConnection.id : ''

@description('The name of the AI Gateway Connection.')
output connectionName string = !empty(connectionName) ? aiGatewayConnection.name : ''

@description('The name of the Teams AI Gateway Connection.')
output teamsConnectionName string = !empty(teamsConnectionName) ? teamsConnection.name : ''

@description('The connection runtime URL for the Teams connection.')
output teamsConnectionRuntimeUrl string = !empty(teamsConnectionName) ? teamsConnection.properties.connectionRuntimeUrl : ''
