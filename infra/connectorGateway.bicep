param name string
param location string
param tags object = {}
param connectionName string = ''
param connectorName string = ''
param teamsConnectionName string = ''
param functionAppPrincipalId string = ''
param tenantId string = tenant().tenantId

resource connectorGateway 'Microsoft.Web/connectorGateways@2026-05-01-preview' = {
  name: name
  location: location
  tags: tags
}

resource connectorGatewayConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = if (!empty(connectionName)) {
  parent: connectorGateway
  name: connectionName
  properties: {
    connectorName: connectorName
  }
}

resource teamsConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = if (!empty(teamsConnectionName)) {
  parent: connectorGateway
  name: teamsConnectionName
  properties: {
    connectorName: 'teams'
  }
}

resource teamsConnectionAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(teamsConnectionName) && !empty(functionAppPrincipalId)) {
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

@description('The resource ID of the Connector Gateway.')
output resourceId string = connectorGateway.id

@description('The name of the Connector Gateway.')
output name string = connectorGateway.name

@description('The resource ID of the Connector Gateway Connection.')
output connectionResourceId string = !empty(connectionName) ? connectorGatewayConnection.id : ''

@description('The name of the Connector Gateway Connection.')
output connectionName string = !empty(connectionName) ? connectorGatewayConnection.name : ''

@description('The name of the Teams Connector Gateway Connection.')
output teamsConnectionName string = !empty(teamsConnectionName) ? teamsConnection.name : ''

@description('The connection runtime URL for the Teams connection.')
output teamsConnectionRuntimeUrl string = !empty(teamsConnectionName) ? teamsConnection.properties.connectionRuntimeUrl : ''
