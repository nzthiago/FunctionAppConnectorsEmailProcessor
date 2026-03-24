param name string
param location string
param tags object = {}
param connectionName string = ''
param connectorName string = ''

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

@description('The resource ID of the AI Gateway.')
output resourceId string = aiGateway.id

@description('The name of the AI Gateway.')
output name string = aiGateway.name

@description('The resource ID of the AI Gateway Connection.')
output connectionResourceId string = !empty(connectionName) ? aiGatewayConnection.id : ''

@description('The name of the AI Gateway Connection.')
output connectionName string = !empty(connectionName) ? aiGatewayConnection.name : ''
