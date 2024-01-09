# smart-motereferat
Script for publisering og arkivering av møtereferater

# Env
Create .env with values
```bash
NODE_ENV="test"
SOURCE_AUTH_TENANT_ID="tenant id"
SOURCE_AUTH_TENANT_NAME="tenant name"
SOURCE_AUTH_CLIENT_ID="client id"
SOURCE_AUTH_PFX_PATH="path to cert.pfx"
SOURCE_AUTH_PFX_THUMBPRINT="thumbprint"
DESTINATION_AUTH_TENANT_ID="tenant id"
DESTINATION_AUTH_TENANT_NAME="tenant name"
DESTINATION_AUTH_CLIENT_ID="client id"
DESTINATION_AUTH_PFX_PATH="path to cert.pfx"
DESTINATION_AUTH_PFX_THUMBPRINT="thumbprint"
GRAPH_URL="https://graph.microsoft.com"
DISABLE_DELTA_QUERY="true"
RETRY_INTERVALS_MINUTES="1,1,1,1,1,1"
MAIL_URL="mail api url"
MAIL_KEY="mail api key"
STATISTICS_URL="stats api url"
STATISTICS_KEY="stats api key"
DELETE_FINISHED_AFTER_DAYS="30"
```
