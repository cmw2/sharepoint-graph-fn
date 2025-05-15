<!--
---
name: SharePoint Graph API Azure Function
description: This repository contains an Azure Function written in Python that connects to SharePoint document libraries via Microsoft Graph API. It uses managed identity for secure authentication and can be deployed with or without a virtual network using Azure Developer CLI (azd).
page_type: sample
languages:
- azdeveloper
- python
- bicep
products:
- azure
- azure-functions
- entra-id
- microsoft-graph
- sharepoint-online
urlFragment: sharepoint-graph-function-azd
---
-->

# SharePoint Document Library API with Azure Functions and Microsoft Graph

This repository contains a serverless API for accessing SharePoint document libraries using Azure Functions, Microsoft Graph API, and managed identity authentication. The solution is deployed to Azure using the Azure Developer CLI (`azd`) and includes security features like managed identity and optional virtual network isolation.

## Features

- **SharePoint Document Library Integration** - Connect to SharePoint Online document libraries and list all documents
- **Microsoft Graph API** - Leverage the power of Microsoft Graph to access SharePoint resources
- **Secure Authentication** - Uses Azure Managed Identity for secure, passwordless authentication
- **Secure Deployment** - Optional virtual network isolation for enhanced security
- **Serverless Architecture** - Pay-per-execution Azure Functions model for cost efficiency
- **Infrastructure as Code** - Complete Bicep templates for repeatable deployments

## Prerequisites

+ [Python 3.11](https://www.python.org/) or newer
+ [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local?pivots=programming-language-python#install-the-azure-functions-core-tools)
+ [Azure Developer CLI (AZD)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
+ [Microsoft Entra ID (Azure AD) tenant](https://learn.microsoft.com/azure/active-directory/fundamentals/active-directory-whatis) with SharePoint Online 
+ To use Visual Studio Code to run and debug locally:
  + [Visual Studio Code](https://code.visualstudio.com/)
  + [Azure Functions extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions)

## Getting Started

### Clone the repository

```shell
git clone https://github.com/cmw2/sharepoint-graph-fn.git
cd sharepoint-graph-fn
```

### Configure SharePoint access

Before running the solution, you'll need to configure the SharePoint settings. The function requires the following environment variables:

- `SHAREPOINT_TENANT_ID` - Your Microsoft Entra ID (Azure AD) tenant ID
- `SHAREPOINT_SITE_NAME` - The name of your SharePoint site 
- `SHAREPOINT_DOCUMENT_LIBRARY` - (Optional) The name of your document library (defaults to "Documents")

## Prepare your local environment

### Create local.settings.json

Create a file named `local.settings.json` in the root of your project with the following contents:

```json
{
    "IsEncrypted": false,
    "Values": {
        "AzureWebJobsStorage": "UseDevelopmentStorage=true",
        "FUNCTIONS_WORKER_RUNTIME": "python",
        "SHAREPOINT_TENANT_ID": "your-tenant-id",
        "SHAREPOINT_SITE_NAME": "your-site-name",
        "SHAREPOINT_DOCUMENT_LIBRARY": "Documents"
    }
}
```

Replace the values with your actual SharePoint settings.

### Create a virtual environment

The way that you create your virtual environment depends on your operating system.
Open the terminal, navigate to the project folder, and run these commands:

#### Linux/macOS/bash

```bash
python -m venv .venv
source .venv/bin/activate
```

#### Windows (Cmd)

```shell
py -m venv .venv
.venv\scripts\activate
```

## Run your app from the terminal

1. For local development, you'll need to authenticate with Azure using the Azure CLI:

    ```shell
    az login
    ```

1. To start the Functions host locally, run these commands in the virtual environment:

    ```shell
    pip3 install -r requirements.txt
    func start
    ```

1. Test the basic HTTP trigger (for verifying the setup):

    ```shell
    curl -i http://localhost:7071/api/httpget?name=YourName
    ```

1. Test the SharePoint document library listing endpoint:

    ```shell
    curl -i http://localhost:7071/api/sharepoint_docs_list
    ```

    This will return a JSON list of all documents in your SharePoint document library.

5. When you're done, press Ctrl+C in the terminal window to stop the host process.

6. Run `deactivate` to shut down the virtual environment.

## Run your app using Visual Studio Code

1. Open the project in Visual Studio Code:
    ```shell
    code .
    ```
2. Press **Run/Debug (F5)** to run in the debugger. Select **Debug anyway** if prompted about local emulator not running.
3. Test the SharePoint endpoint using your HTTP test tool (or browser). If you have the [RestClient](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) extension installed, you can execute requests directly from the [`test.http`](test.http) project file.

## Understanding the Code

### SharePoint Graph Client

The SharePoint integration is handled by the `SharePointGraphClient` class in [`sharepoint_graph.py`](./sharepoint_graph.py), which provides methods to:

- Authenticate with Microsoft Graph API using DefaultAzureCredential (which supports az cli login, environment variables, managed identity, etc.)
- Get SharePoint site and drive IDs
- List documents from SharePoint document libraries

```python
# Initialize SharePoint client
client = SharePointGraphClient(
    sp_tenant_id=sp_tenant_id,
    site_name=site_name,
    document_library=document_library
)

# Get all documents
docs = client.list_documents()
```

### Function App

The [`function_app.py`](./function_app.py) file contains the Azure Functions HTTP triggers:

```python
@app.route(route="sharepoint_docs_list", methods=["GET"])
def sharepoint_docs_list(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP Trigger function to list SharePoint documents.
    This function can be called anytime to get a list of documents.
    """
    # Initialize client
    client = SharePointGraphClient(
        sp_tenant_id=sp_tenant_id,
        site_name=site_name,
        document_library=document_library
    )
    
    # Get all documents
    docs = client.list_documents()
    
    # Return the documents as JSON
    return func.HttpResponse(
        json.dumps({"documents": docs}),
        mimetype="application/json"
    )
```

## Deploy to Azure

Run this command to provision the function app, with any required Azure resources, and deploy your code:

```shell
azd up
```

By default, this sample deploys with a virtual network (VNet) for enhanced security, ensuring that the function app and related resources are isolated within a private network. 
The `VNET_ENABLED` parameter controls whether a VNet is used during deployment:
- When `VNET_ENABLED` is `true` (default), the function app is deployed with a VNet for secure communication and resource isolation.
- When `VNET_ENABLED` is `false`, the function app is deployed without a VNet, allowing public access to resources.

To disable the VNet for this sample, set `VNET_ENABLED` to `false` before running `azd up`:
```bash
azd env set VNET_ENABLED false
azd up
```

You'll be prompted to supply these required deployment parameters:

| Parameter | Description |
| ---- | ---- |
| _Environment name_ | An environment name that's used to maintain a unique deployment context for your app.|
| _Azure subscription_ | Subscription in which your resources are created.|
| _Azure location_ | Azure region in which to create the resource group that contains the new Azure resources. Only regions that currently support the Flex Consumption plan are shown.|

### Configure Azure Managed Identity permissions

After deployment, you'll need to:

1. Configure the following application settings for your function app in the Azure portal:
   - `SHAREPOINT_TENANT_ID`
   - `SHAREPOINT_SITE_NAME`
   - `SHAREPOINT_DOCUMENT_LIBRARY` (optional, defaults to "Documents")

2. Grant the appropriate SharePoint permissions as detailed in the next section.

## Granting SharePoint Access

The function uses Microsoft Graph API to access SharePoint. You can choose either **managed identity** (recommended for same-tenant scenarios) or **app registration** (required for cross-tenant scenarios).

> **Note**: The SharePointGraphClient class in this project uses `DefaultAzureCredential` from the Azure Identity library, which automatically selects the appropriate credential based on the environment. When running locally, it uses your Azure CLI login. When deployed to Azure, it uses the managed identity. If you configure app registration credentials via environment variables, it will use those instead.

Both methods require two key steps:
1. Grant permission to use Microsoft Graph API with the `Sites.Selected` scope
2. Grant permission to access a specific SharePoint site

> **Note on SharePoint site IDs**: The SharePoint site ID format used by Microsoft Graph API is different from what you might see in the SharePoint URL. For Microsoft Graph, the site ID is in the format: `{tenant-name}.sharepoint.com,{site-id},{web-id}`. Use the Microsoft Graph CLI command `mgc sites list --search "site-name"` to get the full site ID in the correct format.

### Option 1: Using Managed Identity (Same Tenant)

1. **Get the Managed Identity Object ID**:
   - Go to your deployed function app in the Azure portal
   - Navigate to "Identity" in the left sidebar
   - Ensure System assigned identity is set to "On"
   - Copy the "Object (principal) ID"
   - Now got to Entra ID Enterprise Applications and find by the object ID
   - Finally copy the Application/Client ID for later use.

2. **Grant Microsoft Graph API permissions using PowerShell**:
   - Since managed identities are not visible in the App registrations section of the Azure portal, you must use PowerShell or Azure CLI to grant permissions
   - Use these PowerShell commands (requires AzureAD PowerShell module) as referenced in the [Microsoft Tech Community article](https://techcommunity.microsoft.com/blog/integrationsonazureblog/grant-graph-api-permission-to-managed-identity-object/2792127):

   ```powershell
   # Install Azure AD module if not already installed
   Install-Module AzureAD

   # Set the variables
   $TenantID="provide the tenant ID"
   $GraphAppId = "00000003-0000-0000-c000-000000000000"
   $DisplayNameOfMSI="Provide the Function App name"
   $PermissionName = "Sites.Selected"
   
   # Connect to Azure AD
   Connect-AzureAD -TenantId $TenantID
   $MSI = (Get-AzureADServicePrincipal -Filter "displayName eq '$DisplayNameOfMSI'")

   $GraphServicePrincipal = Get-AzureADServicePrincipal -Filter "appId eq '$GraphAppId'"

   $AppRole = $GraphServicePrincipal.AppRoles | Where-Object {$_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains "Application"}

   New-AzureAdServiceAppRoleAssignment -ObjectId $MSI.ObjectId -PrincipalId $MSI.ObjectId -ResourceId $GraphServicePrincipal.ObjectId -Id $AppRole.Id
   ```

3. **Grant SharePoint Site-Specific Access**:

   You'll need to use Microsoft Graph CLI to grant site-specific permissions as described in the [official documentation](https://learn.microsoft.com/en-us/graph/permissions-selected-overview?tabs=cli):

   ```bash
   # Install Microsoft Graph CLI if not already installed
   # Follow the installation instructions at: https://learn.microsoft.com/en-us/graph/cli/installation?tabs=windows
   
   # Login to Microsoft Graph
   mgc login

   # Get the site ID (you'll need this for granting permissions)
   # Format: {tenant}.sharepoint.com,{siteId},{webId}
   mgc sites list --search "your-site-name" --output json
   
   # Grant permissions to the managed identity using the site ID from above
   # and your managed identity's object ID
   mgc sites permissions create \
     --site-id "your-tenant.sharepoint.com,site-id,web-id" \
     --body "{\"roles\": [\"read\"], \"grantedTo\": {\"application\": {\"id\": \"your-managed-identity-application-id\", \"displayName\": \"Your Function App Name\"}}}"
   
   # Verify that the permission was added
   mgc sites permissions list --site-id "your-tenant.sharepoint.com,site-id,web-id"
   ```

   You can also use Microsoft Graph REST API directly with Azure CLI:  (Note this is from AI, not yet tested.)

   ```bash
   # Login to Azure
   az login

   # Get an access token for Microsoft Graph
   token=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)

   # Create the permission (replace the placeholder values)
   curl -X POST \
     -H "Authorization: Bearer $token" \
     -H "Content-Type: application/json" \
     -d "{\"roles\": [\"read\"], \"grantedTo\": {\"application\": {\"id\": \"your-managed-identity-application-id\", \"displayName\": \"Your Function App Name\"}}}" \
     "https://graph.microsoft.com/v1.0/sites/your-tenant.sharepoint.com,site-id,web-id/permissions"
   ```

### Option 2: Using App Registration (Cross-Tenant)

For cross-tenant scenarios, you'll need to create an app registration instead:

1. **Create App Registration**:
   - Go to Microsoft Entra ID → App registrations → New registration
   - Provide a name (e.g., "SharePoint Graph Function")
   - Select single tenant
   - Click "Register"

2. **Configure Authentication**:
   - Under "Certificates & secrets", create a new client secret
   - **Important**: Copy the secret value immediately as you won't be able to see it again

3. **Grant API Permissions**:
   - Go to "API permissions" → "Add a permission"
   - Select "Microsoft Graph" → "Application permissions"
   - Find and select "Sites.Selected" permission
   - Click "Add permissions"
   - Click "Grant admin consent for [your tenant]" button

4. **Grant SharePoint Site Access**:

   Use Microsoft Graph CLI to grant site-specific permissions to your app registration:

   ```bash
   # Install Microsoft Graph CLI if not already installed
   # Follow the installation instructions at: https://learn.microsoft.com/en-us/graph/cli/installation?tabs=windows
   
   # Login to Microsoft Graph
   mgc login

   # Get the site ID (you'll need this for granting permissions)
   # Format: {tenant}.sharepoint.com,{siteId},{webId}
   mgc sites list --search "your-site-name" --output json
   
   # Grant permissions to the app registration using the site ID from above
   # and your app's client ID
   mgc sites permissions create \
     --site-id "your-tenant.sharepoint.com,site-id,web-id" \
     --body "{\"roles\": [\"read\"], \"grantedTo\": {\"application\": {\"id\": \"your-app-client-id\", \"displayName\": \"SharePoint Graph Function\"}}}"
   
   # Verify that the permission was added
   mgc sites permissions list --site-id "your-tenant.sharepoint.com,site-id,web-id"
   ```

   You can also use Microsoft Graph REST API with your app's access token: (Not tested.)

   ```bash
   # Using the Microsoft Graph SDK or REST API with your app's credentials
   # First acquire a token for your app registration, then:
   curl -X POST \
     -H "Authorization: Bearer your-app-token" \
     -H "Content-Type: application/json" \
     -d "{\"roles\": [\"read\"], \"grantedTo\": {\"application\": {\"id\": \"your-app-client-id\", \"displayName\": \"SharePoint Graph Function\"}}}" \
     "https://graph.microsoft.com/v1.0/sites/your-tenant.sharepoint.com,site-id,web-id/permissions"
   ```

5. **Update Your Function App Settings**:
   - Add these additional app settings to your function app:
     - `AZURE_CLIENT_ID`: Your app registration's Application (client) ID
     - `AZURE_CLIENT_SECRET`: The secret you created
     - `AZURE_TENANT_ID`: Your tenant ID

   The `DefaultAzureCredential` class used in the SharePoint client will automatically use these settings instead of managed identity when they're available.

## Security Considerations

This solution leverages several security best practices:

1. **Managed Identity Authentication**: No credentials are stored in code or configuration (recommended approach).
2. **App Registration with Secrets**: If cross-tenant access is required, secrets are stored safely in application settings.
3. **Virtual Network Isolation** (optional): Function app and storage resources are protected inside a virtual network.
4. **Network Security Groups**: Control inbound and outbound network traffic.
5. **Least Privilege Principle**: The identity is granted only the specific permissions it needs:
   - Sites.Selected scope instead of Sites.Read.All
   - Site-specific permissions rather than tenant-wide access

## Redeploy your code

You can run the `azd up` command as many times as you need to both provision your Azure resources and deploy code updates to your function app.

## Clean up resources

When you're done working with your function app and related resources, you can use this command to delete the function app and its related resources from Azure and avoid incurring any further costs:

```shell
azd down
```

## Further Enhancements

Here are some ways you could enhance this solution:

1. Add document upload/download functionality
2. Implement search capabilities
3. Add document metadata handling
4. Include version history functionality
5. Add user-based access control
