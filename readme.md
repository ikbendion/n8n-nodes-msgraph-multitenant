# n8n-nodes-msgraph-multitenant

A community node for [n8n](https://n8n.io) that lets you call the **Microsoft Graph API across multiple Azure AD tenants** from a single workflow. Each item in a workflow can target a different tenant — useful for MSPs and multi-tenant SaaS platforms.

Forked from [`advenimus/n8n-nodes-msgraph`](https://github.com/advenimus/n8n-nodes-msgraph).

## Installation

In your n8n instance, go to **Settings > Community Nodes** and install:

```
n8n-nodes-msgraph-multitenant
```

## Prerequisites

- A **multi-tenant Azure AD app registration** ([how to create one](https://learn.microsoft.com/en-us/azure/active-directory/develop/howto-convert-app-to-be-multi-tenant))
- The app granted the required Microsoft Graph API permissions with **admin consent** from each tenant
- A **client secret** for the app

## Azure App Setup

1. Go to [portal.azure.com](https://portal.azure.com) > **Azure Active Directory** > **App registrations** > **New registration**
2. Set **Supported account types** to *Accounts in any organizational directory (Multitenant)*
3. Copy the **Application (client) ID**
4. Under **Certificates & secrets**, create a **client secret** and copy its value
5. Under **API permissions**, add the Graph permissions your workflows need (e.g. `User.Read.All`, `Mail.Send`) and grant admin consent
6. Share the consent URL with each tenant admin so they can authorize your app in their directory

## Credentials

Create a **Microsoft Graph Multi-Tenant** credential in n8n with:

| Field | Value |
|---|---|
| Client ID | Application (client) ID from the app registration |
| Client Secret | Secret value created above |

No redirect URI or OAuth flow is required — the node uses the **client credentials** grant.

## Usage

Add the **Microsoft Graph Multi-Tenant** node to your workflow and configure:

| Parameter | Description |
|---|---|
| Tenant ID | Azure AD tenant (directory) ID for the target organization |
| HTTP Method | GET, POST, PATCH, PUT, or DELETE |
| URL | Full Graph API URL, e.g. `https://graph.microsoft.com/v1.0/users` |
| Query Parameters | Optional key/value pairs appended to the URL |
| Body | JSON body for POST/PATCH/PUT requests |
| Response Format | JSON (default) or string |

The **Tenant ID** field supports n8n expressions, so you can loop over a list of tenant IDs and call Graph for each one in a single workflow execution. Tokens are cached per tenant within a single execution to avoid redundant auth requests.

**Throttling:** the node automatically retries on HTTP 429 responses, respecting the `Retry-After` header returned by Microsoft Graph (up to 5 retries).

## License

MIT — see [LICENSE](./LICENSE).
