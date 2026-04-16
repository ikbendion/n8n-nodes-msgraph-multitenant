"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MsGraph = void 0;
const { NodeOperationError } = require("n8n-workflow");

class MsGraph {
  constructor() {
    this.description = {
      displayName: 'Microsoft Graph Multi-Tenant',
      name: 'msGraph',
      icon: 'file:msgraph.svg',
      group: ['transform'],
      version: 1,
      subtitle: '={{$parameter["method"] + ": " + $parameter["url"]}}',
      description: 'Consume Graph API with multi-tenant support',
      defaults: { name: 'Microsoft Graph Multi-Tenant' },
      inputs: ['main'],
      outputs: ['main'],
      credentials: [{ name: 'msGraphOAuth2Api', required: true }],
      properties: [
        { displayName: 'Tenant ID', name: 'tenantId', type: 'string', default: '', placeholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx', description: 'Azure AD Tenant (Directory) ID', required: true },
        { displayName: 'HTTP Method', name: 'method', type: 'options', options: ['GET','POST','PATCH','PUT','DELETE'].map(m => ({ name: m, value: m })), default: 'GET' },
        { displayName: 'URL', name: 'url', type: 'string', default: 'https://graph.microsoft.com/v1.0/me', placeholder: 'https://graph.microsoft.com/v1.0/users', description: 'Full Graph URL', required: true },
        { displayName: 'Query Parameters', name: 'queryParameters', type: 'fixedCollection', placeholder: 'Add Parameter', typeOptions: { multipleValues: true }, options: [{ name: 'parameter', displayName: 'Parameter', values: [ { displayName: 'Name', name: 'name', type: 'string', default: '' }, { displayName: 'Value', name: 'value', type: 'string', default: '' } ] }], default: {} },
        { displayName: 'Body', name: 'body', type: 'json', displayOptions: { show: { method: ['POST','PATCH','PUT'] } }, default: '', description: 'JSON body' },
        { displayName: 'Response Format', name: 'responseFormat', type: 'options', options: [ { name: 'JSON', value: 'json' }, { name: 'String', value: 'string' } ], default: 'json' },
      ],
    };
  }

  async execute() {
    const items = this.getInputData();
    const returnItems = [];
    const tokenCache = {};

    const oauthCreds = await this.getCredentials('msGraphOAuth2Api');
    const clientId = oauthCreds.clientId || oauthCreds.client_id;
    const clientSecret = oauthCreds.clientSecret || oauthCreds.client_secret;

    // Shared retry logic for both token fetches and Graph API calls (429/503)
    const requestWithRetry = async (options, maxRetries = 5) => {
      let retryCount = 0;
      while (true) {
        try {
          return await this.helpers.request(options);
        } catch (error) {
          if ((error.statusCode === 429 || error.statusCode === 503) && retryCount < maxRetries) {
            retryCount++;
            const retryAfterHeader = parseInt(error.response?.headers?.['retry-after'] || '0', 10);
            // Respect Retry-After header when present; otherwise exponential backoff (2s, 4s, 8s… capped at 60s)
            const delay = retryAfterHeader > 0 ? retryAfterHeader : Math.min(2 * Math.pow(2, retryCount - 1), 60);
            await new Promise(resolve => setTimeout(resolve, delay * 1000));
            continue;
          }
          throw error;
        }
      }
    };

    for (let i = 0; i < items.length; i++) {
      try {
        const tenantId = this.getNodeParameter('tenantId', i);

        // Fetch or reuse token; re-fetch if within 60s of expiry
        const now = Date.now();
        const cached = tokenCache[tenantId];
        let accessToken = cached && cached.expiresAt > now + 60000 ? cached.token : null;
        if (!accessToken) {
          const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/token`;
          const params = new URLSearchParams();
          params.append('grant_type', 'client_credentials');
          params.append('client_id', clientId);
          params.append('client_secret', clientSecret);
          params.append('resource', 'https://graph.microsoft.com');

          const tokenResponse = await requestWithRetry({
            method: 'POST',
            url: tokenUrl,
            body: params.toString(),
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            json: true,
          });

          if (!tokenResponse.access_token) {
            throw new Error(`Failed to retrieve access token for tenant ${tenantId}`);
          }
          accessToken = tokenResponse.access_token;
          const expiresIn = parseInt(tokenResponse.expires_in || '3600', 10);
          tokenCache[tenantId] = { token: accessToken, expiresAt: now + expiresIn * 1000 };
        }

        // Build request parameters
        const method = this.getNodeParameter('method', i);
        const url = this.getNodeParameter('url', i);
        const qsParams = this.getNodeParameter('queryParameters.parameter', i, []);
        const qs = qsParams.reduce((obj, p) => { obj[p.name] = p.value; return obj; }, {});

        let body;
        if (['POST','PATCH','PUT'].includes(method)) {
          body = this.getNodeParameter('body', i, {});
          if (typeof body === 'string' && body.trim()) {
            try { body = JSON.parse(body); } catch {
              throw new NodeOperationError(this.getNode(), 'Body must be valid JSON');
            }
          }
        }

        const headers = { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' };
        if (body) headers['Content-Type'] = 'application/json';

        const responseFormat = this.getNodeParameter('responseFormat', i, 'json');
        const requestOptions = { method, url, headers, qs, body, json: responseFormat === 'json' };

        const response = await requestWithRetry(requestOptions);

        let output = response;
        if (responseFormat === 'string') output = typeof response === 'object' ? JSON.stringify(response) : String(response);

        returnItems.push({ json: output });
      } catch (err) {
        if (this.continueOnFail()) {
          returnItems.push({ json: { error: err.message } });
        } else {
          throw err;
        }
      }
    }

    return this.prepareOutputData(returnItems);
  }
}

exports.MsGraph = MsGraph;
