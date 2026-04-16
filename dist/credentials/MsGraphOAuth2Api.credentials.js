"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MsGraphOAuth2Api = void 0;
class MsGraphOAuth2Api {
    constructor() {
        this.name = 'msGraphOAuth2Api';
        this.displayName = 'Microsoft Graph Multi-Tenant';
        this.documentationUrl = 'https://learn.microsoft.com/en-us/graph/auth-v2-service';
        this.properties = [
            {
                displayName: 'Client ID',
                name: 'clientId',
                type: 'string',
                default: '',
                required: true,
            },
            {
                displayName: 'Client Secret',
                name: 'clientSecret',
                type: 'string',
                typeOptions: {
                    password: true,
                },
                default: '',
                required: true,
            },
        ];
    }
}
exports.MsGraphOAuth2Api = MsGraphOAuth2Api;
