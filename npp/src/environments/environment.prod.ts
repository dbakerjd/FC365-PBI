export const environment = {
  ssoRedirectUrl: 'https://proud-mud-084a1d203.1.azurestaticapps.net/auth-end',
  production: true,
  isInlineApp: false,
  version: '0.6.3',
  functionAppUrl: 'https://func-fc365-pbi-dev.azurewebsites.net/api/PowerBI',
  functionAppDomain: 'func-fc365-pbi-dev.azurewebsites.net',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Dev-NPP/user_impersonation',
  //AAD client ID
  clientID: '3d632646-870b-439e-81a8-7b726b3539c8',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'
};
