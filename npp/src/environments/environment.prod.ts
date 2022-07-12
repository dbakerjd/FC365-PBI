export const environment = {
  ssoRedirectUrl: 'https://purple-grass-0e0729303.1.azurestaticapps.net/auth-end',
  production: true,
  isInlineApp: true,
  version: '0.6.3',
  functionAppUrl: 'https://func-fc365-pbi-dev.azurewebsites.net/api/PowerBI',
  functionAppDomain: 'func-fc365-pbi-dev.azurewebsites.net',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Dev-Inline/user_impersonation',
  //AAD client ID
  clientID: '67bd5fba-7ef2-45bc-a5df-b770dadca012',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'
};
