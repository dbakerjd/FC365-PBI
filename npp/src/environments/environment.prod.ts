export const environment = {
  ssoRedirectUrl: 'https://nice-cliff-0f19cff03.1.azurestaticapps.net/auth-end',
  production: true,
  isInlineApp: false,
  version: '0.6.1',
  functionAppUrl: 'https://func-fc365-test.azurewebsites.net/api/PowerBI',
  functionAppDomain: 'func-fc365-test.azurewebsites.net',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Test-NPP/user_impersonation',
  //AAD client ID
  clientID: '6c76f4df-ba13-4aca-8e16-d7c0bb9d9a51',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'
};
