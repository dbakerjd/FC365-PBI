export const environment = {
  ssoRedirectUrl: 'https://nice-beach-021ec5303.azurestaticapps.net/auth-end',
  production: true,
  isInlineApp: false,
  version: '0.5.7',
  functionAppUrl: 'https://fc365.azurewebsites.net/api/PowerBI',
  functionAppDomain: 'fc365.azurewebsites.net',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365/access_as_user',
  //AAD client ID
  clientID: '9ff5f696-db6b-4373-b076-eab231d4cdcb',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'

};
