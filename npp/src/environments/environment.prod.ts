export const environment = {
  ssoRedirectUrl: 'https://icy-water-0ae577403.1.azurestaticapps.net/auth-end',
  production: true,
  isInlineApp: true,
  version: '0.6.4',
  functionAppUrl: 'https://func-fc365-test.azurewebsites.net/api/PowerBI',
  functionAppDomain: 'func-fc365-test.azurewebsites.net',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Test-Inline/user_impersonation',
  //AAD client ID
  clientID: 'fa5c558e-784f-4950-af4e-7ab724b54808',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'
};
