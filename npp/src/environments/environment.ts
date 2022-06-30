// This file can be replaced during build by using the `fileReplacements` array.
// `ng build` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.

export const environment = {
  ssoRedirectUrl: 'https://localhost:3200/auth-end',
  production: false,
  isInlineApp: true,
  version: '0.6.1',
  functionAppUrl: 'http://localhost:7071/api/PowerBI',
  functionAppDomain: 'localhost',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Test-Inline/user_impersonation',
  //AAD client ID
  clientID: 'fa5c558e-784f-4950-af4e-7ab724b54808',
  //use common for multitenant apps, otherwise use TenantId
  authority: 'common'
};

/*
 * For easier debugging in development mode, you can import the following file
 * to ignore zone related error stack frames such as `zone.run`, `zoneDelegate.invokeTask`.
 *
 * This import should be commented out in production mode because it will have a negative impact
 * on performance if an error is thrown.
 */
// import 'zone.js/plugins/zone-error';  // Included with Angular CLI.
