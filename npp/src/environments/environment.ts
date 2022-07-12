// This file can be replaced during build by using the `fileReplacements` array.
// `ng build` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.

export const environment = {
  ssoRedirectUrl: 'https://localhost:3200/auth-end',
  production: false,
  isInlineApp: false,
  version: '0.6.3',
  functionAppUrl: 'http://localhost:7071/api/PowerBI',
  functionAppDomain: 'localhost',
  //AAD api scope to use. For multitenant must be preceded by domain, otherwise api://{clientID}/{scope}
  apiScope: 'https://janddconsulting.onmicrosoft.com/FC365-Dev-NPP/user_impersonation',
  //AAD client ID
  clientID: '3d632646-870b-439e-81a8-7b726b3539c8',
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
