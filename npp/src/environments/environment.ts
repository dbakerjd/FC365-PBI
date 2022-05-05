// This file can be replaced during build by using the `fileReplacements` array.
// `ng build` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.

export const environment = {
  ssoRedirectUrl: 'https://localhost:4200/auth-end',
  production: false,
  isInlineApp: true,
  version: '0.5.3',
  contact: {
    name: 'PBI Teams',
    email: 'rheath@jdforecasting.com'
  },
  functionAppUrl: 'http://localhost:7071/api/PowerBI',
  functionAppDomain: 'localhost',
  hashUserEmails: false,
  licensingInfo: {
    entityId : 'FC365-Enterprise-Inline-Dev',
    teamSiteDomain : 'janddconsulting.sharepoint.com'
  }
};

/*
 * For easier debugging in development mode, you can import the following file
 * to ignore zone related error stack frames such as `zone.run`, `zoneDelegate.invokeTask`.
 *
 * This import should be commented out in production mode because it will have a negative impact
 * on performance if an error is thrown.
 */
// import 'zone.js/plugins/zone-error';  // Included with Angular CLI.
