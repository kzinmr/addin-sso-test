# addin-sso-test

This is modified sample from [Office Add-In with SSO](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/sso-quickstart).

Some steps are required.

## App Registration on Azure AD

- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2
- Edit .env file and [MSAL.js client](https://learn.microsoft.com/en-us/azure/active-directory/develop/msal-js-initializing-client-applications) settings.

Note. With Office.js, client-side authentication is required due to iframe constraint:
- [Authenticate and authorize with the Office dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api)
- [MSAL.js iframe-usage](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/8cbd07c7100f0790220e84004138ba8a76c203d4/lib/msal-browser/docs/iframe-usage.md)

