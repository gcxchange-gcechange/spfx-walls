# spfx-walls

## Summary

Lock out certain user interface components so only Tenant Admins or special roles can access them.

**_Group id values based on gcxchange. To use use on a different tenant, please update the group id values adminGroupIds either through SharePoint admin (deployed) or serve.json (development)_**

## Deployment

spfx-walls is intended to be deployed tenant wide

## Required API access

These Graph permissions are required for spfx-walls to run properly
- User.ReadBasic.All
- Group.Read.All
- RoleManagement.Read.Directory

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)


## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **Update pageUrl in ./config/serve.json to your target page**
  - **gulp serve**

## Adding and editing properties through SharePoint

This extention pulls values from the extension properties defined on SharePoint. When deployed there are already some default values provided. You can edit these from the app catalog's tenant wide section. The properties follow JSON formatting, and each property is a string that needs to start and end in double quotations. The properties this extension needs to fuction properly are:

- adminGroupIds: A list of comma seperated GUIDs that represent what's considered an administrative level group.
- adminSelectorsCSS: A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- ownerSelectorsCSS:	A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- memberSelectorsCSS:  A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- adminRedirects: A list comma seperated strings. The URL is checked if it **contains** any of these it will redirect amins away from the page.
- ownerRedirects:	A list comma seperated strings. The URL is checked if it **contains** any of these it will redirect owners away from the page.
- memberRedirects: A list comma seperated strings. The URL is checked if it **contains** any of these it will redirect members away from the page.
- redirectLandingPage: The page to redirect to. If blank it will redirect to the home page.
- logging: This turns logging to the web console on or off. A value of "true" is on, anything else is considered off.
