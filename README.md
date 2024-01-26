# SPFX Walls Extension

## Summary
 This extention lock out certain user interface components so only Tenant Admins or special roles can access them. It pulls values from the extension properties defined on SharePoint. When deployed there are already some default values provided. You can edit these from the app catalog's tenant wide section. The properties follow JSON formatting, and each property is a string that needs to start and end in double quotations. This extension is intended to be deployed tenant wide.
 
 The properties this extension needs to fuction properly are:
- adminGroupIds: A list of comma seperated GUIDs that represent what's considered an administrative level group.
- adminSelectorsCSS: A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- ownerSelectorsCSS:	A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- memberSelectorsCSS:  A list of comma seperated CSS selectors to be hidden and removed. Any valid CSS selector will work. Avoid using commas in your selectors!
- adminRedirects: A list comma seperated strings. **The URL is checked if it contains** any of these it will redirect amins away from the page.
- ownerRedirects:	A list comma seperated strings. **The URL is checked if it contains** any of these it will redirect owners away from the page.
- memberRedirects: A list comma seperated strings. **The URL is checked if it contains** any of these it will redirect members away from the page.
- redirectLandingPage: The page to redirect to. If blank it will redirect to the home page.
- logging: This turns logging to the web console on or off. A value of "true" is on, anything else is considered off.

## Prerequisites
None.

## API permission
These Graph permissions are required for spfx-walls to run properly
- User.ReadBasic.All
- Group.Read.All
- RoleManagement.Read.Directory

## Version 
![SPFX](https://img.shields.io/badge/SPFX-1.17.4-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-v16.3+-green.svg)


## Applies to
- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Version history
Version|Date|Comments
-------|----|--------
1.0.0  | Jan 24, 2020 | Initial release
1.0.1  | Jan 15, 2024 | Upgraded to SPFX 1.17.4


## Minimal Path to Awesome
- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - **npm install**
- To debug
  - **in the command-line run:**
    - **gulp clean**
    - **gulp serve**
- To deploy: 
  - **in the command-line run:**
    - **gulp clean**
    - **gulp bundle --ship**
    - **gulp package-solution --ship**
    - **Upload the extension from `\sharepoint\solution` to your tenant's app store**
- To add or modify extension properties
  - **Go to Modern Appcatalog**
  - **Click ...More features in the left side**
  - **Open the tenant wide Extension**
  - **Edit the extension's properties**
- Approve the web API permissions
## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**