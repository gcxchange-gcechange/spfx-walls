# spfx-walls

## Summary

Lock out certain user interface components so only Tenant Admins can access them.

## Deployment

spfx-extendlang is intended to be deployed tenant wide

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

