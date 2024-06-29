# SPfX User Work Anniversaries

## Summary

Displays users that have been hired for this or any selected month/s. E.g. see which users have worked for 1 ot more years and have a work anniversary.

This Web part can be added to any modern customizable SharePoint page and linked to any list with data.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Solution

| Solution    		 | Author(s)       	|
| -------------------| -----------------|
| Work Anniversaries | Nikolay Zhechev 	|

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Features

SharePoint Users List strucutre and columns:
	Id (system), Hired date (date/time), User (person or group)
	- retreived are Id, Hire date and User: User/EMail, User/Title, User/Id
	
- Graphp API is called for user avatars
	
- React is used alongise Typescript
	
- Fluent UI is used as design compoenents and layout

- PnP is utilized for API REST queries


## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
