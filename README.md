## react-spgroupspanel

## Summary
This webparts lets users manage SharePoint Groups in a modern environment. The use case this webpart was written to handle, is when business administrators of the site are supposed to manage only a couple of groups within the site. Instead of giving them link to the groups list (or creating a page with links to those groups), you can use that webpart to display to the user just the groups that they'd be interested in. Also, unlike standard group management in SharePoint, using this webpart does not force the use to leave the modern pages.

![1][figure1]
![2][figure2]
![3][figure3]

## Used SharePoint Framework Version 
![drop](https://img.shields.io/badge/drop-1.6-green.svg)

## Applies to 

* [SharePoint Framework](https:/dev.office.com/sharepoint)
* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)


## Prerequisites
 User has to have permissions to actualy manage the selected groups to use the webpart.

## Solution

Solution|Author(s)
--------|---------
react-spgroupspanel| jspiew (jacek.spiewak@outlook.com)

## Version history

Version|Date|Comments
-------|----|--------
1.0|January 30, 2019|Initial release

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- in the command line run:
  - `npm install`
  - `gulp serve`

## Features
- selecting groups that are displayed in the webpart (useful for exposing just specific groups for business admins)
- editing group's owner
- editing group's members
- editing group's title
- editing group's description
- editing group's view membership and request membership settings
- editing group's request access email address
- people picker based on User Profile Service


<img src="https://telemetry.sharepointpnp.com/sp-dev-fx-webparts/samples/react-spgroupspanel" />

[figure1]: ./assets/1.png
[figure2]: ./assets/2.png
[figure3]: ./assets/3.png