# React sample showing the use of sp-pnp-js with Async / Await

## Summary
This webpart demonstrates how to use [PnP JS Core](https://github.com/SharePoint/PnP-JS-Core) with Async functions into the SharePoint Framework as well as integrating [PnP JS and SPFx Logging systems](https://blog.josequinto.com/2017/04/30/how-to-integrate-pnp-js-core-and-sharepoint-framework-logging-systems/). Real example querying SharePoint library to show document sizes.

![React-sp-pnp-js-async-await](./assets/react-async-await-sp-pnp-js.png)


## Used SharePoint Framework Version 
![drop](https://img.shields.io/badge/drop-GA-green.svg)


## Applies to
* [SharePoint Framework](http://dev.office.com/sharepoint/docs/spfx/sharepoint-framework-overview)
* [Office 365 developer tenant](http://dev.office.com/sharepoint/docs/spfx/set-up-your-developer-tenant)

## Solution

Solution|Author(s)
--------|---------
react-async-await-sp-pnp-js | Jose Quinto ([@jquintozamora](https://twitter.com/jquintozamora) , [blog.josequinto.com](https://blog.josequinto.com))

## Version history

Version|Date|Comments
-------|----|--------
1.0|May 1, 2017|Initial release

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome
- clone this repo
- `$ npm i`
- `$ gulp trust-dev-cert`
- `$ gulp serve `

### Local Mode
A browser in local mode (localhost) will be opened.
https://localhost:4321/temp/workbench.html

### SharePoint Mode
If you want to try on a real environment, open:
https://your-domain.sharepoint.com/_layouts/15/workbench.aspx

## Usage
![React-sp-pnp-js-async-await-code](./assets/react-async-await-sp-pnp-js-2.png)


## Features
- [Async / Await functionality working with PnP JS sample](https://github.com/jquintozamora/spfx-react-async-await-sp-pnp-js/blob/master/src/webparts/asyncAwaitPnPJs/components/AsyncAwaitPnPJs.tsx#L93)
- React Container for the initial load. [Smart Component](https://github.com/jquintozamora/spfx-react-async-await-sp-pnp-js/blob/master/src/webparts/asyncAwaitPnPJs/components/IAsyncAwaitPnPJsState.ts)
- [Interface best practices](https://github.com/jquintozamora/spfx-react-async-await-sp-pnp-js/tree/master/src/webparts/asyncAwaitPnPJs/interfaces)
- [PnP JS and SPFx Logging systems integration](./assets/pnp-js-logging-spfx.png)
  - [SPFx Log class](https://dev.office.com/sharepoint/reference/spfx/sp-core-library/log)
  - [SPFx. Working with the Logging API](https://github.com/SharePoint/sp-dev-docs/wiki/Working-with-the-Logging-API)
  - [PnP JS Logging implementation](https://github.com/SharePoint/PnP-JS-Core/blob/master/src/utils/logging.ts)
  - [PnP JS. Working With: Logging](https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Logging)
  - [React component logging with TypeScript](https://github.com/pepaar/typescript-webpack-react-redux-boilerplate/blob/master/App/Components/BaseComponent.tsx)



<img src="https://telemetry.sharepointpnp.com/sp-dev-fx-webparts/samples/react-async-await-sp-pnp-js" />
