<div align="center">

<img src="https://raw.githubusercontent.com/gevartrix/cai-chatbot-webpart/master/assets/logo.png"/>

</div>

# CAI Chatbot web part

[![spo](https://img.shields.io/badge/SharePoint-Online-important.svg?style=flat-square)](https://www.microsoft.com/en-ww/microsoft-365/sharepoint/collaboration) [![spfx](https://img.shields.io/badge/SPFx-1.13-green.svg?style=flat-square)](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) [![node](https://img.shields.io/badge/Node.js-LTS%20v14-brihtgreen.svg?style=flat-square)](https://nodejs.org/) [![ts](https://img.shields.io/badge/TypeScript-%5E3.7.7-brihtgreen.svg?style=flat-square)](https://www.typescriptlang.org/) [![gulp](https://img.shields.io/badge/gulp.js-%5E4.0.2-brihtgreen.svg?style=flat-square)](https://gulpjs.org/)

A client-side web part for [Microsoft SharePoint Online](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) sites that integrates [SAP Conversational AI](https://cai.tools.sap/) chatbots on a SharePoint page.

## Table of Contents

<details>
  <summary>Click to expand</summary>

  - [Summary](#summary)
    * [Demo](#demo)
    * [Project Structure](#project-structure)
    * [Modules Used](#modules-used)
  - [Installation](#installation)
    * [Disclaimer](#disclaimer)
    * [Prerequisites](#prerequisites)
    * [Creating the .sppkg file](#creating-the-sppkg-file)
  - [Contributing](#contributing)
    * [Style Guide](#style-guide)
  - [License](#license)
</details>

## Summary

This web part enables admins of a SharePoint site to embed a SAP CAI chatbot on a page. It's able to connect via any of a SAP CAI chatbot's primary connection channels — `Webchat` or `SAP CAI Web Client`. The web part works by generating an appropriate script tag for the integrated chatbot based on the chosen connection channel, and by appending said tag at the end of the page's body tag.

**WARNING. The integrated chatbot DOES NOT show up while testing this web part locally. The chatbot shows up only when configured on a SharePoint site page itself.**

The web part is also currently localized in two languages — English and Russian. We also tried to thoroughly document the source-code with JSDoc-like comments, so feel free to browse and explore it.

### Demo

![Demo of CAI Chatbot web part](./assets/demo.gif)

### Project Structure

<details>
  <summary>Click to expand</summary>

    ├── config
    │   ├── config.json
    │   ├── copy-assets.json
    │   ├── deploy-azure-storage.json
    │   ├── package-solution.json
    │   ├── serve.json
    │   └── write-manifests.json
    ├── src
    │   ├── webparts
    │   │   └── cai-chatbot
    │   │       ├── components
    │   │       │   ├── chatbot.module.scss
    │   │       │   ├── chatbot.tsx
    │   │       │   └── chatbotProps.ts
    │   │       ├── loc
    │   │       │   ├── en-us.ts
    │   │       │   ├── mystrings.d.ts
    │   │       │   └── ru-ru.ts
    │   │       ├── chatbotPropertyPane.ts
    │   │       ├── chatbotWebpart.manifest.json
    │   │       ├── chatbotWebpart.ts
    │   │       ├── chatbotWebpartProps.ts
    │   │       └── const.ts
    │   └── index.ts
    ├── .eslintignore
    ├── .eslintrc.js
    ├── .prettierignore
    ├── .prettierrc
    ├── .yo-rc.json
    ├── gulpfile.js
    ├── package.json
    ├── package-lock.json
    ├── tsconfig.json
    └── tslint.json

</details>

### Modules Used

- [SharePoint Framework (SPFx) @1.12.1](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [gulp@4.0.2](https://github.com/gulpjs/gulp)
- [react@16.9.0](https://github.com/facebook/react)
- [react-dom@16.9.0](https://github.com/facebook/react/tree/master/packages/react-dom)
- [eslint@5.16.0](https://github.com/eslint/eslint)
- [prettier@2.3.2](https://github.com/prettier/prettier)

## Installation

### Disclaimer

**This code is provided *as is* without warranty of any kind, either express or implied, including any implied warranties of fitness for a particular purpose, merchantability, or non-infringement. There is also no guarantee or warranty that raised issues will be answered or addressed in future releases.**

### Prerequisites

Before proceeding, please ensure that all of the following conditions are met:

- Your SAP Conversational AI chatbot is configured, and either the Webchat or the SAP CAI Web Client channel is created for the chatbot
- An [App Catalog](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog) has been added to your SharePoint site, and you've got the [permission](https://docs.microsoft.com/en-us/sharepoint/request-app-installation-permissions) to upload a custom web part
- [Node.js LTS v14](https://nodejs.org/download/release/v14.18.1/)[^nodeversion] is installed on your system.

### Creating the .sppkg file

> Note, that you can always download the latest precompiled package of this web part from this repository's [Releases page](https://github.com/gevartrix/cai-chatbot-webpart/releases).

In case you'd like to compile the `.sppkg` file manually, please follow these instructions.

First, clone this repository:

```bat
git clone https://github.com/gevartrix/cai-chatbot-webpart.git
```

Change directory to this project's root folder and install the required modules (it may take several minutes):

```bat
cd cai-chatbot-webpart\
npm install
```

Finally, run the following package script

```bat
npm run release
```

to generate the `cai-chatbot.sppkg` file inside the `sharepoint\solution\` folder. That script executes several `gulp` commands, please check the [package.json](package.json) file for more details.

Once the `.sppkg` file has been compiled, you can upload it to your site's App Catalog, and the web part should appear in the "Edit" mode.

To run the web part locally, execute `npm run serve`.

## Contributing

If you'd like to contribute to this little project, please follow these steps:

1. [Fork this repository](https://github.com/gevartrix/cai-chatbot-webpart/fork).
2. Create a branch: `git checkout -b feature/foo`.
3. Make your changes and commit them: `git commit -am 'Add foo feature'`.
4. Push your changes to the branch: `git push origin feature/foo`.
5. Create a new pull request.

__Pull requests are warmly welcome!__

### Style Guide

This application is configured with [ESLint](https://eslint.org/) and [Prettier](https://prettier.io), so you may want to check the configuration in the [.eslintrc.js](.eslintrc.js), [.prettier.rc](.prettierrc), [tsconfig.json](tsconfig.json) and [tslint.json](.tslint.json) files. We also try our best adhering to the best practices listed in the [Google TypeScript Style Guide](https://google.github.io/styleguide/tsguide.html).

Therefore it is recommended to lint the code before committing changes by running the following package script from the root folder:

```bat
npm run lint
```

## License

[MIT](LICENSE).

[^nodeversion]: Please refer to the [SPFx requirements page](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment#install-nodejs) to check the current latest supported version of Node.js. You may want to use a version management tool like [nvm](https://github.com/coreybutler/nvm-windows) to install the required LTS version along with your current Node.js installation.
