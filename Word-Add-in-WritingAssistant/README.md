# Word Writing Assistant

![](./assets/sampleDemo.gif)

This add-in demonstrates Word add-in capabilities to help users check errors, improve writing, and rephrase content.

### Features
- Use insertFileFromBase64 to insert a document into the open Word document.
- Check erros, provid recommendations.

## How to run this sample

### Prerequisites
- [Node.js](https://nodejs.org) 16/18 (Tested on 16.14.0)
- [Office Add-in Debugger](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) version 0.4.0 and higher.
- Office connected to a Microsoft 365 subscription (including Office on the web). If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](
https://developer.microsoft.com/en-us/microsoft-365/dev-program);
for details, see the [FAQ](
https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).
Alternatively, you can [sign up for a 1-month free trial](
https://www.microsoft.com/en-us/microsoft-365/try?rtc=1)
or [purchase a Microsoft 365 plan](
https://www.microsoft.com/en-us/microsoft-365/buy/compare-all-microsoft-365-products).


### Run and debug the add-in
1. Open Office Add-in Debugger
<br>![](./assets/toolkit_development.png)
2. Click `Check and Install Dependencies`
3. Launch and debug
    * **For Office on Windows/macOS**, click `Preview Your Office Add-in(F5)` button on tree view and select a launch config. A Word/Excel/PowerPoint app will launch with add-in sample side-loaded. **Note:** Debugging on macOS is not supported yet.
    * To debug in Desktop (Edge Legacy), go to [Debug Edge Legacy Webview](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy)
    * **For Office on the web**: [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
4. Click `Stop Previewing Your Office Add-in` to stop debugging.


### How to use this sample add-in
1. Click the button "Import" to choose the document: assets\Company Name.docx to inserted into the current Word document.
2. Click the button "Check" to check all the potential errors.
3. Move the mouse hover the annotation to check details.
4. Select the last paragraph, click "Rewrite" to rewrite the paragraph.
5. Click the button "Ignore" to ignore all annotations.

### File structure
```
| .eslintrc.json
| .gitignore
| .vscode/
|   | extensions.json
|   | launch.json               Launch and debug configurations
|   | settings.json             
|   | tasks.json                
| assets/                       Static assets like image/gif
| babel.config.json
| manifest*.xml                 Manifest file
| package.json                  
| README.md                     Get started here
| src/                          Add-ins source code
|   | commands/
|   |   | commands.html
|   |   | commands.js
|   | taskpane/
|   |   | componets/            React components used in this sample.
|   |   |   |Annotations.tsx
|   |   |   |App.tsx
|   |   |   |FileUploader.tsx
|   |   |   |Header.tsx
|   |   |   |NewModal.tsx
|   |   | index.tsx             React component
|   |   | office-document.ts    API usages.
|   |   | taskpane.html         Taskpane entry html
| webpack.config.js             Webpack config
```

## Feedback
Did you experience any problems with the sample? [Create an issue]( https://github.com/OfficeDev/Office-Samples/issues/new) and we'll help you out.

## Copyright
Copyright (c) 2024 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
<br>**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.
<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc">

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
