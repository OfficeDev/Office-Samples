## Using chart API to analyse data

<img src="./assets/thumbnail.png" width="800">

The sample add-in demonstrates Excel add-in capablities to help users using chart API to analyse data

### Features
- Use Tabel and Range related APIs to insert sample data.
- Use Chart related APIs to analyse data like formatting chart, adding trendline and calculate top 5/10.

## How to run this sample

### Prerequisites
- [Node.js](https://nodejs.org) 16, 18, or 20 (18 is preferred) and [npm](https://www.npmjs.com/get-npm). To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- [Visual Studio Code](https://visualstudio.microsoft.com/downloads/) and [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) version 0.5.0 and higher.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](
https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](
https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details.
Alternatively, you can [sign up for a 1-month free trial](
https://www.microsoft.com/microsoft-365/try?rtc=1)
or [purchase a Microsoft 365 plan](
https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).


### Run the add-in using Office Add-ins Development Kit extension
The simpliest way to run this add-in project is using the Office Add-ins Development Kit. The [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) is an end-to end developer tool for building Office add-ins. You can use this tool to easily creating, running and debugging an Office add-in.
<br><img src="./assets/devkit_preview.png" width="800"/>

1. **Download the sample code**

    Clone or download this sample to a folder on your computer, then open the folder in Visual Studio Code.

1. **Install the Office Add-ins Development Kit**
    
    Install the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) extension in Visual Studio Code marketplace. Click the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30"/> button in the side panel to open the extension.

1. **Preview Your Office Add-in (F5)**
    
    Select `Preview Your Office Add-in(F5)` to start debugging the add-in code. 
    
    <img src="./assets/DevKit_Preview.png" width="500"/>

    <br>After selecting the button, the extension will first check prerequites before debugging starts. Check the terminal for detailed information and guiduance to get the environment ready. After theis process, a Word/Excel/PowerPoint desktop app will launch with the add-in sample side-loaded.
    
1.  **Stop Previewing Your Office Add-in**

    After you complete the debugging process, select `Stop Previewing Your Office Add-in` to stop debugging.
    
### Common questions running an add-in
    
* To debug on Office on the web, go to [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
* To debug in Desktop (Edge Legacy), go to [Debug Edge Legacy Webview](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy)
* If you meet sideload errors, please check the following items to avoid some common errors:
    * You have installed dependencies.
    * You have closed all Word/Excel/PowerPoint apps.
    * You have stopped your last add-in previewing session.

    If you still have problems, check [troubleshoot development errors]( https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors) or [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.  


### How to use this sample
1. Click the button "Add Sample data" to import some sample data.
2. Then a new table with content will be inserted into worksheet "Sample"
3. Click the button "Sales in different locations", it will insert a new chart using sales data.
4. Click the button "Reverse vertical axis order", it will reverse the vertical axis order.
5. Click the button "Customize title and legend" to change the title and legend.
6. Click the buttons under "Do analysis against source data" to have more analysis using the sample data and show the result as charts.


### Explore sample files
To explore the components of the add-in project, review the key files listed below. 
<br>You can check whether your manifest file is valid by selecting `Validate Manifest` in the `Office Add-ins Development Kit` extension tree view.
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
| manifest.xml                  Manifest file
| package.json                  
| README.md                     Get started here
| RUN_WITH_EXTENSION.md         Run the add-in with Office Add-ins Development Kit extension
| src/                          Add-ins source code
|   | commands/
|   |   | commands.html
|   |   | commands.js
|   | taskpane/
|   |   | taskpane.css          Taskpane style
|   |   | taskpane.html         Taskpane entry html
|   |   | taskpane.js           Add API calls and logic here
| webpack.config.js             Webpack config
```

### Make code changes
**Resources to learn more Office add-ins capabilities:**
* Select `View Samples` on `Office Add-ins Development Kit` tree view for real-world examples and code structures.
* [Read the documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins) of Office add-ins.

## Engage with the team
Did you experience any problems with the sample? [Create an issue]( https://github.com/OfficeDev/Office-Samples/issues/new) and we'll help you out.

Want to learn more about new features, development practices, and additional information? [Join the Microsoft Office Add-ins community call.](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call)

## Copyright
Copyright (c) 2024 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
<br>**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.
<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc">

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
