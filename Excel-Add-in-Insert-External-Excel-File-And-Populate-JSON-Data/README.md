![](./assets/sampleDemo.gif)

This sample shows how to insert an existing template from an external Excel file into the currently open Excel file. Then it retrieves data from a JSON web service and populates the template for the customer.

### Features
- Use insertWorksheetsFromBase64 to insert a worksheet from another Excel file into the open Excel file.
- Get JSON data and add it to the worksheet.

## Steps & How to Run

### Prerequisites
- [Node.js](https://nodejs.org) 16/18 (Tested on 16.14.0)

### Run & Debug
1. Open M365/Teams Toolkit
<br>![](./assets/toolkit_development.png)
2. Click `Check and Install Dependencies`
3. Launch and debug
    * **For Office on Windows/macOS**, click `Preview Your Office Add-in(F5)` button on tree view and select a launch config. A Word/Excel/PowerPoint app will launch with add-in sample side-loaded. **Note:** Debugging on macOS is not supported yet.
    * **For Office on the web**: [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
4. lick `Stop Preview Your Office Add-in` to stop debugging.


### How to use this sample
1. Click the button "Choose File" to choose the Excel file you want to inserted into the worksheet. You can also use the provided file nameds as "SalesTemplate.xlsx" to have a test.
2. Then, the provided Excel file will be inserted into current worksheet.
3. What's more, we will also automate appendix an worksheet using the provided "JSONData.json"


## Feedback
Did you experience any problems with the sample? [Create an issue]( https://github.com/OfficeDev/Office-Samples/issues/new) and we'll help you out.

## Copyright
Copyright (c) 2024 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
<br>**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.
<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc">

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
