# Create an artificial intelligence-generated content (AIGC) add-in in Word

![A Word document that contains content generated by the add-in.](./assets/thumbnail.png)

This sample shows how to use a Word add-in to insert predefined or AI-generated content into a document.

## Features

- Uses the `insertFileFromBase64` method to import the document template.
- Uses the `insertParagraph`, `insertComment`, and `insertFootnote` methods to insert predefined content or content generated by Azure OpenAI Service.
- Uses the `insertInlinePictureFromBase64` method to insert a sample image into the document.
- Uses the Style API to customize the document.

## How to run this sample

### Prerequisites

- Download and install [Visual Studio Code](https://visualstudio.microsoft.com/downloads/).
- Install the latest version of [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) in Visual Studio Code.
- Install Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Sign in to Office with an account connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).
- Sign up for an [Azure OpenAI Service](https://learn.microsoft.com/azure/ai-services/openai/overview) account.
  
### Run the add-in using Office Add-ins Development Kit extension

We recommend you try this sample by using the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger). The Office Add-ins Development Kit is an end-to-end developer tool for building Office add-ins. It helps create, run, and debug an Office Add-in.

1. **Download the sample code**

   To download this sample code, either:
   - Open the Office Add-ins Development Kit extension and view samples in the **Sample gallery**. Select the **Create** button in the top-right corner of the sample page.
   - [Clone](https://docs.github.com/repositories/creating-and-managing-repositories/cloning-a-repository) this repository or download this sample to a folder on your computer. Then, open the folder in Visual Studio Code.

1. **Open the Office Add-ins Development Kit**

    Select the [Office Add-ins Development Kit](./assets/Icon_Office_Add-ins_Development_Kit.png) icon in the **Activity Bar** to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in (F5)** to launch the add-in and debug the code. In the dropdown menu, select **Desktop (Edge Chromium)**.

    ![The "Preview your Office Add-in" option in the Office Add-ins Development Kit extension.](./assets/devkit_preview.png)

    The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Word desktop application launches and opens a new document with the sample add-in.

1. **Stop Previewing Your Office Add-in**

    Once you finish testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the sample add-in

1. In the add-in's task pane, select one of the following options.
    - Select **Generate sample content** to insert sample content into the document.
    - Select **Create custom content** to create your own content using a template.
1. Perform the following steps to add a title, comment, and footnote citation into the document.
    1. Hover over **Add a title**, **Add a comment**, or **Add a footnote citation**.
    1. Select **Add a predefined title** or **Add a title generated by AI**. If you select **Add a title generated by AI**, you must connect to your Azure OpenAI Service account by providing your endpoint, deployment, and API key.
1. Select **Add a sample image** to insert an image into the document.
1. Select **Format the document** to customize the document style.
1. Select **Back** to return to the main page of the task pane.

## Explore sample files

These are the important files in the sample project.

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
| README.md                     
| RUN_WITH_EXTENSION.md         
| src/                          Add-ins source code
|   | taskpane/
|   |   | components/           React components
|   |   | css/                  Task pane style
|   |   | taskpane.html         Task pane entry HTML
|   |   | index.tsx             Task pane React component
| webpack.config.js             Webpack config
| tsconfig.json
```

## Troubleshooting

If you experience any issues while running the sample, follow these steps.

- Close any open instances of Word.
- Close the previous web server started for the sample using the **Stop Previewing Your Office Add-in** option in the Office Add-ins Development Kit.

If you continue to experience issues, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue).

For information on how to run the sample in Word on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on how to debug on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

Once you understand the sample, make it your own! All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you errors in the manifest syntax.

## Engage with the team

Did you experience any problems with the sample? [Create an issue]( https://github.com/OfficeDev/Office-Samples/issues/new) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The **taskpane.html** file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-localhost"><br>

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
