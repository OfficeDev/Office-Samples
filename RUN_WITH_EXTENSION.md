# Run the add-in with Office Add-ins Development Kit extension
We recommend you try this sample by using the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger). The Office Add-ins Development Kit is an end-to end developer tool for building Office add-ins. It helps create, run, and debug an Office Add-in.

1. **Open the Office Add-ins Development Kit**
    
    Click the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30" alt="Office Add-ins Development Kit"/> button in the side panel to open the extension.
   
1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code.

    <img src="./assets/devkit_preview.png" width="500"/>

    <br>The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Excel desktop application launches and opens a new workbook with the sample add-in.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.
    
## Troubleshooting

If you have problems running the sample, take these steps.

- Ensure you have installed the project dependencies. Select **Check and Install Dependencies** from the Office Add-in Dev Kit extension to install these.
- Close any open instances of Excel.
- Close the previous web server started for the sample with the **Stop Previewing Your Office Add-in** Office Add-in Dev Kit extension option.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the sample on Excel on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

# See also
    
* **See detailed introduction to this add-in project:** `README.md` file.
* **Explore More add-in samples:** `View Samples` in `Office Add-ins Development Kit` tree view.
* **Read the documentation:** [Office add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* **Experience any problems?** [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.
* **Engage with the team to learn more about updates:** [Join the Microsoft Office Add-ins community call.](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call)
