# Run the add-in with Office Add-ins Development Kit extension
The simpliest way to run this add-in project is using Office Add-ins Development Kit. Here are steps to follow:

1. **Open the Office Add-ins Development Kit**
    
    Click the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30"/> button in the side panel to open the extension.

1. **Preview Your Office Add-in (F5)**
    
    Select `Preview Your Office Add-in(F5)` to start debugging the add-in code. 
    
    <img src="./assets/devkit_preview.png" width="500"/>

    <br>After selecting the button, the extension will first check prerequites before debugging starts. Check the terminal for detailed information and guiduance to get the environment ready. After theis process, a Word/Excel/PowerPoint desktop app will launch with the add-in sample side-loaded.
    
1.  **Stop Previewing Your Office Add-in**

    After you complete the debugging process, select `Stop Previewing Your Office Add-in` to stop debugging.
    
## Common questions running an add-in
    
* To debug on Office on the web, go to [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
* To debug in Desktop (Edge Legacy), go to [Debug Edge Legacy Webview](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy)
* If you meet sideload errors, please check the following items to avoid some common errors:
    * You have installed dependencies.
    * You have closed all Word/Excel/PowerPoint apps.
    * You have stopped your last add-in previewing session.

    If you still have problems, check [troubleshoot development errors]( https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors) or [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.  


# See also
    
* **See detailed introduction to this add-in project:** `README.md` file.
* **Explore More add-in samples:** `View Samples` in `Office Add-ins Development Kit` tree view.
* **Read the documentation:** [Office add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* **Experience any problems?** [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.
* **Engage with the team to learn more about updates:** [Join the Microsoft Office Add-ins community call.](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call)