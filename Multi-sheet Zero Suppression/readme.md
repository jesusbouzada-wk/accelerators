## Zero Suppression

### Challenge
Customers leverage Workiva solutions to generate and file hundreds of reports. They create documents with linked tables that often end up with rows with empty data due to the entity. These empty rows are not relevant and must be hidden in each entity document.

### About this accelerator
Remove all your empty rows from all your documents just by clicking a button. 

![Before and After hidding empty rows](images/zero-suppression_before-after.jpg)

### How it works
This accelerator iterates through a list of documents. For each document, it finds existing tables. For each table, it unhides rows with data and hides those that are empty or contain only zeros.

The accelerator will report back whether hidding empty rows was completed successfully for each document and the time when that was completed. In case there is an error, the accelerator will also report the error.

### Input Sheet
The Input Sheet allows tackling multiple documents at once. Add the list of all the documents you want to hide empty rows. Each row of the input sheet will contain information for a different document containing empty rows. You provide the following data:
- Workspace ID: The ID of the workspace where the document containing empty rows is. You can find the workspace ID within Workiva url (the first long string of characters and numbers). If this is not provided, a default hardcoded workspace ID will be used. You can add the hyperlink to the workspace to the workspace id for users to easily navigate into that workspace.
- Document ID. The ID of the document containing empty rows. You can find the document ID within Workiva document url (the long string of characters and numbers that goes after "doc/"). This is mandatory. If this is not provided, the automation will fail. You can add the hyperlink to the document to the document id for users to easily navigate into that workspace.
- Suppress. Type "Yes" if you want the document to be processed. Type "No" to skip processing the document.
- client id and client secret. Provide both client id and client secret if the document requires specific id and secret. If this is not provided, a default hardcoded client id and client secret will be used.
- Output: The accelerator will report back whether hidding empty rows was completed successfully (i.e., "Done!"). In case there is an error, the accelerator will also report the error.
- Timpestamp: The accelerator will report back the time when the document was processed.

![Input Sheet](images/zero-suppression_input-sheet.jpg)

### Try it

#### Step 1: Create your own zero suppression
Download the zero suppression script and upload it into your Scripting Workspace in the Workiva Platform (i.e., the workspace where you have scripting enabled). Alternatively, copy the code and paste it into a new script into your Scripting Workspace.

#### Step 2: Fill the Input Sheet
Import the Input Sheet provided here into your Scripting Workspace and replace the demo data by your actual data. That is, list the documents that you want to have empty rows hidden and provide additional information as described above in the Input Sheet section. It is not required that your documents are in the Scripting Workspace. You can process documents that are in different workspaces.

#### Step 3: Create an Integrated Automation
Create an Integrated Automation in your Input Spreadsheet. 

Select Manual Execution trigger and Execute Script action. 

Enter the script ID of the script you created in Step 1. You can find the script ID within Workiva script url (the long string of characters and numbers that goes after "script/").

Select the Input Spreadsheet and the Input Section of the Input Spreadsheet that you currently are.

#### Step 4: Manually run the automation
You now can run the automation by clicking the Run button in the dropdown menu of the Integrated Automation you just created.

![Punch line](images/zero-suppression_punch-line.jpg)
