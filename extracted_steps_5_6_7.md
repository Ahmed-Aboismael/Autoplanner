### Step 5: Update the Add-in Manifest (`manifest.xml`)

After registering your application in Azure AD, you need to update your Outlook add-in's manifest file (`manifest.xml`) to include the Application (client) ID and the Application ID URI. This allows Office to identify your add-in and request tokens for the correct resource.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/manifest.xml`

**Modifications:**

1.  **Locate the `<Resources>` section:** This section is usually towards the end of the manifest file, within the `<VersionOverrides>` for `xsi:type="WebApplicationInfo"`.
2.  **Add or update the `<WebApplicationInfo>` element:**
    *   **`<Id>`:** Set this to the **Application (client) ID** you obtained from Azure AD.
    *   **`<Resource>`:** Set this to the **Application ID URI** you obtained from Azure AD (e.g., `api://YOUR_CLIENT_ID`).
    *   **`<Scopes>`:** Add the necessary Microsoft Graph API scopes that your add-in will require. For creating Planner tasks and accessing user profiles/plans, common scopes include:
        *   `profile`
        *   `openid`
        *   `User.Read` (to get user profile)
        *   `Group.ReadWrite.All` (or `Group.Read.All` if only reading plans/members, but creating tasks usually needs write access to the group backing Planner)
        *   `Tasks.ReadWrite` (Planner specific tasks permission)
        *   `Files.ReadWrite.All` (if you plan to upload .eml files to OneDrive/SharePoint)

**Example `WebApplicationInfo` section in `manifest.xml`:**

```xml
<WebApplicationInfo>
  <Id>60ca32af-6d83-4369-8a0a-dce7bb909d9d</Id> <Resource>api://60ca32af-6d83-4369-8a0a-dce7bb909d9d</Resource>
  <Scopes>
    <Scope>profile</Scope>
    <Scope>openid</Scope>
    <Scope>User.Read</Scope>
    <Scope>Group.ReadWrite.All</Scope>
    <Scope>Tasks.ReadWrite</Scope>
    <Scope>Files.ReadWrite.All</Scope>
  </Scopes>
</WebApplicationInfo>
```

**Important:**

*   Ensure this `WebApplicationInfo` element is correctly placed within the `<Resources>` tag, which itself is inside a `<VersionOverrides>` tag, typically one with `xsi:type="DesktopFormFactor"` or similar, and it should be associated with the part of your add-in that requires SSO (e.g., the task pane).
*   The exact structure might vary slightly based on your `yo office` template version. Refer to Microsoft's documentation for the most current manifest schema if unsure.
*   After saving these changes, you will need to re-sideload your add-in in Outlook for the changes to take effect.

### Step 6: Implement the Core Add-in UI (HTML & CSS)

This step involves creating the basic user interface for the add-in's task pane. This UI will include elements for selecting a Planner plan, an assignee, a due date, and fields for the task title and description (pre-filled from the email), and a button to create the task.

**Files to Modify/Create:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.html`
*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.css`

**HTML Structure (`taskpane.html`):**

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Outlook to Planner</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <!-- Fabric UI (Fluent UI) for styling - Optional, but recommended for Office look and feel -->
    <!-- You might need to install this or use a CDN -->
    <!-- <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css"/> -->

    <!-- Add your CSS styles -->
    <link rel="stylesheet" type="text/css" href="taskpane.css" />
</head>
<body class="ms-Fabric" dir="ltr">
    <div id="container">
        <h1 class="ms-fontSize-xl">Create Planner Task</h1>
        
        <div class="ms-Grid" dir="ltr">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <label for="taskTitle" class="ms-Label">Task Title (from Email Subject):</label>
                    <input type="text" id="taskTitle" class="ms-TextField-field" readonly>
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <label for="taskDescription" class="ms-Label">Description (from Email Body):</label>
                    <textarea id="taskDescription" class="ms-TextField-multiline" rows="5" readonly></textarea>
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <label for="planSelector" class="ms-Label">Select Planner Plan:</label>
                    <select id="planSelector" class="ms-Dropdown-select">
                        <option value="">Loading plans...</option>
                    </select>
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <label for="assigneeSelector" class="ms-Label">Assign To:</label>
                    <select id="assigneeSelector" class="ms-Dropdown-select">
                        <option value="">Select a plan first...</option>
                    </select>
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <label for="dueDate" class="ms-Label">Due Date:</label>
                    <input type="date" id="dueDate" class="ms-TextField-field">
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style="margin-top: 10px;">
                    <button id="createTaskButton" class="ms-Button ms-Button--primary">
                        <span class="ms-Button-label">Create Task</span>
                    </button>
                </div>
            </div>

            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12" id="status" style="margin-top: 10px;">
                    <!-- Status messages will appear here -->
                </div>
            </div>
             <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12" id="loadingIndicator" style="display:none; margin-top:10px;">
                    Loading...
                </div>
            </div>
        </div>
    </div>

    <!-- Your task pane JavaScript -->
    <script type="text/javascript" src="taskpane.js"></script>
</body>
</html>
```

**CSS Styling (`taskpane.css`):**

```css
body {
    font-family: "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
    font-size: 14px;
    margin: 20px;
    display: flex;
    flex-direction: column;
}

#container {
    width: 100%;
}

.ms-Label {
    display: block;
    margin-bottom: 5px;
    font-weight: 600;
}

.ms-TextField-field,
.ms-TextField-multiline,
.ms-Dropdown-select {
    width: calc(100% - 10px); /* Account for padding/border */
    padding: 5px;
    margin-bottom: 15px;
    border: 1px solid #c8c6c4;
    border-radius: 2px;
    box-sizing: border-box;
}

.ms-TextField-multiline {
    min-height: 80px;
    resize: vertical;
}

.ms-Button {
    padding: 8px 20px;
    border: none;
    border-radius: 2px;
    cursor: pointer;
    font-size: 14px;
    background-color: #0078d4; /* Primary button color */
    color: white;
}

.ms-Button:hover {
    background-color: #005a9e;
}

#status {
    margin-top: 15px;
    font-size: 0.9em;
}

#status.error {
    color: red;
}

#loadingIndicator {
    font-style: italic;
    color: #333;
}

/* Basic Fluent UI inspired styling - consider using Office UI Fabric/Fluent UI for a more native look */
.ms-fontSize-xl { font-size: 21px; font-weight: 600; margin-bottom: 15px; }

/* Basic Grid - for more complex layouts, consider CSS Flexbox or Grid, or Fluent UI's grid system */
.ms-Grid { display: block; }
.ms-Grid-row { margin-bottom: 10px; }
.ms-Grid-col { padding: 0 5px; }

/* For simplicity, not implementing full responsive grid here, but you would for a real add-in */
```

**Key UI Elements:**

*   Read-only fields for Task Title and Description (to be populated from email).
*   Dropdowns for Planner Plan and Assignee.
*   A date picker for the Due Date.
*   A "Create Task" button.
*   A status area to display messages to the user (e.g., success, error, loading).
*   A loading indicator.

**Notes:**

*   The HTML uses basic structure and some class names inspired by Microsoft's Fluent UI (formerly Fabric UI). For a truly native Office look and feel, you would typically integrate the Fluent UI library.
*   The CSS provides some basic styling. You can expand this to match your desired appearance.
*   The JavaScript file `taskpane.js` (to be created/modified in the next steps) will handle the logic for these UI elements.

### Step 7: Implement Email Data Extraction (Office.js)

This step focuses on using the Office JavaScript API (`Office.js`) to retrieve information from the currently selected email in Outlook. We need to get the email's subject (for the task title), body (for the task description), and a way to save the entire email as an `.eml` file for attachment.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Office.js APIs to Use:**

*   `Office.onReady()` or `Office.initialize`: To ensure the Office environment is ready before running any Office-specific code.
*   `Office.context.mailbox.item`: This object represents the currently selected email (or appointment, if that's the context).
    *   `item.subject`: To get the email subject.
    *   `item.body.getAsync(Office.CoercionType.Text, callback)`: To get the email body as plain text. You can also get it as HTML if needed.
    *   `item.getAsFileAsync({ asyncContext: null }, callback)`: This is a newer API that allows you to get the entire item (email) as an `.eml` file. This is ideal for attachments. (Note: This API might have specific availability requirements across Outlook versions/platforms. Always check compatibility or have fallbacks if targeting very old clients).

**Initial `taskpane.js` Structure and Email Data Retrieval:**

```javascript
// Office.initialize function is called when the add-in is loaded
Office.initialize = function (reason) {
    // Ensure the DOM is ready before interacting with it
    $(document).ready(function () {
        console.log("Outlook Planner Add-in initialized.");
        // Add event listener for the create task button
        $("#createTaskButton").on("click", createTask);
        // Get email details when the task pane loads
        getEmailDetails();
    });
};

// Function to get details from the current email
function getEmailDetails() {
    const item = Office.context.mailbox.item;

    // Get Subject
    if (item.subject) {
        item.subject.getAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                $("#taskTitle").val(asyncResult.value);
            } else {
                console.error("Error getting subject: " + asyncResult.error.message);
                updateStatus("Error getting email subject.", true);
            }
        });
    }

    // Get Body (as plain text)
    if (item.body) {
        item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // Truncate if too long for a simple textarea display, or handle scrolling in UI
                let bodyContent = asyncResult.value;
                if (bodyContent.length > 1000) { // Example truncation
                    bodyContent = bodyContent.substring(0, 1000) + "...";
                }
                $("#taskDescription").val(bodyContent);
            } else {
                console.error("Error getting body: " + asyncResult.error.message);
                updateStatus("Error getting email body.", true);
            }
        });
    }
}

// Function to update the status message area
function updateStatus(message, isError) {
    const statusDiv = $("#status");
    statusDiv.text(message);
    if (isError) {
        statusDiv.addClass("error");
    } else {
        statusDiv.removeClass("error");
    }
}

// Function to show/hide loading indicator
function showLoading(isLoading) {
    if (isLoading) {
        $("#loadingIndicator").show();
        $("#createTaskButton").prop("disabled", true);
    } else {
        $("#loadingIndicator").hide();
        $("#createTaskButton").prop("disabled", false);
    }
}

// Placeholder for the createTask function (will be expanded in later steps)
function createTask() {
    updateStatus("Attempting to create task...", false);
    showLoading(true);
    console.log("Create Task button clicked.");

    // Get values from the form
    const planId = $("#planSelector").val();
    const assigneeId = $("#assigneeSelector").val();
    const dueDate = $("#dueDate").val();
    const title = $("#taskTitle").val();
    const description = $("#taskDescription").val();

    if (!title) {
        updateStatus("Task title is missing.", true);
        showLoading(false);
        return;
    }

    // In later steps, we will add authentication and Graph API calls here.
    console.log("Plan ID:", planId, "Assignee:", assigneeId, "Due:", dueDate);
    updateStatus("Task creation logic not yet implemented.", false);
    // Simulate work
    setTimeout(() => {
        showLoading(false);
        updateStatus("Task creation simulated (not actually created).", false);
    }, 2000);
}

// Note: jQuery ($) is used here for simplicity. If not using jQuery,
// replace with vanilla JavaScript DOM manipulation (e.g., document.getElementById).
```

**Explanation:**

1.  **`Office.initialize`**: This is the entry point for your add-in's JavaScript logic. It ensures that the Office application is ready to host your add-in.
2.  **`getEmailDetails()`**: This function is called when the task pane loads.
    *   It accesses `Office.context.mailbox.item` to get a reference to the current email.
    *   `item.subject.getAsync()`: Asynchronously retrieves the email subject and updates the `#taskTitle` input field.
    *   `item.body.getAsync(Office.CoercionType.Text, ...)`: Asynchronously retrieves the email body as plain text and updates the `#taskDescription` textarea. The body content is truncated for display purposes in this example.
3.  **`updateStatus()` and `showLoading()`**: Helper functions to provide feedback to the user in the UI.
4.  **`createTask()`**: A placeholder function that will eventually handle the logic for creating the task in Microsoft Planner. It currently just logs to the console and simulates an action.
5.  **`.eml` File**: The logic for `item.getAsFileAsync()` to retrieve the .eml content will be integrated into the `createTask` flow in a later step when we are ready to handle file uploads/attachments to Planner via the Graph API.

**To Test This Part:**

1.  Ensure your `manifest.xml` is correctly configured.
2.  Sideload your add-in in Outlook (desktop or web).
3.  Open an email.
4.  Open your add-in's task pane.
5.  You should see the Task Title and Description fields populated with the content from the selected email.

This step provides the basic UI and the logic to extract necessary information from the selected email. The subsequent steps will focus on authentication and interacting with the Microsoft Graph API to use this data.
