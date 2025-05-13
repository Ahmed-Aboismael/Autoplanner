

### Step 8: Implement Microsoft Graph API Authentication (JavaScript/Office.js)

To interact with Microsoft Planner (and potentially OneDrive for attachments), your add-in needs to securely obtain an access token for the Microsoft Graph API. Office Add-ins provide a Single Sign-On (SSO) mechanism (`Office.auth.getAccessToken()`) which is the preferred method. If SSO fails or is unavailable (e.g., the user is using an older Office version or a personal Microsoft account that doesn\'t support SSO for the add-in), you need a fallback authentication mechanism, typically the OAuth 2.0 authorization code flow initiated in a dialog.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Key Concepts:**

1.  **Single Sign-On (SSO):**
    *   The `Office.auth.getAccessToken()` method attempts to get an access token for your add-in\'s web API (which you defined in the "Expose an API" step in Azure AD). This token can then be exchanged for a Microsoft Graph API token on your server-side (or, for simple client-side calls in some scenarios, used directly if the scopes are pre-consented and appropriate).
    *   This provides a seamless experience as the user doesn\'t need to sign in again if they are already signed into Office.
    *   Requires proper configuration in `manifest.xml` (`<WebApplicationInfo>`) and Azure AD app registration.

2.  **Fallback Authentication (OAuth 2.0 Authorization Code Flow):**
    *   If `getAccessToken()` fails (e.g., error code 13001, 13002, 13003, 13005, 13007, etc.), you must implement a fallback. This usually involves:
        *   Opening a dialog using `Office.context.ui.displayDialogAsync()`.
        *   Directing this dialog to an Azure AD authorization endpoint.
        *   User signs in and consents in the dialog.
        *   Azure AD redirects the dialog to your specified redirect URI with an authorization code.
        *   Your page at the redirect URI (or a server-side component) exchanges the code for an access token and a refresh token.
        *   The access token is then passed back to the task pane from the dialog using `Office.context.ui.messageParent()`.
    *   This is more complex to implement entirely client-side due to the need to securely handle the client secret for token exchange. Often, a minimal server-side endpoint is used for the token exchange part of the auth code flow. However, for pure client-side SPAs, PKCE (Proof Key for Code Exchange) is used with the auth code flow to avoid needing a client secret.

**Simplified SSO Approach for `taskpane.js` (Client-Side Token Handling):**

For this MVP, we\'ll focus on the `Office.auth.getAccessToken()` flow. The token returned by `getAccessToken()` is for *your add-in\'s own web API* (the resource specified in `manifest.xml`). To call Microsoft Graph, you typically need to exchange this token for a Graph token. This exchange is usually done server-side using the OAuth 2.0 on-behalf-of (OBO) flow. 

However, if the user has already consented to the Graph permissions directly (which can happen, especially in simpler scenarios or if an admin has granted tenant-wide consent), and if your Azure AD app registration is configured appropriately, sometimes you can get a token that works directly with Graph or a token that can be used in a client-side OBO-like flow if your backend is minimal. For a robust solution, a server-side OBO flow is recommended.

Let\'s start by implementing the `getAccessToken` call. We will add a function to attempt to get this token. The actual Graph API calls will use this token.

**Add Authentication Logic to `taskpane.js`:**

```javascript
// (Add this to your existing taskpane.js)

// Microsoft Graph API endpoint
const graphApiEndpoint = "https://graph.microsoft.com/v1.0";

// Placeholder for the access token
let graphAccessToken = null;

// Function to initiate authentication and get a Graph token
async function ensureGraphToken() {
    if (graphAccessToken) {
        // TODO: Check if token is expired and refresh if necessary
        console.log("Using existing Graph token.");
        return graphAccessToken;
    }

    console.log("Attempting to get access token via Office SSO...");
    try {
        // The token returned here is for your add-in\'s web API (resource defined in manifest)
        // It needs to be exchanged for a Graph token, typically server-side using OBO flow.
        // For a pure client-side add-in without a backend, this is more complex.
        // Office.auth.getAccessToken() can sometimes return a token usable with Graph if configured for it,
        // but the standard pattern is an exchange.

        // For this example, we\'ll assume a simplified scenario or a future server-side exchange.
        // The `Office.auth.getAccessToken` call itself:
        const ssoToken = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
        
        // If `forMSGraphAccess: true` is honored and configured correctly in Azure AD (scopes pre-consented or user consents),
        // this token *might* be directly usable for Graph. Otherwise, an OBO exchange is needed.
        // For now, we\'ll assign it and try to use it, but a robust app needs the OBO flow.
        console.log("SSO token received (may need exchange for Graph token):", ssoToken);
        graphAccessToken = ssoToken; // This is an oversimplification for MVP without backend.
                                     // A real app would exchange this for a Graph token.
        updateStatus("Successfully obtained an SSO token.", false);
        return graphAccessToken;

    } catch (error) {
        console.error("SSO Error: ", error);
        graphAccessToken = null;
        // Handle specific error codes for fallback
        if (error.code) {
            handleSSOError(error);
        } else {
            updateStatus(`Error getting SSO token: ${error.message || error}`, true);
        }
        throw error; // Re-throw to indicate failure
    }
}

function handleSSOError(error) {
    // Error codes for Office.auth.getAccessToken:
    // 13001: UserNotSignedIn - User is not signed in to Office. 
    // 13002: UserAborted - User aborted the sign-in or consent prompt.
    // 13003: InvalidGrant - Generally, the Office user identity does not map to a valid AAD identity, or admin consent is required.
    // 13004: InvalidResourceUrl - The resource URI in the manifest is invalid.
    // 13005: InvalidSSORequest - SSO is not supported by the Office version or platform.
    // 13006: ClientError - Generic client-side error.
    // 13007: AddinIsAlreadyRequestingToken - A previous token request is still pending.
    // 13010: UserConsentNotReceived - User did not grant consent.
    // 13012: AddinCannotConsent - Add-in is not authorized to request the specified scopes, or other consent issue.

    let message = `SSO Error Code: ${error.code}\nMessage: ${error.message}`;
    switch (error.code) {
        case 13001:
        case 13002:
        case 13003:
        case 13005:
        case 13010:
        case 13012:
            message += "\nConsider implementing a fallback OAuth2 flow.";
            // Here you would trigger your fallback authentication dialog
            // dialogFallback(); // Example function call
            updateStatus("SSO failed. Fallback authentication is required. (Not implemented in this MVP)", true);
            break;
        default:
            updateStatus(`SSO Error: ${error.message || error}`, true);
            break;
    }
    console.error(message);
}

// Modify the Office.initialize to call ensureGraphToken on load (optional, or call before first Graph API use)
Office.initialize = function (reason) {
    $(document).ready(function () {
        console.log("Outlook Planner Add-in initialized.");
        $("#createTaskButton").on("click", createTask);
        getEmailDetails();

        // Attempt to get a token silently on load, or before the first Graph call.
        // ensureGraphToken().catch(err => console.log("Initial token fetch failed, will try on action."));
    });
};

// Modify your createTask function to use the token
async function createTask() {
    updateStatus("Attempting to create task...", false);
    showLoading(true);
    console.log("Create Task button clicked.");

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Could not obtain authentication token. Please try signing in.", true);
            showLoading(false);
            return;
        }

        // If we reach here, `token` is assumed to be a Graph-compatible token (simplification for MVP)
        console.log("Using token for Graph API calls: ", token.substring(0, 20) + "...");

        const planId = $("#planSelector").val();
        const assigneeId = $("#assigneeSelector").val(); // This will be an array of IDs
        const dueDate = $("#dueDate").val();
        const title = $("#taskTitle").val();
        const description = $("#taskDescription").val();

        if (!title) {
            updateStatus("Task title is missing.", true);
            showLoading(false);
            return;
        }
        if (!planId) {
            updateStatus("Please select a Planner plan.", true);
            showLoading(false);
            return;
        }

        // Placeholder for Graph API calls
        updateStatus(`Task creation for "${title}" initiated. Planner integration pending.`, false);
        console.log("Plan ID:", planId, "Assignee:", assigneeId, "Due:", dueDate);
        
        // TODO: Implement actual Graph API calls for:
        // 1. Fetching plans (already needed for planSelector)
        // 2. Fetching plan members (already needed for assigneeSelector)
        // 3. Creating the task
        // 4. Getting .eml content
        // 5. Attaching .eml file

    } catch (error) {
        console.error("Error in createTask: ", error);
        updateStatus("Failed to create task. " + (error.message || "Please check console."), true);
    } finally {
        showLoading(false);
    }
}

// (Keep other functions like getEmailDetails, updateStatus, showLoading)
```

**Important Notes on this Authentication Code:**

*   **SSO Token vs. Graph Token:** The crucial part is that `Office.auth.getAccessToken()` provides a token for *your add-in itself*. To call Microsoft Graph, this token usually needs to be exchanged for a Graph token using the On-Behalf-Of (OBO) flow. This exchange **must happen on a server-side component** that can securely store a client secret for your Azure AD app. The `forMSGraphAccess: true` option is a newer feature that *can* simplify this if all conditions are met (permissions pre-consented, specific Office versions), but relying on it without a server-side OBO flow is not robust for all scenarios.
*   **For this MVP (without a backend):** The code above *oversimplifies* by assuming the token from `getAccessToken({ forMSGraphAccess: true })` might be directly usable. In a real-world public add-in, you would need a backend service to perform the OBO exchange.
*   **Fallback Authentication:** A complete implementation requires a full OAuth 2.0 Authorization Code Flow with PKCE (if client-side) or a standard Auth Code Flow (if using a backend for token exchange) as a fallback. This involves using `Office.context.ui.displayDialogAsync` to show a login page from Azure AD. This is a significant piece of work and is marked as "Not implemented in this MVP" in the `handleSSOError` function.
*   **jQuery:** The example uses jQuery for DOM manipulation for brevity (`$`). If your `yo office` template doesn\'t include jQuery by default (modern templates might not), you\'ll need to add it to `taskpane.html` or convert these to vanilla JavaScript DOM operations.
*   **Error Handling:** The `handleSSOError` function provides a basic structure for dealing with common SSO errors.

This step lays the groundwork for authentication. The next steps will involve using the obtained (or assumed) `graphAccessToken` to make actual calls to the Microsoft Graph API to fetch Planner plans, members, and create tasks.




### Step 9: Implement Planner Plan Selection (Graph API & JS)

Now that we have a way to get an access token, we can start making calls to the Microsoft Graph API. The first feature we'll implement using the Graph API is fetching the list of Planner plans that the user has access to and populating the "Select Planner Plan" dropdown in our task pane.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Graph API Endpoint for Planner Plans:**

*   To get the plans that the signed-in user is a member of, you can use the endpoint: `GET /me/planner/plans`
*   Each plan object returned will have an `id` and a `title`, which are what we need for our dropdown.

**Implementation Steps:**

1.  **Create a function to fetch plans:** This function will use the `graphAccessToken` obtained from `ensureGraphToken()` to make an authenticated GET request to the `/me/planner/plans` endpoint.
2.  **Populate the dropdown:** Once the plans are fetched, clear any existing options from the `#planSelector` dropdown and add new `<option>` elements for each plan.
3.  **Call this function on load:** After successful authentication (or when the add-in loads and a token is available), call this function to populate the plans.
4.  **Handle errors:** If the API call fails, display an appropriate message to the user.

**Add Planner Plan Fetching Logic to `taskpane.js`:**

```javascript
// (Add this to your existing taskpane.js, typically after the authentication functions)

async function fetchPlannerPlans() {
    console.log("Fetching Planner plans...");
    showLoading(true);
    updateStatus("Loading your Planner plans...", false);

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Authentication token not available. Cannot fetch plans.", true);
            showLoading(false);
            return;
        }

        const response = await fetch(`${graphApiEndpoint}/me/planner/plans`, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });

        if (!response.ok) {
            const error = await response.json();
            console.error("Error fetching plans: ", error);
            throw new Error(`Graph API Error ${response.status}: ${error.error ? error.error.message : 'Failed to fetch plans'}`);
        }

        const plansData = await response.json();
        const plans = plansData.value;
        console.log("Plans fetched: ", plans);

        const planSelector = $("#planSelector");
        planSelector.empty(); // Clear existing options

        if (plans && plans.length > 0) {
            planSelector.append($("<option></option>").attr("value", "").text("Select a plan..."));
            plans.forEach(plan => {
                planSelector.append($("<option></option>").attr("value", plan.id).text(plan.title));
            });
            updateStatus("Planner plans loaded.", false);
        } else {
            planSelector.append($("<option></option>").attr("value", "").text("No plans found or accessible."));
            updateStatus("No Planner plans found or you may not have access to any.", false);
        }

    } catch (error) {
        console.error("Failed to fetch Planner plans: ", error);
        updateStatus(`Error loading plans: ${error.message}`, true);
        const planSelector = $("#planSelector");
        planSelector.empty().append($("<option></option>").attr("value", "").text("Error loading plans"));
    } finally {
        showLoading(false);
    }
}

// Modify Office.initialize to call fetchPlannerPlans after ensuring a token (or on successful auth)
Office.initialize = function (reason) {
    $(document).ready(function () {
        console.log("Outlook Planner Add-in initialized.");
        $("#createTaskButton").on("click", createTask);
        $("#planSelector").on("change", handlePlanSelectionChange); // Add event listener for plan changes

        getEmailDetails();

        // Attempt to get a token and then fetch plans
        ensureGraphToken()
            .then(token => {
                if (token) {
                    fetchPlannerPlans(); // Fetch plans if token is successfully obtained
                }
            })
            .catch(err => {
                console.log("Initial token fetch or plan loading failed: ", err);
                // Status is already updated by ensureGraphToken or fetchPlannerPlans
            });
    });
};

// Add a handler for when a plan is selected (to then load assignees)
function handlePlanSelectionChange() {
    const selectedPlanId = $("#planSelector").val();
    if (selectedPlanId) {
        console.log("Plan selected: " + selectedPlanId);
        // TODO: Call a new function to fetch assignees for this planId
        // fetchPlanMembers(selectedPlanId);
        $("#assigneeSelector").empty().append($("<option></option>").attr("value", "").text("Loading members..."));
        // For now, just a placeholder for assignee logic
        // In the next step, we will implement fetchPlanMembers(selectedPlanId)
    } else {
        $("#assigneeSelector").empty().append($("<option></option>").attr("value", "").text("Select a plan first..."));
    }
}

// (Keep other functions like getEmailDetails, ensureGraphToken, handleSSOError, createTask, updateStatus, showLoading)
```

**Explanation of Changes:**

*   **`fetchPlannerPlans()` function:**
    *   Calls `ensureGraphToken()` to get the access token.
    *   Makes a `fetch` request to the `/me/planner/plans` Graph API endpoint.
    *   Includes the access token in the `Authorization` header as a Bearer token.
    *   Parses the JSON response.
    *   Clears the `#planSelector` dropdown and populates it with the fetched plans (`id` as value, `title` as text).
    *   Includes error handling and updates the status message.
*   **`Office.initialize` Modification:**
    *   After the DOM is ready and basic setup, it now calls `ensureGraphToken()`.
    *   If `ensureGraphToken()` is successful (returns a token), it then calls `fetchPlannerPlans()`.
    *   This ensures that we attempt to load plans as soon as the add-in is ready and authenticated.
*   **`handlePlanSelectionChange()` function:**
    *   A new event listener is added to the `#planSelector` dropdown.
    *   When the selected plan changes, this function will be called.
    *   It logs the selected plan ID and currently just clears the assignee dropdown with a "Loading members..." message. The actual fetching of plan members will be implemented in the next step.

**Testing this Step:**

*   To test this, you would need to sideload your add-in in Outlook.
*   When the task pane opens, it should attempt to authenticate.
*   If authentication is successful, it will then try to fetch your Planner plans and populate the dropdown.
*   Check the browser's developer console (accessible for Outlook add-ins) for any errors or log messages.

With this step, your add-in can now display a list of the user's Planner plans. The next step will be to populate the "Assign to" dropdown based on the selected plan.




### Step 10: Implement Assignee Selection (Graph API & JS)

After the user selects a Planner plan, the next step is to populate the "Assign to" dropdown with members of that selected plan. This requires another call to the Microsoft Graph API.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Graph API Endpoints for Plan Members:**

Planner plans are typically contained within a Microsoft 365 Group. To get the members of a plan, we first need the ID of the Group that owns the plan. 

1.  **Get Group ID from Plan:** The `plannerPlan` resource (which we fetched in the previous step) has a `container` property. If `container.type` is `group`, then `container.id` is the Group ID we need.
2.  **Get Group Members:** Once we have the Group ID, we can use the endpoint `GET /groups/{group-id}/members` to list the users who are members of that group. These users are the potential assignees for tasks in the plan.

**Implementation Steps:**

1.  **Modify `fetchPlannerPlans`:** When fetching plans, we need to store the `container.id` (Group ID) along with the `plan.id` and `plan.title`. We can store this as a data attribute on the `<option>` elements in the plan selector, or in a separate JavaScript object mapping plan IDs to their container Group IDs.
2.  **Create `fetchPlanMembers(groupId)` function:** This function will take a `groupId` as input, make an authenticated GET request to `/groups/{groupId}/members`, and populate the `#assigneeSelector` dropdown with the `user.displayName` and `user.id`.
3.  **Update `handlePlanSelectionChange`:** When a plan is selected, retrieve its associated Group ID and call `fetchPlanMembers(groupId)`.

**Updated `taskpane.js` Logic:**

```javascript
// (Add/Modify these parts in your existing taskpane.js)

// Store a mapping of planId to groupId
let planToGroupMap = {};

async function fetchPlannerPlans() {
    console.log("Fetching Planner plans...");
    showLoading(true);
    updateStatus("Loading your Planner plans...", false);
    planToGroupMap = {}; // Reset map

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Authentication token not available. Cannot fetch plans.", true);
            showLoading(false);
            return;
        }

        // Requesting container details along with plans
        const response = await fetch(`${graphApiEndpoint}/me/planner/plans?$select=id,title,container`, { // Added $select for container
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
            console.error("Error fetching plans: ", error);
            throw new Error(`Graph API Error ${response.status}: ${error.error ? error.error.message : 'Failed to fetch plans'}`);
        }

        const plansData = await response.json();
        const plans = plansData.value;
        console.log("Plans fetched: ", plans);

        const planSelector = $("#planSelector");
        planSelector.empty();

        if (plans && plans.length > 0) {
            planSelector.append($("<option></option>").attr("value", "").text("Select a plan..."));
            plans.forEach(plan => {
                planSelector.append($("<option></option>").attr("value", plan.id).text(plan.title));
                // Store the mapping from planId to its container's groupId if type is group
                if (plan.container && plan.container.type === 'group') {
                    planToGroupMap[plan.id] = plan.container.id;
                }
            });
            updateStatus("Planner plans loaded.", false);
        } else {
            planSelector.append($("<option></option>").attr("value", "").text("No plans found or accessible."));
            updateStatus("No Planner plans found or you may not have access to any.", false);
        }

    } catch (error) {
        console.error("Failed to fetch Planner plans: ", error);
        updateStatus(`Error loading plans: ${error.message}`, true);
        const planSelector = $("#planSelector");
        planSelector.empty().append($("<option></option>").attr("value", "").text("Error loading plans"));
    } finally {
        showLoading(false);
    }
}

async function fetchPlanMembers(groupId) {
    console.log(`Fetching members for Group ID: ${groupId}`);
    showLoading(true);
    updateStatus("Loading plan members...", false);
    const assigneeSelector = $("#assigneeSelector");
    assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Loading members..."));

    if (!groupId) {
        updateStatus("Cannot load members: Group ID not found for this plan.", true);
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Error: Group ID missing"));
        showLoading(false);
        return;
    }

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Authentication token not available. Cannot fetch members.", true);
            showLoading(false);
            return;
        }

        const response = await fetch(`${graphApiEndpoint}/groups/${groupId}/members?$select=id,displayName,userPrincipalName`, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
            console.error("Error fetching group members: ", error);
            throw new Error(`Graph API Error ${response.status}: ${error.error ? error.error.message : 'Failed to fetch members'}`);
        }

        const membersData = await response.json();
        const members = membersData.value.filter(member => member["@odata.type"] === "#microsoft.graph.user"); // Filter for users only
        console.log("Members fetched: ", members);

        assigneeSelector.empty();
        if (members && members.length > 0) {
            assigneeSelector.append($("<option></option>").attr("value", "").text("Select an assignee (optional)..."));
            members.forEach(member => {
                assigneeSelector.append($("<option></option>").attr("value", member.id).text(member.displayName || member.userPrincipalName));
            });
            updateStatus("Plan members loaded.", false);
        } else {
            assigneeSelector.append($("<option></option>").attr("value", "").text("No members found in this plan."));
            updateStatus("No members found in this plan's group.", false);
        }

    } catch (error) {
        console.error("Failed to fetch plan members: ", error);
        updateStatus(`Error loading members: ${error.message}`, true);
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Error loading members"));
    } finally {
        showLoading(false);
    }
}

function handlePlanSelectionChange() {
    const selectedPlanId = $("#planSelector").val();
    const assigneeSelector = $("#assigneeSelector");

    if (selectedPlanId) {
        console.log("Plan selected: " + selectedPlanId);
        const groupId = planToGroupMap[selectedPlanId];
        if (groupId) {
            fetchPlanMembers(groupId);
        } else {
            console.warn("No Group ID found mapped for plan ID: " + selectedPlanId);
            updateStatus("Could not determine group for this plan to fetch members.", true);
            assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Cannot load members"));
        }
    } else {
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Select a plan first..."));
    }
}

// Ensure Office.initialize and other functions (getEmailDetails, ensureGraphToken, handleSSOError, createTask, updateStatus, showLoading) are kept.
// The Office.initialize function should already be calling ensureGraphToken and fetchPlannerPlans.
// The event listener for #planSelector change is also set up in Office.initialize.

```

**Explanation of Key Changes:**

*   **`planToGroupMap`:** A global JavaScript object to store the mapping between a `plan.id` and its owning `container.id` (which is the Group ID).
*   **`fetchPlannerPlans()` Modification:**
    *   It now includes `$select=id,title,container` in the Graph API call to ensure the `container` property is returned.
    *   When populating the plan selector, it also populates `planToGroupMap` if `plan.container.type` is `'group'`.
*   **`fetchPlanMembers(groupId)` function:**
    *   This new asynchronous function takes a `groupId`.
    *   It calls `ensureGraphToken()` for authentication.
    *   Makes a GET request to `/groups/{groupId}/members` to fetch users. We also use `$select=id,displayName,userPrincipalName` to get only necessary fields.
    *   Filters the results to include only users (members can also be other groups or service principals).
    *   Populates the `#assigneeSelector` dropdown with the `displayName` (or `userPrincipalName` as fallback) and `id` of each member.
    *   Includes error handling and status updates.
*   **`handlePlanSelectionChange()` Modification:**
    *   When a plan is selected, it now looks up the `groupId` from `planToGroupMap` using the `selectedPlanId`.
    *   If a `groupId` is found, it calls `fetchPlanMembers(groupId)`.
    *   Handles cases where a `groupId` might not be found for a plan (e.g., if the plan container is not a group, though typically Planner plans are group-based).

**Important Considerations:**

*   **Permissions:** Ensure your Azure AD app registration has the `Group.Read.All` permission (delegated) for this to work. We added this during the Azure AD app registration step.
*   **Large Number of Members:** If a group has a very large number of members, you might need to implement paging for the `/groups/{groupId}/members` Graph API call. The API supports `$top` and `$skipToken` (or a `@odata.nextLink`) for this. For an MVP, fetching the first page (default is 100 members) might be sufficient, but a production add-in should handle paging.
*   **User Experience:** The `showLoading(true)` and `showLoading(false)` calls help provide feedback to the user while data is being fetched.

After implementing these changes, when a user selects a plan from the first dropdown, the add-in should attempt to fetch the members of that plan's underlying group and populate the "Assign to" dropdown.




### Step 11: Implement Task Creation Logic (Graph API & JS)

With plan and assignee selection in place, the next core feature is to actually create the task in Microsoft Planner using the Graph API. This will happen when the user clicks the "Create Task" button.

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Graph API Endpoint for Creating a Planner Task:**

*   `POST /planner/tasks`
*   The request body needs to include at least `planId` and `title`. Other properties like `assignments`, `dueDateTime`, and `details` (for description) can also be included.

**Implementation Steps:**

1.  **Enhance `createTask()` function:**
    *   Ensure an access token is available using `ensureGraphToken()`.
    *   Retrieve all necessary values from the form: `planId`, `assigneeId` (if selected), `dueDate`, `taskTitle`, `taskDescription`.
    *   Construct the JSON payload for the new task.
        *   The `assignments` property is an object where keys are user IDs and values are objects specifying the `orderHint` and `assignedBy` (which can be the current user).
        *   The `dueDateTime` needs to be in ISO 8601 format (e.g., `YYYY-MM-DDTHH:mm:ssZ`).
        *   The description will be set via a subsequent call to update the task details, as the initial task creation endpoint doesn't directly take a rich description in the same way.
    *   Make an authenticated POST request to `/planner/tasks`.
    *   Handle the response: if successful, display a success message. If not, show an error.
2.  **Update Task Details for Description:**
    *   After successfully creating a task, the response will include the `id` of the newly created task.
    *   To add the description (from the email body), you need to update the task's details using: `PATCH /planner/tasks/{task-id}/details`.
    *   The request body for this PATCH request will include the `description`.
    *   The `ETag` header from the GET task details response is required for the PATCH request to avoid conflicts. So, first GET details, then PATCH.

**Updated `createTask()` function in `taskpane.js`:**

```javascript
// (Modify the createTask function in your existing taskpane.js)

async function createTask() {
    updateStatus("Attempting to create task...", false);
    showLoading(true);
    console.log("Create Task button clicked.");

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Could not obtain authentication token. Please try signing in.", true);
            showLoading(false);
            return;
        }

        const planId = $("#planSelector").val();
        const selectedAssigneeId = $("#assigneeSelector").val(); // User ID of the selected assignee
        const dueDateValue = $("#dueDate").val(); // YYYY-MM-DD
        const title = $("#taskTitle").val();
        const description = $("#taskDescription").val();

        if (!title) {
            updateStatus("Task title is missing.", true);
            showLoading(false);
            return;
        }
        if (!planId) {
            updateStatus("Please select a Planner plan.", true);
            showLoading(false);
            return;
        }

        // Construct the task object for Graph API
        const taskPayload = {
            planId: planId,
            title: title,
            assignments: {}
        };

        if (selectedAssigneeId) {
            // Planner API expects assignments as an object where keys are user IDs
            taskPayload.assignments[selectedAssigneeId] = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !"
            };
        }

        if (dueDateValue) {
            // Convert YYYY-MM-DD to ISO 8601 format (YYYY-MM-DDTHH:mm:ssZ)
            // For simplicity, setting time to midday UTC. Adjust as needed.
            taskPayload.dueDateTime = new Date(dueDateValue + "T12:00:00Z").toISOString();
        }

        console.log("Creating task with payload: ", JSON.stringify(taskPayload));

        // Make the Graph API call to create the task
        const response = await fetch(`${graphApiEndpoint}/planner/tasks`, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json",
                "Accept": "application/json"
            },
            body: JSON.stringify(taskPayload)
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
            console.error("Error creating task: ", error);
            throw new Error(`Graph API Error ${response.status}: ${error.error ? error.error.message : 'Failed to create task'}`);
        }

        const createdTask = await response.json();
        console.log("Task created successfully: ", createdTask);
        updateStatus(`Task "${createdTask.title}" created successfully! ID: ${createdTask.id}`, false);

        // Now, update the task details with the description
        if (description) {
            await updateTaskDescription(token, createdTask.id, description);
        }
        
        // TODO: Next step - Implement .eml file attachment
        // await attachEmailAsEml(token, createdTask.id);

    } catch (error) {
        console.error("Error in createTask process: ", error);
        updateStatus(`Failed to create task: ${error.message}`, true);
    } finally {
        showLoading(false);
    }
}

async function updateTaskDescription(token, taskId, description) {
    console.log(`Updating description for task ID: ${taskId}`);
    updateStatus("Adding task description...", false);

    try {
        // 1. GET current task details to obtain the ETag
        const detailsResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${taskId}/details`, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });

        if (!detailsResponse.ok) {
            const error = await detailsResponse.json().catch(() => ({ error: { message: detailsResponse.statusText } }));
            console.error("Error getting task details for ETag: ", error);
            throw new Error(`Graph API Error ${detailsResponse.status}: ${error.error ? error.error.message : 'Failed to get task details'}`);
        }
        
        const taskDetails = await detailsResponse.json();
        const etag = detailsResponse.headers.get("ETag"); // Get ETag from headers

        if (!etag) {
            console.warn("ETag not found for task details. Description update might fail or overwrite changes.");
            // Proceeding without ETag is risky but some APIs might allow it or have workarounds.
            // For robust solution, ETag is important.
        }

        // 2. PATCH task details with the new description
        const descriptionPayload = {
            description: description
        };

        const updateResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${taskId}/details`, {
            method: "PATCH",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json",
                "Accept": "application/json",
                "If-Match": etag || "*" // Use ETag if available, otherwise wildcard (less safe)
            },
            body: JSON.stringify(descriptionPayload)
        });

        if (!updateResponse.ok) {
            const error = await updateResponse.json().catch(() => ({ error: { message: updateResponse.statusText } }));
            console.error("Error updating task description: ", error);
            // Don't throw an error that overwrites the main task creation success message, just log and perhaps a softer status update.
            updateStatus(`Task created, but failed to update description: ${error.error ? error.error.message : 'Unknown error'}`, true);
            return; // Exit without throwing to keep main success message
        }

        console.log("Task description updated successfully.");
        updateStatus("Task created and description added.", false); // Update overall status

    } catch (error) {
        console.error("Failed to update task description: ", error);
        updateStatus(`Task created, but error updating description: ${error.message}`, true);
    }
}

// (Keep other functions: Office.initialize, getEmailDetails, ensureGraphToken, handleSSOError, fetchPlannerPlans, fetchPlanMembers, handlePlanSelectionChange, updateStatus, showLoading)
```

**Explanation of Key Changes in `createTask()` and `updateTaskDescription()`:**

*   **`createTask()`:**
    *   Gathers all form values.
    *   Constructs `taskPayload` including `planId`, `title`, `assignments` (if an assignee is selected), and `dueDateTime` (if a due date is set).
        *   `assignments`: The API expects an object where keys are user IDs. The `orderHint` is a string used by Planner to determine the order of assignments; `" !"` is a common starting value.
        *   `dueDateTime`: Converted to ISO 8601 format.
    *   Makes a `POST` request to `/planner/tasks`.
    *   If successful, it logs the created task and then calls `updateTaskDescription()`.
*   **`updateTaskDescription(token, taskId, description)`:**
    *   This new asynchronous function is responsible for adding the description to the task.
    *   **GET ETag:** It first makes a `GET` request to `/planner/tasks/{taskId}/details` to retrieve the current details, primarily to get the `ETag` from the response headers. The `ETag` is crucial for preventing lost updates when making `PATCH` requests.
    *   **PATCH Description:** It then makes a `PATCH` request to the same endpoint (`/planner/tasks/{taskId}/details`), including the `description` in the payload and the `ETag` in the `If-Match` header.
    *   Handles errors specifically for the description update, trying not to overwrite the main task creation success message if only the description update fails.

**Important Considerations:**

*   **Permissions:** Ensure your Azure AD app registration has the `Tasks.ReadWrite` delegated permission for Microsoft Graph.
*   **ETags:** Using ETags with `If-Match` header for `PATCH` requests is a best practice to ensure data consistency. The code attempts to fetch and use it.
*   **Error Handling:** The functions include `try...catch` blocks to handle errors from the Graph API calls and update the UI status accordingly.
*   **Assignee Object:** The structure for `assignments` is specific. Each key is a user ID, and the value is an object defining the assignment.

With these changes, the add-in should now be able to create a Planner task with a title, due date, assignee, and description based on the email content and user input. The next step will be to implement the .eml file attachment.




### Step 12: Implement .eml File Attachment (Graph API & JS)

After creating the task and adding its description, the final core functionality for the MVP is to attach a copy of the original email as an .eml file to the Planner task. This involves several sub-steps:

1.  **Get .eml Content:** Use `Office.context.mailbox.item.getAsFileAsync()` to retrieve the raw EML content of the current email. This API returns the content as a base64 encoded string.
2.  **Upload to OneDrive:** Upload this .eml file to a location in the user's OneDrive (e.g., a dedicated folder for add-in attachments). This requires `Files.ReadWrite.All` or a more specific OneDrive permission.
    *   Endpoint: `PUT /me/drive/root:/{folder}/{filename}:/content` or `PUT /me/drive/items/{parent-item-id}/children/{filename}/content`.
    *   The body of the PUT request will be the EML content (after decoding from base64 if the API expects raw bytes, or directly if it handles base64 with appropriate content type).
3.  **Add Reference to Planner Task:** Once uploaded, get the `webUrl` of the file in OneDrive and add it as an external reference to the Planner task.
    *   Endpoint: `PATCH /planner/tasks/{task-id}/details` (same as for description).
    *   The payload will include a `references` object, where you add the new reference. The key for the reference can be the `webUrl` itself, and the value is an object describing the reference (alias, type, previewPriority, resourceUrl).

**File to Modify:**

*   `/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/src/taskpane/taskpane.js`

**Updated `createTask()` and new `getAndAttachEmailAsEml()` function in `taskpane.js`:**

```javascript
// (Modify the createTask function and add getAndAttachEmailAsEml in your existing taskpane.js)

async function createTask() {
    updateStatus("Attempting to create task...", false);
    showLoading(true);
    console.log("Create Task button clicked.");

    let createdTaskForAttachment = null; // To store created task for attachment step

    try {
        const token = await ensureGraphToken();
        if (!token) {
            updateStatus("Could not obtain authentication token. Please try signing in.", true);
            showLoading(false);
            return;
        }

        const planId = $("#planSelector").val();
        const selectedAssigneeId = $("#assigneeSelector").val();
        const dueDateValue = $("#dueDate").val();
        const title = $("#taskTitle").val();
        const description = $("#taskDescription").val();

        if (!title) {
            updateStatus("Task title is missing.", true);
            showLoading(false);
            return;
        }
        if (!planId) {
            updateStatus("Please select a Planner plan.", true);
            showLoading(false);
            return;
        }

        const taskPayload = {
            planId: planId,
            title: title,
            assignments: {}
        };

        if (selectedAssigneeId) {
            taskPayload.assignments[selectedAssigneeId] = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !"
            };
        }

        if (dueDateValue) {
            taskPayload.dueDateTime = new Date(dueDateValue + "T12:00:00Z").toISOString();
        }

        console.log("Creating task with payload: ", JSON.stringify(taskPayload));

        const response = await fetch(`${graphApiEndpoint}/planner/tasks`, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json",
                "Accept": "application/json"
            },
            body: JSON.stringify(taskPayload)
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
            console.error("Error creating task: ", error);
            throw new Error(`Graph API Error ${response.status}: ${error.error ? error.error.message : 'Failed to create task'}`);
        }

        const createdTask = await response.json();
        createdTaskForAttachment = createdTask; // Save for attachment step
        console.log("Task created successfully: ", createdTask);
        updateStatus(`Task "${createdTask.title}" created! ID: ${createdTask.id}. Adding details...`, false);

        if (description) {
            await updateTaskDescription(token, createdTask.id, description);
        }
        
        // Now, get .eml and attach it
        await getAndAttachEmailAsEml(token, createdTask.id, title); // Pass original email title for filename

    } catch (error) {
        console.error("Error in createTask process: ", error);
        let finalMessage = `Failed to create task: ${error.message}`;
        if (createdTaskForAttachment && createdTaskForAttachment.title) { // If task was created but a later step failed
            finalMessage = `Task "${createdTaskForAttachment.title}" created, but error in subsequent step: ${error.message}`;
        }
        updateStatus(finalMessage, true);
    } finally {
        showLoading(false);
    }
}

async function getAndAttachEmailAsEml(token, taskId, emailSubject) {
    console.log("Attempting to get and attach .eml file...");
    updateStatus("Getting email content for attachment...", false);

    if (!Office.context.mailbox.item) {
        updateStatus("No email selected to attach.", true);
        throw new Error("No email selected to attach."); // This will be caught by createTask
    }

    try {
        // 1. Get EML data using Office.js (returns base64 string)
        const item = Office.context.mailbox.item;
        const base64Eml = await new Promise((resolve, reject) => {
            item.getAsFileAsync({ asyncContext: null }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(asyncResult.value); // This is the base64 encoded EML content
                } else {
                    console.error("Error getting .eml file: ", asyncResult.error);
                    reject(new Error("Failed to get .eml content: " + asyncResult.error.message));
                }
            });
        });

        console.log(".eml content retrieved (first 100 chars of base64):", base64Eml.substring(0,100) + "...");
        updateStatus("Email content retrieved. Uploading as attachment...", false);

        // Convert base64 string to Blob for upload
        const byteCharacters = atob(base64Eml);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const emlBlob = new Blob([byteArray], { type: "message/rfc822" });

        // Sanitize email subject to create a valid filename
        const safeFileName = (emailSubject || "email_attachment").replace(/[^a-z0-9_.-]/gi, '_').substring(0, 100) + ".eml";

        // 2. Upload .eml to OneDrive (e.g., to /Attachments/OutlookPlannerAddin/{taskId}/{safeFileName}.eml)
        // Ensure the folder path is URL-encoded if it contains special characters (though taskId and safeFileName should be okay)
        const uploadFolderPath = `Attachments/OutlookPlannerAddin/${taskId}`;
        const fullUploadPath = `${uploadFolderPath}/${safeFileName}`;
        console.log(`Uploading .eml to OneDrive: /me/drive/root:/${fullUploadPath}:/content`);

        // First, ensure the folder exists or create it (optional, but good for organization)
        // For simplicity, this step is omitted, assuming the path can be created directly by PUT if parent exists.
        // A robust solution would check/create the folder `Attachments/OutlookPlannerAddin` and then `taskId` folder.

        const uploadResponse = await fetch(`${graphApiEndpoint}/me/drive/root:/${fullUploadPath}:/content`, {
            method: "PUT",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "message/rfc822"
            },
            body: emlBlob
        });

        if (!uploadResponse.ok) {
            const error = await uploadResponse.json().catch(() => ({ error: { message: uploadResponse.statusText } }));
            console.error("Error uploading .eml file to OneDrive: ", error);
            throw new Error(`Graph API Error ${uploadResponse.status} uploading to OneDrive: ${error.error ? error.error.message : 'Failed to upload'}`);
        }

        const uploadedFile = await uploadResponse.json();
        console.log(".eml file uploaded to OneDrive: ", uploadedFile);
        updateStatus("Email uploaded to OneDrive. Adding as task reference...", false);

        // 3. Add a reference to this uploaded file in the Planner task
        const referencePayload = {
            "@odata.type": "#microsoft.graph.plannerExternalReference",
            "alias": safeFileName,
            "previewPriority": " !", 
            "type": "Other", 
            "resourceUrl": uploadedFile.webUrl 
        };

        // Get current task details for ETag before adding reference
        const detailsResponseForRef = await fetch(`${graphApiEndpoint}/planner/tasks/${taskId}/details`, {
            method: "GET",
            headers: { "Authorization": "Bearer " + token, "Accept": "application/json" }
        });
        if (!detailsResponseForRef.ok) {
            const error = await detailsResponseForRef.json().catch(() => ({ error: { message: detailsResponseForRef.statusText } }));
            throw new Error(`Failed to get task details for ETag before adding reference: ${error.error ? error.error.message : 'Unknown error'}`);
        }
        const etagForRef = detailsResponseForRef.headers.get("ETag");

        const addReferenceResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${taskId}/details`, {
            method: "PATCH",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json",
                "Accept": "application/json",
                "If-Match": etagForRef || "*",
                "Prefer": "return=representation"
            },
            body: JSON.stringify({ 
                references: { 
                    // Ensure the key for the reference is unique, webUrl is a good candidate
                    [uploadedFile.webUrl]: referencePayload 
                }
            })
        });

        if (!addReferenceResponse.ok) {
            const error = await addReferenceResponse.json().catch(() => ({ error: { message: addReferenceResponse.statusText } }));
            console.error("Error adding .eml reference to task: ", error);
            throw new Error(`Graph API Error ${addReferenceResponse.status} adding reference: ${error.error ? error.error.message : 'Failed to add reference'}`);
        }

        console.log("EML reference added to task successfully.");
        updateStatus("Task created, description and email attachment added!", false);

    } catch (error) {
        console.error("Failed to get and attach .eml: ", error);
        // This error will be caught by the calling createTask function's catch block
        // and update the status appropriately (e.g., "Task created, but error attaching email...")
        throw error; 
    }
}

// (Keep other functions: Office.initialize, getEmailDetails, ensureGraphToken, handleSSOError, fetchPlannerPlans, fetchPlanMembers, handlePlanSelectionChange, updateTaskDescription, updateStatus, showLoading)
```

**Explanation of `getAndAttachEmailAsEml()`:**

*   **Get .eml Data:** Uses `item.getAsFileAsync()` to get the base64 encoded EML content.
*   **Convert to Blob:** Converts the base64 string to a `Blob` object with the correct MIME type (`message/rfc822`) for uploading.
*   **Sanitize Filename:** Creates a safe filename from the email subject.
*   **Upload to OneDrive:** Makes a `PUT` request to `/me/drive/root:/{path_to_file}:/content` to upload the blob. A folder structure like `Attachments/OutlookPlannerAddin/{taskId}/{filename}.eml` is suggested for organization.
    *   *Note:* For a production app, you might want to create the `Attachments` and `OutlookPlannerAddin` folders if they don't exist using Graph API calls before uploading.
*   **Add Reference to Task:** After successful upload, it gets the `webUrl` of the uploaded file. Then, it makes a `PATCH` request to `/planner/tasks/{taskId}/details` to add this `webUrl` as an external reference. The `references` property in the PATCH body is an object where keys are unique identifiers for the references (using the `webUrl` itself is a common practice) and values are `plannerExternalReference` objects.
*   **ETag for Reference:** Similar to updating the description, an ETag is fetched before patching the task details to add the reference.
*   **Error Handling:** Includes `try...catch` to manage errors during any of these steps. Errors are thrown to be caught by the main `createTask` function's error handler, which will update the UI.

**Permissions:**

*   Ensure `Files.ReadWrite.All` (or a more scoped permission like `Files.ReadWrite.AppFolder` if you restrict uploads to a specific app folder and it meets your needs) is granted in your Azure AD app registration.
*   `Tasks.ReadWrite` is still needed for updating the task with the reference.

This completes the core logic for creating a task with details and an .eml attachment. The next step in the development guide would be to consolidate error handling and provide final testing instructions.

### Step 13: Finalize Error Handling and User Feedback

Throughout the `taskpane.js` code, we've added `try...catch` blocks and calls to `updateStatus(message, isError)` and `showLoading(isLoading)`. This step is about ensuring these are comprehensive and provide clear, user-friendly feedback for various scenarios.

**Review and Enhance:**

1.  **Network Errors:** Ensure `fetch` calls handle network failures gracefully (e.g., if the user is offline).
2.  **Graph API Errors:** The current code attempts to parse error messages from Graph API responses. Ensure these are displayed clearly.
3.  **Office.js API Errors:** Errors from `Office.context.mailbox.item` calls (e.g., `getAsync`, `getAsFileAsync`) should be caught and reported.
4.  **Authentication Errors:** `handleSSOError` provides a starting point. For a production app, the fallback mechanism would need full implementation.
5.  **User Input Validation:** Basic checks (e.g., for task title, plan selection) are in place. Review if more are needed.
6.  **Clear Status Messages:** Ensure messages in `updateStatus` are informative and guide the user on what happened or what to do next.
    *   Success messages should be clear (e.g., "Task created successfully with attachment!").
    *   Error messages should be actionable if possible (e.g., "Failed to load plans. Please check your internet connection or try again.").
7.  **Loading Indicators:** `showLoading` is used. Ensure it's called appropriately before long operations and hidden afterwards, in both success and error paths (`finally` blocks are good for this).

The existing code in `taskpane.js` already incorporates many of these principles. A final review of all `catch` blocks and `updateStatus` calls would be to ensure consistency and clarity.

For example, in the `createTask` function's final `catch` block, the error message construction was improved to be more specific if the task was created but a subsequent step (like attachment) failed.

This iterative refinement of error handling and user feedback is crucial for a good user experience.




### Step 14: Testing and Sideloading the Add-in in Outlook

With the core development of the MVP complete, it's time to test it thoroughly by sideloading it into Outlook. Sideloading allows you to run and test your add-in in Outlook without publishing it to a store.

**Prerequisites for Local Testing:**

1.  **HTTPS Development Server:** Office Add-ins require that the web files (HTML, JS, CSS) are served over HTTPS. The `yo office` generator typically sets up a development server (`npm start` or `yarn start`) that uses HTTPS and handles self-signed certificates.
2.  **Outlook Client:** You need a compatible Outlook client (Outlook on Windows, Outlook on Mac, or Outlook on the web) that supports sideloading and is connected to a Microsoft 365 account where you can test Planner integration.

**How to Start the Local Development Server:**

1.  Open a terminal or command prompt.
2.  Navigate to your add-in project directory: `cd /home/ubuntu/outlook-planner-addin/OutlookPlannerAddin`
3.  Run the start command: `npm start` (or `yarn start` if you used Yarn).
    *   This command usually does a few things:
        *   Builds your add-in (transpiles JavaScript if using TypeScript, bundles files, etc.).
        *   Starts a local web server (e.g., on `https://localhost:3000`).
        *   It might also attempt to automatically sideload the add-in if Office developer tools are configured, or it will provide instructions.

**Sideloading in Outlook on the Web (OWA):**

This is often the easiest way to quickly test.

1.  Go to Outlook on the web (e.g., `https://outlook.office.com/mail/`).
2.  Open any email.
3.  Click the **"Get Add-ins"** button (often found in the top action bar of an email, or under the "..." more actions menu).
4.  In the "Add-Ins for Outlook" dialog, select **"My add-ins"** from the left-hand menu.
5.  Scroll down to the **"Custom Addins"** section.
6.  Click **"+ Add a custom add-in"** and then choose **"Add from file..."**.
7.  Browse to your project directory (`/home/ubuntu/outlook-planner-addin/OutlookPlannerAddin/`) and select the `manifest.xml` file.
8.  Click **"Open"** and then **"Install"**.
9.  Once installed, close the dialog. You should now see your add-in available when you open an email (usually under the "..." menu or as a button on the ribbon, depending on how it was configured in the manifest).

**Sideloading in Outlook Desktop (Windows/Mac):**

The process is similar but might vary slightly based on the Outlook version.

*   **Outlook on Windows (Newer versions with centralized deployment support):** Often, the "Get Add-ins" button on the Home ribbon leads to a similar dialog as OWA where you can manage "My add-ins" and add a custom add-in from its manifest file.
*   **Outlook on Windows (Older versions or different configurations):** You might need to go to `File > Info > Manage Add-ins`, which usually opens a web browser to manage add-ins, similar to the OWA experience.
*   **Outlook on Mac:** Look for the "Get Add-ins" button on the Home ribbon or in the message surface toolbar. The process to add a custom add-in from a file should be available under "My add-ins."

**Testing Checklist:**

Once sideloaded and the local server is running (`https://localhost:3000` is accessible):

1.  **Open the Add-in:** Select an email and open your "OutlookPlannerAddin" task pane.
2.  **Email Details:** Verify that the Task Title is pre-filled with the email subject and the Description with the email body.
3.  **Authentication:** The add-in should attempt to sign you in via SSO. Check for any errors or prompts.
4.  **Planner Plan Loading:** Verify that the "Select Planner Plan" dropdown populates with your accessible plans. Check for loading messages and error handling if plans don_t load.
5.  **Assignee Loading:** Select a plan. Verify that the "Assign to" dropdown populates with members of that plan. Check loading/error messages.
6.  **Due Date Picker:** Ensure the due date picker works.
7.  **Create Task (Full Flow):**
    *   Fill in all fields (or use pre-filled ones).
    *   Click "Create Task."
    *   Observe status messages for task creation, description update, and .eml attachment.
    *   **Verify in Planner:** Go to Microsoft Planner (web or Teams app) and check if the task was created in the selected plan with the correct title, assignee (if chosen), due date, description, and if the .eml file is attached as a reference and opens correctly.
8.  **Error Handling:**
    *   Try creating a task without selecting a plan.
    *   Try creating a task without a title.
    *   If possible, simulate network errors or permission issues (harder to do systematically without more tools) and observe feedback.
    *   Check the browser_s developer console (for OWA) or the add-in_s debug console (for desktop, often accessible via a right-click menu in the task pane if developer features are enabled) for any errors.
9.  **Responsiveness (if applicable):** Check if the UI looks okay if the task pane is resized.

**Troubleshooting Common Issues:**

*   **Add-in not appearing:** Double-check the `manifest.xml` for errors. Ensure the local server is running and accessible over HTTPS. Clear Outlook_s cache if needed.
*   **HTTPS/Certificate errors:** Your browser might complain about the self-signed certificate from `localhost:3000`. You usually need to accept the risk or install the certificate as trusted for the session.
*   **Authentication failures:** Review Azure AD app registration settings (Redirect URIs, API permissions, Expose an API). Check console logs for specific error codes from `Office.auth.getAccessToken()`.
*   **Graph API errors:** Check console logs for details from Graph API responses. Ensure the correct permissions are consented to.

### Step 15: Hosting the Add-in for Broader Use (Beyond Local Development)

For users other than yourself to use the add-in, or for you to use it without running a local development server, the add-in_s web files (HTML, CSS, JavaScript, images) need to be hosted on a web server that is accessible over HTTPS.

**Key Requirements for Hosting:**

1.  **HTTPS:** Office Add-ins *must* be served from an HTTPS-enabled endpoint. HTTP is not allowed.
2.  **Static File Hosting:** Since our add-in (as developed in this MVP) is purely client-side (HTML, JS, CSS), it can be hosted on any static web hosting service.
3.  **Update Manifest:** The `manifest.xml` file contains URLs pointing to your task pane, commands file, and icons (e.g., `https://localhost:3000/taskpane.html`). These URLs **must be updated** to point to your new hosting location before you distribute the manifest for others to sideload or before you submit it to AppSource.

**Hosting Options:**

1.  **Azure Static Web Apps:**
    *   A service from Microsoft for hosting static web apps with global distribution, CI/CD from GitHub/Azure DevOps, and free SSL certificates.
    *   **Steps:**
        1.  Create an Azure Static Web App resource in the Azure portal.
        2.  Connect it to a GitHub repository containing your add-in_s web files (typically the contents of the `dist` folder after running `npm run build`, or the relevant files from `src` if no build step is strictly necessary for these simple files).
        3.  Azure Static Web Apps will build and deploy your site and provide you with an HTTPS URL (e.g., `https://<your-app-name>.azurestaticapps.net`).
        4.  Update all `https://localhost:3000/...` URLs in your `manifest.xml` to point to this new base URL.

2.  **GitHub Pages:**
    *   If your project is in a public GitHub repository, you can use GitHub Pages to host static files for free.
    *   **Steps:**
        1.  Ensure your add-in_s web files are in your GitHub repository (e.g., in the root, a `/docs` folder, or a specific branch).
        2.  Go to your repository settings on GitHub, find the "Pages" section.
        3.  Configure the source for GitHub Pages. It will provide an HTTPS URL (e.g., `https://<your-username>.github.io/<repository-name>/`).
        4.  Update URLs in `manifest.xml`.

3.  **Other Cloud Providers (AWS S3, Google Cloud Storage, Netlify, Vercel, etc.):**
    *   Most cloud providers offer static website hosting services that can be configured with HTTPS.
    *   The general process involves uploading your built web files to their storage service and configuring it for web hosting with an SSL certificate.

4.  **Your Own Web Server (e.g., Nginx, Apache):**
    *   If you have your own web server, you can host the files there. Ensure it_s configured for HTTPS with a valid SSL certificate.

**Build Step (If Applicable):**

*   The `yo office` generated project might have a build command like `npm run build`. This command typically prepares your files for production (minification, bundling, copying to a `dist` folder).
*   You would host the contents of this `dist` folder.
*   For our current project structure (`OutlookPlannerAddin/src/taskpane/`), if we are not using TypeScript or complex bundling, the files in `src/taskpane` (and `assets`) are what need to be hosted. However, a build step is good practice.
    *   Run `cd /home/ubuntu/outlook-planner-addin/OutlookPlannerAddin`
    *   Run `npm run build` (if this script exists in `package.json` and is configured).
    *   The output will likely be in a `dist` folder. These are the files to deploy.

**Updating the `manifest.xml` for Hosted Deployment:**

Let_s say you host your add-in at `https://myoutlookaddin.example.com/`.

You would need to change entries in `manifest.xml` like:

*   `<IconUrl DefaultValue="https://localhost:3000/assets/icon-64.png"/>` to `<IconUrl DefaultValue="https://myoutlookaddin.example.com/assets/icon-64.png"/>`
*   `<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>` to `<SourceLocation DefaultValue="https://myoutlookaddin.example.com/taskpane.html"/>`
*   And all other URLs in `<bt:Urls>` and `<AppDomains>`.

**Distributing the Updated Manifest:**

Once the files are hosted and the manifest is updated with the production URLs, this new `manifest.xml` is what users would use to sideload the add-in, or what you would submit to Microsoft AppSource.

**Important for Azure AD App Registration:**

*   When you move from `https://localhost:3000` to a production URL (e.g., `https://myoutlookaddin.example.com`), you **must update the Redirect URIs** in your Azure AD app registration to include the new production task pane URL (e.g., `https://myoutlookaddin.example.com/taskpane.html`). Otherwise, fallback authentication (if needed) will fail.

This guide provides the foundational steps for testing and preparing your add-in for wider use. For AppSource submission, there are additional validation steps and requirements from Microsoft.

