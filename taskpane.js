// Office.initialize function that runs when the add-in is loaded
Office.initialize = function (reason) {
    // Ensure the DOM is loaded before we try to manipulate it
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

// Function to get details from the currently selected email
function getEmailDetails() {
    if (Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;

        // Get subject
        item.subject.getAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                $("#taskTitle").val(asyncResult.value);
            } else {
                console.error("Error getting email subject: " + asyncResult.error.message);
                updateStatus("Error getting email subject: " + asyncResult.error.message, true);
            }
        });

        // Get body (plain text)
        item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                $("#taskDescription").val(asyncResult.value);
            } else {
                console.error("Error getting email body: " + asyncResult.error.message);
                updateStatus("Error getting email body: " + asyncResult.error.message, true);
            }
        });
    } else {
        updateStatus("No email item selected or available.", true);
        console.log("No email item selected or available.");
    }
}

// Microsoft Graph API endpoint
const graphApiEndpoint = "https://graph.microsoft.com/v1.0";

// Placeholder for the access token
let graphAccessToken = null;
// Store a mapping of planId to groupId
let planToGroupMap = {};

// Function to initiate authentication and get a Graph token
async function ensureGraphToken() {
    if (graphAccessToken) {
        // TODO: Check if token is expired and refresh if necessary
        console.log("Using existing Graph token.");
        return graphAccessToken;
    }

    console.log("Attempting to get access token via Office SSO...");
    try {
        const ssoToken = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
        
        console.log("SSO token received (may need exchange for Graph token):", ssoToken);
        graphAccessToken = ssoToken; 
        updateStatus("Successfully obtained an SSO token.", false);
        return graphAccessToken;

    } catch (error) {
        console.error("SSO Error: ", error);
        graphAccessToken = null;
        if (error.code) {
            handleSSOError(error);
        } else {
            updateStatus(`Error getting SSO token: ${error.message || error}`, true);
        }
        throw error; // Re-throw to indicate failure
    }
}

function handleSSOError(error) {
    let message = `SSO Error Code: ${error.code}\nMessage: ${error.message}`;
    switch (error.code) {
        case 13001:
        case 13002:
        case 13003:
        case 13005:
        case 13010:
        case 13012:
            message += "\nConsider implementing a fallback OAuth2 flow.";
            updateStatus("SSO failed. Fallback authentication is required. (Not implemented in this MVP)", true);
            break;
        default:
            updateStatus(`SSO Error: ${error.message || error}`, true);
            break;
    }
    console.error(message);
}

async function fetchPlannerPlans() {
    console.log("Fetching Planner plans...");
    showLoading(true);
    updateStatus("Loading your Planner plans...", false);
    planToGroupMap = {}; // Reset map

    try {
        const token = await ensureGraphToken(); // Ensure we have a token
        if (!token) {
            updateStatus("Authentication token not available. Cannot fetch plans.", true);
            showLoading(false);
            return;
        }

        const response = await fetch(`${graphApiEndpoint}/me/planner/plans?$select=id,title,container`, { 
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
        const members = membersData.value.filter(member => member["@odata.type"] === "#microsoft.graph.user"); 
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
        createdTaskForAttachment = createdTask; // Save for attachment
        console.log("Task created successfully: ", createdTask);
        updateStatus(`Task "${createdTask.title}" created! ID: ${createdTask.id}. Adding details...`, false);

        if (description) {
            await updateTaskDescription(token, createdTask.id, description);
        }

        // Now, get .eml and attach it
        await getAndAttachEmailAsEml(token, createdTask.id, title);

    } catch (error) {
        console.error("Error in createTask process: ", error);
        let finalMessage = `Failed to create task: ${error.message}`;
        if (createdTaskForAttachment) { // If task was created but a later step failed
            finalMessage = `Task "${createdTaskForAttachment.title}" created, but error in subsequent step: ${error.message}`;
        }
        updateStatus(finalMessage, true);
    } finally {
        showLoading(false);
    }
}

async function updateTaskDescription(token, taskId, description) {
    console.log(`Updating description for task ID: ${taskId}`);
    updateStatus("Adding task description...", false);

    try {
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
        const etag = detailsResponse.headers.get("ETag");

        const descriptionPayload = {
            description: description
        };

        const updateResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${taskId}/details`, {
            method: "PATCH",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json",
                "Accept": "application/json",
                "If-Match": etag || "*"
            },
            body: JSON.stringify(descriptionPayload)
        });

        if (!updateResponse.ok) {
            const error = await updateResponse.json().catch(() => ({ error: { message: updateResponse.statusText } }));
            console.error("Error updating task description: ", error);
            updateStatus(`Task created, but failed to update description: ${error.error ? error.error.message : 'Unknown error'}`, true);
            return; 
        }

        console.log("Task description updated successfully.");
        updateStatus("Task created, description added. Attaching email...", false);

    } catch (error) {
        console.error("Failed to update task description: ", error);
        updateStatus(`Task created, but error updating description: ${error.message}`, true);
        throw error; // Re-throw to be caught by createTask's final catch
    }
}

async function getAndAttachEmailAsEml(token, taskId, emailSubject) {
    console.log("Attempting to get and attach .eml file...");
    updateStatus("Getting email content for attachment...", false);

    if (!Office.context.mailbox.item) {
        updateStatus("No email selected to attach.", true);
        throw new Error("No email selected.");
    }

    try {
       
(Content truncated due to size limit. Use line ranges to read in chunks)