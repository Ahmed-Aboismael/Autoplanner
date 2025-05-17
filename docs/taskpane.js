// Enhanced debug version of taskpane.js with additional logging

// Office.initialize function that runs when the add-in is loaded
Office.initialize = function (reason) {
    console.log("Office.initialize started with reason:", reason);
    
    // Check if jQuery is available
    if (typeof $ === 'undefined') {
        console.error("jQuery ($) is not defined! Make sure jQuery is loaded before taskpane.js");
        alert("Error: jQuery not loaded. Please check the console for details.");
        return;
    }
    
    console.log("jQuery is available, continuing initialization...");
    
    // Ensure the DOM is loaded before we try to manipulate it
    $(document).ready(function () {
        console.log("DOM ready, setting up event handlers...");
        
        // Log initialization
        console.log("Outlook Planner Add-in initialized.");
        
        // Set up event handlers
        $("#createTaskButton").on("click", function() {
            console.log("Create Task button clicked");
            createTask();
        });
        
        $("#planSelector").on("change", function() {
            console.log("Plan selection changed to:", $(this).val());
            handlePlanSelectionChange();
        });
        
        // Attempt to get email details
        console.log("Attempting to get email details...");
        getEmailDetails();
        
        // Attempt to get a token and then fetch plans
        console.log("Starting authentication process...");
        ensureGraphToken()
            .then(token => {
                console.log("Token obtained:", token ? "Yes (token available)" : "No (token is null or empty)");
                if (token) {
                    console.log("Fetching Planner plans...");
                    fetchPlannerPlans();
                } else {
                    console.warn("No token available, cannot fetch plans");
                    updateStatus("Authentication failed. Cannot fetch plans.", true);
                }
            })
            .catch(err => {
                console.error("Initial token fetch or plan loading failed: ", err);
                updateStatus("Error during authentication: " + (err.message || err), true);
            });
    });
};

// Function to get details from the currently selected email
function getEmailDetails() {
    console.log("getEmailDetails() called");
    
    if (!Office || !Office.context || !Office.context.mailbox) {
        console.error("Office.context.mailbox is not available!");
        updateStatus("Error: Office mailbox context not available", true);
        return;
    }
    
    if (Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        console.log("Email item found, attempting to get subject...");
        
        // Get subject
        item.subject.getAsync(function (asyncResult) {
            console.log("Subject getAsync result:", asyncResult);
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully got email subject:", asyncResult.value);
                $("#taskTitle").val(asyncResult.value);
            } else {
                console.error("Error getting email subject: ", asyncResult.error);
                updateStatus("Error getting email subject: " + asyncResult.error.message, true);
            }
        });
        
        // Get body (plain text)
        console.log("Attempting to get email body...");
        item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            console.log("Body getAsync result:", asyncResult);
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully got email body (length: " + asyncResult.value.length + " chars)");
                $("#taskDescription").val(asyncResult.value);
            } else {
                console.error("Error getting email body: ", asyncResult.error);
                updateStatus("Error getting email body: " + asyncResult.error.message, true);
            }
        });
    } else {
        console.warn("No email item selected or available.");
        updateStatus("No email item selected or available.", true);
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
    console.log("ensureGraphToken() called");
    
    if (graphAccessToken) {
        console.log("Using existing Graph token.");
        return graphAccessToken;
    }
    
    console.log("No existing token, attempting to get access token via Office SSO...");
    
    // Check if Office.auth is available
    if (!Office.auth) {
        console.error("Office.auth is not available! This might be due to missing permissions or incorrect manifest configuration.");
        updateStatus("Authentication API not available. Check add-in permissions.", true);
        return null;
    }
    
    try {
        console.log("Calling Office.auth.getAccessToken...");
        const ssoToken = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true
        });
        
        console.log("SSO token received (length: " + (ssoToken ? ssoToken.length : 0) + " chars)");
        graphAccessToken = ssoToken;
        updateStatus("Successfully obtained an SSO token.", false);
        return graphAccessToken;
    } catch (error) {
        console.error("SSO Error details:", error);
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
    console.log("handleSSOError() called with error code:", error.code);
    
    let message = `SSO Error Code: ${error.code}\nMessage: ${error.message}`;
    console.error(message);
    
    switch (error.code) {
        case 13001:
            console.log("The user is not logged in, or the user cancelled without providing consent.");
            updateStatus("Please sign in to Microsoft 365 and grant permissions to this add-in.", true);
            break;
        case 13002:
            console.log("The user's identity or access token is invalid.");
            updateStatus("Authentication error: Invalid user identity or access token.", true);
            break;
        case 13003:
            console.log("A resource that the application requires is unavailable.");
            updateStatus("Required resource is unavailable. Check add-in configuration.", true);
            break;
        case 13005:
            console.log("The add-in requested a token for a resource that hasn't been configured.");
            updateStatus("Authentication error: Resource not configured in Azure AD.", true);
            break;
        case 13010:
            console.log("User interaction is required to get the access token.");
            updateStatus("Please complete the sign-in process when prompted.", true);
            break;
        case 13012:
            console.log("The user or administrator has not consented to use the application.");
            updateStatus("Please grant consent to this application when prompted.", true);
            break;
        default:
            updateStatus(`SSO Error: ${error.message || error}`, true);
            break;
    }
}

async function fetchPlannerPlans() {
    console.log("fetchPlannerPlans() called");
    showLoading(true);
    updateStatus("Loading your Planner plans...", false);
    planToGroupMap = {}; // Reset map
    
    try {
        console.log("Ensuring Graph token is available...");
        const token = await ensureGraphToken();
        
        if (!token) {
            console.error("Authentication token not available. Cannot fetch plans.");
            updateStatus("Authentication token not available. Cannot fetch plans.", true);
            showLoading(false);
            return;
        }
        
        console.log("Token obtained, making request to Microsoft Graph API for plans...");
        console.log("Request URL:", `${graphApiEndpoint}/me/planner/plans?$select=id,title,container`);
        
        const response = await fetch(`${graphApiEndpoint}/me/planner/plans?$select=id,title,container`, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });
        
        console.log("Graph API response status:", response.status, response.statusText);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error("Error response body:", errorText);
            
            let errorObj;
            try {
                errorObj = JSON.parse(errorText);
            } catch (e) {
                errorObj = { error: { message: response.statusText } };
            }
            
            console.error("Error fetching plans: ", errorObj);
            throw new Error(`Graph API Error ${response.status}: ${errorObj.error ? errorObj.error.message : 'Failed to fetch plans'}`);
        }
        
        const plansData = await response.json();
        console.log("Plans data received:", plansData);
        
        const plans = plansData.value;
        console.log("Plans fetched: ", plans);
        
        const planSelector = $("#planSelector");
        planSelector.empty();
        
        if (plans && plans.length > 0) {
            console.log(`Found ${plans.length} plans, populating selector...`);
            planSelector.append($("<option></option>").attr("value", "").text("Select a plan..."));
            
            plans.forEach(plan => {
                console.log(`Adding plan: ${plan.title} (${plan.id})`);
                planSelector.append($("<option></option>").attr("value", plan.id).text(plan.title));
                
                if (plan.container && plan.container.type === 'group') {
                    console.log(`Mapping plan ${plan.id} to group ${plan.container.id}`);
                    planToGroupMap[plan.id] = plan.container.id;
                } else {
                    console.warn(`Plan ${plan.id} does not have a group container`);
                }
            });
            
            updateStatus("Planner plans loaded.", false);
        } else {
            console.warn("No plans found or accessible.");
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
    console.log(`fetchPlanMembers() called for Group ID: ${groupId}`);
    showLoading(true);
    updateStatus("Loading plan members...", false);
    
    const assigneeSelector = $("#assigneeSelector");
    assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Loading members..."));
    
    if (!groupId) {
        console.error("Cannot load members: Group ID not found for this plan.");
        updateStatus("Cannot load members: Group ID not found for this plan.", true);
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Error: Group ID missing"));
        showLoading(false);
        return;
    }
    
    try {
        console.log("Ensuring Graph token is available...");
        const token = await ensureGraphToken();
        
        if (!token) {
            console.error("Authentication token not available. Cannot fetch members.");
            updateStatus("Authentication token not available. Cannot fetch members.", true);
            showLoading(false);
            return;
        }
        
        console.log("Token obtained, making request to Microsoft Graph API for group members...");
        console.log("Request URL:", `${graphApiEndpoint}/groups/${groupId}/members?$select=id,displayName,userPrincipalName`);
        
        const response = await fetch(`${graphApiEndpoint}/groups/${groupId}/members?$select=id,displayName,userPrincipalName`, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + token,
                "Accept": "application/json"
            }
        });
        
        console.log("Graph API response status:", response.status, response.statusText);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error("Error response body:", errorText);
            
            let errorObj;
            try {
                errorObj = JSON.parse(errorText);
            } catch (e) {
                errorObj = { error: { message: response.statusText } };
            }
            
            console.error("Error fetching group members: ", errorObj);
            throw new Error(`Graph API Error ${response.status}: ${errorObj.error ? errorObj.error.message : 'Failed to fetch members'}`);
        }
        
        const membersData = await response.json();
        console.log("Members data received:", membersData);
        
        const members = membersData.value.filter(member => member["@odata.type"] === "#microsoft.graph.user");
        console.log("Members fetched: ", members);
        
        assigneeSelector.empty();
        
        if (members && members.length > 0) {
            console.log(`Found ${members.length} members, populating selector...`);
            assigneeSelector.append($("<option></option>").attr("value", "").text("Select an assignee (optional)..."));
            
            members.forEach(member => {
                console.log(`Adding member: ${member.displayName || member.userPrincipalName} (${member.id})`);
                assigneeSelector.append($("<option></option>").attr("value", member.id).text(member.displayName || member.userPrincipalName));
            });
            
            updateStatus("Plan members loaded.", false);
        } else {
            console.warn("No members found in this plan's group.");
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
    console.log("handlePlanSelectionChange() called with planId:", selectedPlanId);
    
    const assigneeSelector = $("#assigneeSelector");
    
    if (selectedPlanId) {
        console.log("Plan selected: " + selectedPlanId);
        const groupId = planToGroupMap[selectedPlanId];
        
        if (groupId) {
            console.log("Found group ID for plan:", groupId);
            fetchPlanMembers(groupId);
        } else {
            console.warn("No Group ID found mapped for plan ID: " + selectedPlanId);
            updateStatus("Could not determine group for this plan to load members.", true);
            assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Cannot load members for this plan"));
        }
    } else {
        console.log("No plan selected, resetting assignee selector");
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Select a plan first..."));
    }
}

async function createTask() {
    console.log("createTask() called");
    
    // Get form values
    const planId = $("#planSelector").val();
    const assigneeId = $("#assigneeSelector").val();
    const taskTitle = $("#taskTitle").val();
    const dueDate = $("#dueDate").val();
    const description = $("#taskDescription").val();
    
    console.log("Form values:", {
        planId,
        assigneeId,
        taskTitle,
        dueDate,
        description: description ? description.substring(0, 50) + "..." : "(empty)"
    });
    
    // Validate required fields
    if (!planId) {
        console.error("No plan selected");
        updateStatus("Please select a plan.", true);
        return;
    }
    
    if (!taskTitle) {
        console.error("No task title provided");
        updateStatus("Please provide a task title.", true);
        return;
    }
    
    showLoading(true);
    updateStatus("Creating task...", false);
    
    try {
        console.log("Ensuring Graph token is available...");
        const token = await ensureGraphToken();
        
        if (!token) {
            console.error("Authentication token not available. Cannot create task.");
            updateStatus("Authentication token not available. Cannot create task.", true);
            showLoading(false);
            return;
        }
        
        // Create the task
        console.log("Creating task in plan:", planId);
        
        // Prepare the task details
        const taskDetails = {
            planId: planId,
            title: taskTitle,
            details: {
                description: description
            }
        };
        
        // Add due date if provided
        if (dueDate) {
            console.log("Adding due date:", dueDate);
            const dueDateObj = new Date(dueDate);
            taskDetails.dueDateTime = dueDateObj.toISOString();
        }
        
        console.log("Task details for creation:", taskDetails);
        console.log("Request URL:", `${graphApiEndpoint}/planner/tasks`);
        
        // Create the task
        const createResponse = await fetch(`${graphApiEndpoint}/planner/tasks`, {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(taskDetails)
        });
        
        console.log("Task creation response status:", createResponse.status, createResponse.statusText);
        
        if (!createResponse.ok) {
            const errorText = await createResponse.text();
            console.error("Error response body:", errorText);
            
            let errorObj;
            try {
                errorObj = JSON.parse(errorText);
            } catch (e) {
                errorObj = { error: { message: createResponse.statusText } };
            }
            
            console.error("Error creating task: ", errorObj);
            throw new Error(`Graph API Error ${createResponse.status}: ${errorObj.error ? errorObj.error.message : 'Failed to create task'}`);
        }
        
        const createdTask = await createResponse.json();
        console.log("Task created successfully:", createdTask);
        
        // If assignee is specified, assign the task
        if (assigneeId) {
            console.log("Assigning task to user:", assigneeId);
            
            // Get the etag for the task
            const taskResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${createdTask.id}/details`, {
                method: "GET",
                headers: {
                    "Authorization": "Bearer " + token,
                    "Accept": "application/json"
                }
            });
            
            if (!taskResponse.ok) {
                console.error("Error getting task details for assignment:", taskResponse.statusText);
                throw new Error(`Failed to get task details for assignment: ${taskResponse.statusText}`);
            }
            
            const taskDetails = await taskResponse.json();
            const etag = taskResponse.headers.get("etag");
            
            console.log("Task details retrieved, etag:", etag);
            
            // Assign the task
            const assignmentResponse = await fetch(`${graphApiEndpoint}/planner/tasks/${createdTask.id}`, {
                method: "PATCH",
                headers: {
                    "Authorization": "Bearer " + token,
                    "Content-Type": "application/json",
                    "If-Match": etag,
                    "Prefer": "return=representation"
                },
                body: JSON.stringify({
                    assignments: {
                        [assigneeId]: {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !"
                        }
                    }
                })
            });
            
            console.log("Assignment response status:", assignmentResponse.status, assignmentResponse.statusText);
            
            if (!assignmentResponse.ok) {
                console.error("Error assigning task:", assignmentResponse.statusText);
                // We don't throw here because the task was created successfully
                updateStatus("Task created but assignment failed.", true);
            } else {
                console.log("Task assigned successfully");
            }
        }
        
        // Success!
        updateStatus("Task created successfully!", false);
        
        // Reset form
        $("#taskTitle").val("");
        $("#taskDescription").val("");
        $("#dueDate").val("");
        
    } catch (error) {
        console.error("Failed to create task: ", error);
        updateStatus(`Error creating task: ${error.message}`, true);
    } finally {
        showLoading(false);
    }
}

// Helper function to show/hide loading indicator
function showLoading(show) {
    console.log("showLoading(" + show + ") called");
    if (show) {
        $("#loadingIndicator").css("display", "block");
    } else {
        $("#loadingIndicator").css("display", "none");
    }
}

// Helper function to update status message
function updateStatus(message, isError) {
    console.log("updateStatus() called with message:", message, "isError:", isError);
    const statusElement = $("#statusMessage");
    statusElement.text(message);
    
    if (isError) {
        statusElement.addClass("ms-fontColor-error");
    } else {
        statusElement.removeClass("ms-fontColor-error");
    }
}

// Add a global error handler to catch unhandled exceptions
window.onerror = function(message, source, lineno, colno, error) {
    console.error("Unhandled error:", message, "at", source, ":", lineno, ":", colno);
    console.error("Error object:", error);
    updateStatus("An unexpected error occurred. See console for details.", true);
    return true; // Prevents the default browser error handling
};

// Log that the script has loaded
console.log("Debug version of taskpane.js loaded successfully");
