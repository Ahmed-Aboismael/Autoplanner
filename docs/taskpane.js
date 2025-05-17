// Enhanced version of taskpane.js with improved error reporting

// Office.initialize function that runs when the add-in is loaded
Office.initialize = function (reason) {
    try {
        console.log("Office.initialize started with reason:", reason);
        
        // Check if jQuery is available
        if (typeof $ === 'undefined') {
            displayError("jQuery ($) is not defined! Make sure jQuery is loaded before taskpane.js");
            return;
        }
        
        // Ensure the DOM is loaded before we try to manipulate it
        $(document).ready(function () {
            try {
                console.log("DOM ready, setting up event handlers...");
                
                // Create error display area if it doesn't exist
                if ($("#errorDisplay").length === 0) {
                    $("body").prepend('<div id="errorDisplay" style="display:none; color:red; background-color:#ffeeee; padding:10px; margin-bottom:10px; border:1px solid red;"></div>');
                }
                
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
                            displayError("Authentication failed. No token available to access Microsoft Graph API.");
                        }
                    })
                    .catch(err => {
                        console.error("Initial token fetch or plan loading failed: ", err);
                        updateStatus("Error during authentication: " + (err.message || err), true);
                        displayError("Authentication error: " + (err.message || err));
                    });
            } catch (error) {
                console.error("Error in document ready handler:", error);
                displayError("Initialization error: " + (error.message || error));
            }
        });
    } catch (error) {
        console.error("Error in Office.initialize:", error);
        // Can't use jQuery here as it might not be loaded
        alert("Critical initialization error: " + (error.message || error));
    }
};

// Function to display errors prominently in the UI
function displayError(message) {
    console.error("ERROR:", message);
    
    try {
        if ($("#errorDisplay").length > 0) {
            $("#errorDisplay").html("<strong>Error:</strong> " + message)
                .show();
        } else {
            // Fallback if jQuery or the error display div isn't available
            updateStatus("ERROR: " + message, true);
        }
    } catch (e) {
        // Last resort if even the above fails
        console.error("Failed to display error:", e);
        alert("Error: " + message);
    }
}

// Function to get details from the currently selected email
function getEmailDetails() {
    console.log("getEmailDetails() called");
    
    try {
        if (!Office || !Office.context || !Office.context.mailbox) {
            displayError("Office.context.mailbox is not available! This add-in must be run in Outlook.");
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
                    displayError("Failed to get email subject: " + asyncResult.error.message);
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
                    displayError("Failed to get email body: " + asyncResult.error.message);
                }
            });
        } else {
            console.warn("No email item selected or available.");
            updateStatus("No email item selected or available.", true);
            displayError("No email item is selected or available. Please select an email first.");
        }
    } catch (error) {
        console.error("Error in getEmailDetails:", error);
        displayError("Failed to get email details: " + (error.message || error));
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
    
    try {
        if (graphAccessToken) {
            console.log("Using existing Graph token.");
            return graphAccessToken;
        }
        
        console.log("No existing token, attempting to get access token via Office SSO...");
        
        // Check if Office.auth is available
        if (!Office.auth) {
            console.error("Office.auth is not available! This might be due to missing permissions or incorrect manifest configuration.");
            displayError("Authentication API not available. Check add-in permissions in manifest.xml and Azure AD configuration.");
            return null;
        }
        
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
            displayError(`Authentication error: ${error.message || error}`);
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
            displayError("Authentication error: You need to sign in to Microsoft 365 and grant permissions to this add-in.");
            break;
        case 13002:
            console.log("The user's identity or access token is invalid.");
            updateStatus("Authentication error: Invalid user identity or access token.", true);
            displayError("Authentication error: Your identity or access token is invalid.");
            break;
        case 13003:
            console.log("A resource that the application requires is unavailable.");
            updateStatus("Required resource is unavailable. Check add-in configuration.", true);
            displayError("Authentication error: A required resource is unavailable. Check add-in configuration.");
            break;
        case 13005:
            console.log("The add-in requested a token for a resource that hasn't been configured.");
            updateStatus("Authentication error: Resource not configured in Azure AD.", true);
            displayError("Authentication error: The resource hasn't been configured in Azure AD. Check your app registration.");
            break;
        case 13010:
            console.log("User interaction is required to get the access token.");
            updateStatus("Please complete the sign-in process when prompted.", true);
            displayError("Authentication requires your interaction. Please complete the sign-in process when prompted.");
            break;
        case 13012:
            console.log("The user or administrator has not consented to use the application.");
            updateStatus("Please grant consent to this application when prompted.", true);
            displayError("You need to grant consent to this application. Please complete the permission request when prompted.");
            break;
        default:
            updateStatus(`SSO Error: ${error.message || error}`, true);
            displayError(`Authentication error (${error.code}): ${error.message || error}`);
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
            displayError("Failed to get authentication token. Cannot access your Planner plans.");
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
                displayError(`Graph API Error (${response.status}): ${errorObj.error ? errorObj.error.message : 'Failed to fetch plans'}`);
            } catch (e) {
                errorObj = { error: { message: response.statusText } };
                displayError(`Graph API Error (${response.status}): ${response.statusText}`);
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
            displayError("No Planner plans were found or you don't have access to any plans. Make sure you have the necessary permissions.");
        }
    } catch (error) {
        console.error("Failed to fetch Planner plans: ", error);
        updateStatus(`Error loading plans: ${error.message}`, true);
        displayError(`Failed to load Planner plans: ${error.message}`);
        
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
        displayError("Cannot load members: Group ID not found for this plan.");
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
            displayError("Failed to get authentication token. Cannot load plan members.");
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
                displayError(`Graph API Error (${response.status}): ${errorObj.error ? errorObj.error.message : 'Failed to fetch members'}`);
            } catch (e) {
                errorObj = { error: { message: response.statusText } };
                displayError(`Graph API Error (${response.status}): ${response.statusText}`);
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
            displayError("No members were found in this plan's group.");
        }
    } catch (error) {
        console.error("Failed to fetch plan members: ", error);
        updateStatus(`Error loading members: ${error.message}`, true);
        displayError(`Failed to load plan members: ${error.message}`);
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
            displayError("Could not determine the group for this plan to load members.");
            assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Cannot load members for this plan"));
        }
    } else {
        console.log("No plan selected, resetting assignee selector");
        assigneeSelector.empty().append($("<option></option>").attr("value", "").text("Select a plan first..."));
    }
}

async function createTask() {
    console.log("createTask() called");
    
    try {
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
            displayError("Please select a plan before creating a task.");
            return;
        }
        
        if (!taskTitle) {
            console.error("No task title provided");
            updateStatus("Please provide a task title.", true);
            displayError("Please provide a task title.");
            return;
        }
        
        showLoading(true);
        updateStatus("Creating task...", false);
        
        console.log("Ensuring Graph token is available...");
        const token = await ensureGraphToken();
        
        if (!token) {
            console.error("Authentication token not available. Cannot create task.");
            updateStatus("Authentication token not available. Cannot create task.", true);
            displayError("Failed to get authentication token. Cannot create task.");
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
                displayError(`Failed to create task (${createResponse.status}): ${errorObj.error ? errorObj.error.message : 'Unknown error'}`);
            } catch (e) {
                errorObj = { error: { message: createResponse.statusText } };
                displayError(`Failed to create task (${createResponse.status}): ${createResponse.statusText}`);
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
                displayError("Task was created successfully, but assigning it to the selected user failed.");
            } else {
                console.log("Task assigned successfully");
            }
        }
        
        // Success!
        updateStatus("Task created successfully!", false);
        $("#errorDisplay").hide(); // Hide any previous errors
        
        // Reset form
        $("#taskTitle").val("");
        $("#taskDescription").val("");
        $("#dueDate").val("");
        
    } catch (error) {
        console.error("Failed to create task: ", error);
        updateStatus(`Error creating task: ${error.message}`, true);
        displayError(`Failed to create task: ${error.message}`);
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
    displayError(`Unhandled error: ${message} (at line ${lineno})`);
    return true; // Prevents the default browser error handling
};

// Log that the script has loaded
console.log("Enhanced error reporting version of taskpane.js loaded successfully");
