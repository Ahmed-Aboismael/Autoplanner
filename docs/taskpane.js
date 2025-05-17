// Reduced permissions taskpane.js without admin approval requirement
// This version uses delegated permissions that don't require admin consent

// Use Office.initialize instead of Office.onReady to ensure proper loading sequence
Office.initialize = function (reason) {
    // This function is called when the Office.js library is fully loaded
    console.log("[DEBUG] Office.initialize called with reason:", reason);
    
    // Set up the UI once Office.js is fully initialized
    $(document).ready(function() {
        console.log("[DEBUG] Document ready event fired");
        
        // Get the current page URL to use as redirect URI
        const currentUrl = window.location.href.split('?')[0]; // Remove any query parameters
        console.log("[DEBUG] Using current URL as redirect URI:", currentUrl);
        
        // MSAL configuration with reduced permission scopes
        const msalConfig = {
            auth: {
                clientId: '60ca32af-6d83-4369-8a0a-dce7bb909d9d',
                // Use 'organizations' for multi-tenant business apps
                authority: 'https://login.microsoftonline.com/organizations',
                redirectUri: currentUrl, // Use current page URL as redirect URI
                postLogoutRedirectUri: currentUrl,
                navigateToLoginRequestUrl: true
            },
            cache: {
                cacheLocation: 'localStorage',
                storeAuthStateInCookie: true
            },
            system: {
                loggerOptions: {
                    loggerCallback: (level, message, containsPii) => {
                        if (!containsPii) console.log("[MSAL]", message);
                    },
                    piiLoggingEnabled: false,
                    logLevel: 3 // Verbose logging for debugging
                }
            }
        };

        // Create the MSAL application object
        let msalInstance;
        try {
            msalInstance = new Msal.UserAgentApplication(msalConfig);
            console.log("[DEBUG] MSAL initialized successfully");
            updateStatus("MSAL initialized successfully");
            
            // Register redirect callback
            msalInstance.handleRedirectCallback((error, response) => {
                if (error) {
                    console.error("[DEBUG] Redirect callback error:", error);
                    showError("Authentication error: " + error.message);
                } else {
                    console.log("[DEBUG] Redirect callback success:", response);
                    updateStatus("Authentication successful");
                    loadPlannerPlans();
                }
            });
        } catch (error) {
            console.error("Failed to initialize MSAL:", error);
            showError("Failed to initialize MSAL: " + error.message);
        }

        // Configure the request with REDUCED permission scopes that don't require admin consent
        const requestObj = {
            scopes: [
                'User.Read', // Basic profile - doesn't require admin consent
                'Tasks.Read', // Read-only access to tasks - less privileged than Tasks.ReadWrite
                'Mail.Read' // Read mail - still needed but can be consented by user
                // Removed Group.Read.All which requires admin consent
            ]
        };
        
        // Set up status message area
        const statusElement = document.getElementById('statusMessage');
        if (statusElement) {
            statusElement.textContent = "Add-in initialized successfully";
        }
        
        // Load email data (subject and body) using the correct API pattern
        try {
            console.log("[DEBUG] Getting email subject");
            
            // Check if we're in a valid Outlook context
            if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
                // Direct property access for subject
                const emailSubject = Office.context.mailbox.item.subject;
                console.log("[DEBUG] Email subject:", emailSubject);
                
                const titleInput = document.getElementById('taskTitle');
                if (titleInput) {
                    titleInput.value = emailSubject || "";
                    console.log("[DEBUG] Email subject loaded");
                } else {
                    console.error("taskTitle element not found");
                }
                
                // For body, we need to use getBodyAsync
                Office.context.mailbox.item.body.getAsync("text", function(asyncResult) {
                    console.log("[DEBUG] Email body result:", asyncResult);
                    
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const emailBody = asyncResult.value;
                        const descriptionTextarea = document.getElementById('taskDescription');
                        
                        if (descriptionTextarea) {
                            descriptionTextarea.value = emailBody || "";
                            console.log("[DEBUG] Email body loaded, length:", emailBody.length);
                        } else {
                            console.error("taskDescription element not found");
                        }
                    } else {
                        console.error("Error getting email body:", asyncResult.error);
                    }
                });
            } else {
                console.error("Not in a valid Outlook mail item context");
                if (statusElement) {
                    statusElement.textContent = "Error: Not in a valid email context";
                    statusElement.style.color = "red";
                }
            }
        } catch (error) {
            console.error("Error accessing email data:", error);
            if (statusElement) {
                statusElement.textContent = "Error accessing email data: " + error.message;
                statusElement.style.color = "red";
            }
        }
        
        // Set up create task button
        const createButton = document.getElementById('createTaskButton');
        if (createButton) {
            createButton.onclick = createPlannerTask;
            console.log("[DEBUG] Create task button handler set");
        } else {
            console.error("createTaskButton not found");
        }
        
        // Set up planner selection dropdown
        const plannerSelect = document.getElementById('planSelector');
        if (plannerSelect) {
            plannerSelect.onchange = onPlannerSelectionChange;
            console.log("[DEBUG] Planner selection handler set");
        } else {
            console.error("planSelector not found");
        }
        
        // Check if user is already signed in
        if (msalInstance && msalInstance.getAccount()) {
            const account = msalInstance.getAccount();
            updateStatus('User already signed in as: ' + account.userName + '. Loading plans...');
            console.log("[DEBUG] User account:", account);
            loadPlannerPlans();
        } else {
            updateStatus('Please sign in to access your Planner plans');
            // Add a sign-in button instead of auto-authenticating
            addSignInButton();
        }

        // Add a sign-in button to the UI
        function addSignInButton() {
            const signInArea = document.createElement('div');
            signInArea.style.margin = '20px 0';
            signInArea.style.textAlign = 'center';
            
            const signInButton = document.createElement('button');
            signInButton.textContent = 'Sign in with Microsoft';
            signInButton.className = 'ms-Button ms-Button--primary';
            signInButton.style.padding = '8px 16px';
            signInButton.onclick = authenticateWithPopup;
            
            signInArea.appendChild(signInButton);
            
            // Find a good place to insert the button
            const formElement = document.querySelector('form') || document.body;
            formElement.insertBefore(signInArea, formElement.firstChild);
            
            console.log("[DEBUG] Sign-in button added");
        }

        // Authenticate using popup for Outlook desktop
        function authenticateWithPopup() {
            updateStatus('Authenticating...');
            
            if (!msalInstance) {
                showError('Authentication library not initialized');
                return;
            }
            
            // Add loginHint if available to improve the sign-in experience
            const loginHint = Office.context.mailbox.userProfile.emailAddress;
            const authParams = { ...requestObj };
            if (loginHint) {
                authParams.loginHint = loginHint;
            }
            
            console.log("[DEBUG] Starting popup authentication");
            msalInstance.loginPopup(authParams)
                .then(function(loginResponse) {
                    updateStatus('Authentication successful. Loading plans...');
                    console.log("[DEBUG] Login response:", loginResponse);
                    loadPlannerPlans();
                })
                .catch(function(error) {
                    // Handle specific multi-tenant errors
                    if (error.errorCode === "AADSTS700016") {
                        showError('This application is not available in your organization. Please contact your IT administrator.');
                    } else if (error.errorCode === "AADSTS900971") {
                        showError('Authentication error: No reply address provided. Please check Azure AD app registration.');
                        console.error("Redirect URI error. Current URL:", window.location.href);
                    } else if (error.errorCode === "AADSTS65001") {
                        showError('You need to consent to the permissions requested by this application.');
                    } else {
                        showError('Authentication error: ' + error.message);
                    }
                    console.error("Authentication error:", error);
                });
        }

        // Get access token for Microsoft Graph API
        function getAccessToken() {
            console.log("[DEBUG] Getting access token");
            
            if (!msalInstance) {
                return Promise.reject(new Error('Authentication library not initialized'));
            }
            
            return msalInstance.acquireTokenSilent(requestObj)
                .then(function(tokenResponse) {
                    console.log("[DEBUG] Token acquired silently");
                    return tokenResponse.accessToken;
                })
                .catch(function(error) {
                    console.log("[DEBUG] Error in silent token acquisition:", error);
                    if (error.name === "InteractionRequiredAuthError") {
                        console.log("[DEBUG] Interaction required, using popup");
                        return msalInstance.acquireTokenPopup(requestObj)
                            .then(function(tokenResponse) {
                                console.log("[DEBUG] Token acquired via popup");
                                return tokenResponse.accessToken;
                            });
                    }
                    throw error;
                });
        }

        // Load Planner plans - MODIFIED to work with reduced permissions
        function loadPlannerPlans() {
            updateStatus('Loading your Planner plans...');
            showElement('loadingIndicator');
            
            getAccessToken()
                .then(function(accessToken) {
                    console.log("[DEBUG] Access token obtained, length:", accessToken.length);
                    // Get plans the user has access to - using /me/planner/tasks instead of /me/planner/plans
                    // This works with Tasks.Read permission instead of requiring Group.Read.All
                    updateStatus('Fetching tasks from Microsoft Graph API...');
                    return fetch('https://graph.microsoft.com/v1.0/me/planner/tasks', {
                        headers: {
                            'Authorization': 'Bearer ' + accessToken
                        }
                    });
                })
                .then(function(response) {
                    console.log("[DEBUG] Tasks API response status:", response.status);
                    if (!response.ok) {
                        throw new Error('Failed to fetch tasks: ' + response.status);
                    }
                    return response.json();
                })
                .then(function(data) {
                    console.log("[DEBUG] Tasks data received:", data);
                    
                    // Extract unique plan IDs from tasks
                    const planIds = new Set();
                    if (data.value && data.value.length > 0) {
                        data.value.forEach(task => {
                            if (task.planId) {
                                planIds.add(task.planId);
                            }
                        });
                    }
                    
                    // If we found plan IDs, fetch details for each plan
                    if (planIds.size > 0) {
                        const planPromises = Array.from(planIds).map(planId => {
                            return getAccessToken().then(accessToken => {
                                return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${planId}`, {
                                    headers: {
                                        'Authorization': 'Bearer ' + accessToken
                                    }
                                })
                                .then(response => {
                                    if (!response.ok) {
                                        console.warn(`Could not fetch details for plan ${planId}: ${response.status}`);
                                        return null;
                                    }
                                    return response.json();
                                })
                                .catch(error => {
                                    console.warn(`Error fetching plan ${planId}:`, error);
                                    return null;
                                });
                            });
                        });
                        
                        return Promise.all(planPromises);
                    } else {
                        return [];
                    }
                })
                .then(function(plansData) {
                    console.log("[DEBUG] Plans data received:", plansData);
                    hideElement('loadingIndicator');
                    
                    const plannerSelect = document.getElementById('planSelector');
                    if (!plannerSelect) {
                        throw new Error('planSelector element not found');
                    }
                    
                    // Filter out null results (failed plan fetches)
                    const validPlans = plansData.filter(plan => plan !== null);
                    
                    if (validPlans && validPlans.length > 0) {
                        updateStatus('Plans loaded successfully: ' + validPlans.length + ' plans found');
                        
                        // Clear existing options
                        plannerSelect.innerHTML = '';
                        
                        // Add default option
                        const defaultOption = document.createElement('option');
                        defaultOption.value = '';
                        defaultOption.text = '-- Select a plan --';
                        plannerSelect.appendChild(defaultOption);
                        
                        // Add plans to dropdown
                        validPlans.forEach(function(plan) {
                            const option = document.createElement('option');
                            option.value = plan.id;
                            option.text = plan.title;
                            plannerSelect.appendChild(option);
                            console.log("[DEBUG] Added plan:", plan.title);
                        });
                    } else {
                        updateStatus('No plans found. You may need to create a plan in Microsoft Planner first.');
                        console.log("[DEBUG] No plans found in data");
                    }
                })
                .catch(function(error) {
                    hideElement('loadingIndicator');
                    showError('Error loading plans: ' + error.message);
                    console.error("Error loading plans:", error);
                });
        }

        // Handle plan selection change - MODIFIED to work with reduced permissions
        function onPlannerSelectionChange() {
            const plannerSelect = document.getElementById('planSelector');
            if (!plannerSelect) {
                showError('planSelector element not found');
                return;
            }
            
            const selectedPlanId = plannerSelect.value;
            console.log("[DEBUG] Plan selection changed to:", selectedPlanId);
            
            if (selectedPlanId) {
                updateStatus('Loading tasks for selected plan...');
                showElement('loadingIndicator');
                
                getAccessToken()
                    .then(function(accessToken) {
                        // Get tasks for the selected plan
                        return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${selectedPlanId}/tasks`, {
                            headers: {
                                'Authorization': 'Bearer ' + accessToken
                            }
                        });
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Plan tasks API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch plan tasks: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(tasksData) {
                        console.log("[DEBUG] Plan tasks received:", tasksData);
                        hideElement('loadingIndicator');
                        
                        // Since we can't get group members with reduced permissions,
                        // we'll just populate the assignee dropdown with the current user
                        const assigneeSelect = document.getElementById('assigneeSelector');
                        if (!assigneeSelect) {
                            throw new Error('assigneeSelector element not found');
                        }
                        
                        // Clear existing options
                        assigneeSelect.innerHTML = '';
                        
                        // Add default option
                        const defaultOption = document.createElement('option');
                        defaultOption.value = '';
                        defaultOption.text = '-- Select an assignee --';
                        assigneeSelect.appendChild(defaultOption);
                        
                        // Add current user as the only assignee option
                        if (msalInstance && msalInstance.getAccount()) {
                            const currentUser = msalInstance.getAccount();
                            const option = document.createElement('option');
                            option.value = currentUser.accountIdentifier || currentUser.userName;
                            option.text = currentUser.name || currentUser.userName;
                            assigneeSelect.appendChild(option);
                            console.log("[DEBUG] Added current user as assignee:", option.text);
                        }
                        
                        updateStatus('Ready to create task');
                    })
                    .catch(function(error) {
                        hideElement('loadingIndicator');
                        showError('Error loading plan details: ' + error.message);
                        console.error("Error loading plan details:", error);
                    });
            }
        }

        // Create a new Planner task - MODIFIED to work with reduced permissions
        function createPlannerTask() {
            const taskTitle = document.getElementById('taskTitle').value;
            const taskDescription = document.getElementById('taskDescription').value;
            const selectedPlanId = document.getElementById('planSelector').value;
            const selectedAssigneeId = document.getElementById('assigneeSelector').value;
            
            if (!taskTitle) {
                showError('Please enter a task title');
                return;
            }
            
            if (!selectedPlanId) {
                showError('Please select a plan');
                return;
            }
            
            updateStatus('Creating task...');
            showElement('loadingIndicator');
            
            // With reduced permissions, we can't create tasks directly
            // Instead, we'll show a message with instructions
            setTimeout(() => {
                hideElement('loadingIndicator');
                
                // Create a message with the task details
                const taskDetails = `
                    <div style="border: 1px solid #ccc; padding: 15px; margin: 15px 0; background: #f9f9f9;">
                        <h3>Task Details</h3>
                        <p><strong>Title:</strong> ${taskTitle}</p>
                        <p><strong>Description:</strong> ${taskDescription}</p>
                        <p><strong>Plan ID:</strong> ${selectedPlanId}</p>
                        ${selectedAssigneeId ? `<p><strong>Assignee ID:</strong> ${selectedAssigneeId}</p>` : ''}
                    </div>
                    <p>Due to permission limitations, this task cannot be created automatically. Please copy these details and create the task manually in Microsoft Planner.</p>
                    <p>To enable automatic task creation, your administrator would need to grant additional permissions to this application.</p>
                `;
                
                // Display the message
                const messageArea = document.createElement('div');
                messageArea.innerHTML = taskDetails;
                
                // Find a good place to insert the message
                const formElement = document.querySelector('form') || document.body;
                formElement.appendChild(messageArea);
                
                updateStatus('Task details prepared for manual creation');
            }, 1000);
        }

        // Helper functions for UI updates
        function updateStatus(message) {
            console.log("[STATUS] " + message);
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = message;
                statusElement.style.color = '#333';
            }
        }
        
        function showError(message) {
            console.error("[ERROR] " + message);
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = message;
                statusElement.style.color = 'red';
            }
        }
        
        function showSuccess(message) {
            console.log("[SUCCESS] " + message);
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = message;
                statusElement.style.color = 'green';
            }
        }
        
        function showElement(id) {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = 'block';
            }
        }
        
        function hideElement(id) {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = 'none';
            }
        }
    });
};
