// Optimized taskpane.js with improved authentication flow
// This version maintains all permissions while addressing admin approval issues

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
        
        // MSAL configuration with improved settings for public use
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

        // Configure the request with your existing permissions
        // All these permissions are user-consentable according to your screenshot
        const requestObj = {
            scopes: [
                'User.Read',
                'Files.ReadWrite.All',
                'Mail.Read',
                'offline_access',
                'openid',
                'profile',
                'Tasks.ReadWrite',
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

        // Authenticate using popup for Outlook desktop with improved error handling
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
                // Also add prompt=consent to force a fresh consent experience
                authParams.prompt = 'consent';
            }
            
            console.log("[DEBUG] Starting popup authentication with params:", authParams);
            msalInstance.loginPopup(authParams)
                .then(function(loginResponse) {
                    updateStatus('Authentication successful. Loading plans...');
                    console.log("[DEBUG] Login response:", loginResponse);
                    loadPlannerPlans();
                })
                .catch(function(error) {
                    // Enhanced error handling with specific guidance
                    if (error.errorCode === "AADSTS700016") {
                        showError('This application is not available in your organization. Please contact your IT administrator.');
                        console.error("Application not found in directory:", error);
                        
                        // Show additional guidance
                        const guidanceElement = document.createElement('div');
                        guidanceElement.innerHTML = `
                            <div style="margin-top: 20px; padding: 10px; border: 1px solid #e0e0e0; background-color: #f9f9f9;">
                                <h3>Troubleshooting Guidance</h3>
                                <p>This error occurs when the application is not registered in your organization's directory.</p>
                                <p>Options to resolve:</p>
                                <ul>
                                    <li>Try signing in with a different account</li>
                                    <li>Ask your IT administrator to allow external applications</li>
                                </ul>
                            </div>
                        `;
                        document.body.appendChild(guidanceElement);
                    } else if (error.errorCode === "AADSTS900971") {
                        showError('Authentication error: No reply address provided. Please check Azure AD app registration.');
                        console.error("Redirect URI error. Current URL:", window.location.href);
                        
                        // Show redirect URI guidance
                        const guidanceElement = document.createElement('div');
                        guidanceElement.innerHTML = `
                            <div style="margin-top: 20px; padding: 10px; border: 1px solid #e0e0e0; background-color: #f9f9f9;">
                                <h3>Redirect URI Issue</h3>
                                <p>The current page URL needs to be registered in Azure AD:</p>
                                <code>${currentUrl}</code>
                            </div>
                        `;
                        document.body.appendChild(guidanceElement);
                    } else if (error.errorCode === "AADSTS65001") {
                        showError('You need to consent to the permissions requested by this application.');
                        
                        // Offer retry with explicit consent
                        const retryButton = document.createElement('button');
                        retryButton.textContent = 'Retry with explicit consent';
                        retryButton.className = 'ms-Button ms-Button--primary';
                        retryButton.style.margin = '20px 0';
                        retryButton.onclick = function() {
                            const retryParams = { ...authParams, prompt: 'consent' };
                            msalInstance.loginPopup(retryParams);
                        };
                        document.body.appendChild(retryButton);
                    } else {
                        showError('Authentication error: ' + error.message);
                    }
                    console.error("Authentication error:", error);
                });
        }

        // Get access token for Microsoft Graph API with improved token acquisition
        function getAccessToken() {
            console.log("[DEBUG] Getting access token");
            
            if (!msalInstance) {
                return Promise.reject(new Error('Authentication library not initialized'));
            }
            
            // First try to get token silently
            return msalInstance.acquireTokenSilent(requestObj)
                .then(function(tokenResponse) {
                    console.log("[DEBUG] Token acquired silently");
                    return tokenResponse.accessToken;
                })
                .catch(function(error) {
                    console.log("[DEBUG] Error in silent token acquisition:", error);
                    
                    // If silent acquisition fails, try popup with improved error handling
                    if (error.name === "InteractionRequiredAuthError") {
                        console.log("[DEBUG] Interaction required, using popup");
                        updateStatus("Additional authentication required...");
                        
                        return msalInstance.acquireTokenPopup(requestObj)
                            .then(function(tokenResponse) {
                                console.log("[DEBUG] Token acquired via popup");
                                return tokenResponse.accessToken;
                            })
                            .catch(function(popupError) {
                                console.error("[DEBUG] Popup token acquisition failed:", popupError);
                                
                                // If popup fails due to consent issues, try one more time with explicit consent prompt
                                if (popupError.errorCode === "AADSTS65001") {
                                    console.log("[DEBUG] Consent required, retrying with explicit consent prompt");
                                    const consentRequestObj = { ...requestObj, prompt: 'consent' };
                                    
                                    return msalInstance.acquireTokenPopup(consentRequestObj)
                                        .then(function(tokenResponse) {
                                            console.log("[DEBUG] Token acquired with explicit consent");
                                            return tokenResponse.accessToken;
                                        });
                                }
                                
                                throw popupError;
                            });
                    }
                    
                    throw error;
                });
        }

        // Load Planner plans
        function loadPlannerPlans() {
            updateStatus('Loading your Planner plans...');
            showElement('loadingIndicator');
            
            getAccessToken()
                .then(function(accessToken) {
                    console.log("[DEBUG] Access token obtained, length:", accessToken.length);
                    // Get all plans the user has access to
                    updateStatus('Fetching plans from Microsoft Graph API...');
                    return fetch('https://graph.microsoft.com/v1.0/me/planner/plans', {
                        headers: {
                            'Authorization': 'Bearer ' + accessToken
                        }
                    });
                })
                .then(function(response) {
                    console.log("[DEBUG] Plans API response status:", response.status);
                    if (!response.ok) {
                        throw new Error('Failed to fetch plans: ' + response.status);
                    }
                    return response.json();
                })
                .then(function(data) {
                    console.log("[DEBUG] Plans data received:", data);
                    hideElement('loadingIndicator');
                    
                    const plannerSelect = document.getElementById('planSelector');
                    if (!plannerSelect) {
                        throw new Error('planSelector element not found');
                    }
                    
                    if (data.value && data.value.length > 0) {
                        updateStatus('Plans loaded successfully: ' + data.value.length + ' plans found');
                        
                        // Clear existing options
                        plannerSelect.innerHTML = '';
                        
                        // Add default option
                        const defaultOption = document.createElement('option');
                        defaultOption.value = '';
                        defaultOption.text = '-- Select a plan --';
                        plannerSelect.appendChild(defaultOption);
                        
                        // Add plans to dropdown
                        data.value.forEach(function(plan) {
                            const option = document.createElement('option');
                            option.value = plan.id;
                            option.text = plan.title;
                            plannerSelect.appendChild(option);
                            console.log("[DEBUG] Added plan:", plan.title);
                        });
                    } else {
                        updateStatus('No plans found. Please create a plan in Microsoft Planner first.');
                        console.log("[DEBUG] No plans found in data:", data);
                    }
                })
                .catch(function(error) {
                    hideElement('loadingIndicator');
                    showError('Error loading plans: ' + error.message);
                    console.error("Error loading plans:", error);
                    
                    // Provide recovery options for common errors
                    if (error.message.includes('401') || error.message.includes('403')) {
                        const retryButton = document.createElement('button');
                        retryButton.textContent = 'Retry with new authentication';
                        retryButton.className = 'ms-Button ms-Button--primary';
                        retryButton.style.margin = '20px 0';
                        retryButton.onclick = function() {
                            msalInstance.logout();
                            setTimeout(() => {
                                authenticateWithPopup();
                            }, 1000);
                        };
                        document.body.appendChild(retryButton);
                    }
                });
        }

        // Handle plan selection change
        function onPlannerSelectionChange() {
            const plannerSelect = document.getElementById('planSelector');
            if (!plannerSelect) {
                showError('planSelector element not found');
                return;
            }
            
            const selectedPlanId = plannerSelect.value;
            console.log("[DEBUG] Plan selection changed to:", selectedPlanId);
            
            if (selectedPlanId) {
                updateStatus('Loading assignees for selected plan...');
                showElement('loadingIndicator');
                
                getAccessToken()
                    .then(function(accessToken) {
                        // Get group members for the selected plan
                        return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${selectedPlanId}/details`, {
                            headers: {
                                'Authorization': 'Bearer ' + accessToken
                            }
                        });
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Plan details API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch plan details: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(planDetails) {
                        console.log("[DEBUG] Plan details received:", planDetails);
                        
                        // Get the group ID associated with the plan
                        return getAccessToken().then(accessToken => {
                            return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${selectedPlanId}`, {
                                headers: {
                                    'Authorization': 'Bearer ' + accessToken
                                }
                            });
                        });
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Plan API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch plan: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(planData) {
                        console.log("[DEBUG] Plan data received:", planData);
                        
                        if (planData.owner) {
                            const groupId = planData.owner;
                            
                            // Try to get members of the group
                            return getAccessToken().then(accessToken => {
                                return fetch(`https://graph.microsoft.com/v1.0/groups/${groupId}/members`, {
                                    headers: {
                                        'Authorization': 'Bearer ' + accessToken
                                    }
                                });
                            })
                            .catch(error => {
                                // If we can't get group members, fall back to current user only
                                console.warn("Could not fetch group members:", error);
                                return { ok: false, status: 403 };
                            });
                        } else {
                            throw new Error('No owner group found for this plan');
                        }
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Members API response status:", response.status);
                        
                        // If we can access group members, use them
                        if (response.ok) {
                            return response.json().then(membersData => {
                                return { type: 'group', data: membersData };
                            });
                        } 
                        // Otherwise fall back to just the current user
                        else {
                            console.log("[DEBUG] Falling back to current user as assignee");
                            return { 
                                type: 'currentUser', 
                                data: { 
                                    value: [{ 
                                        id: msalInstance.getAccount().accountIdentifier,
                                        displayName: msalInstance.getAccount().name || msalInstance.getAccount().userName
                                    }] 
                                } 
                            };
                        }
                    })
                    .then(function(result) {
                        console.log("[DEBUG] Assignee data received:", result);
                        hideElement('loadingIndicator');
                        
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
                        
                        // Add members to dropdown
                        const membersData = result.data;
                        if (membersData.value && membersData.value.length > 0) {
                            membersData.value.forEach(function(member) {
                                const option = document.createElement('option');
                                option.value = member.id;
                                option.text = member.displayName || member.userPrincipalName || member.id;
                                assigneeSelect.appendChild(option);
                                console.log("[DEBUG] Added assignee:", option.text);
                            });
                            
                            if (result.type === 'currentUser') {
                                updateStatus('Only your user account is available as assignee');
                            } else {
                                updateStatus('Assignees loaded successfully');
                            }
                        } else {
                            updateStatus('No assignees found for this plan');
                        }
                    })
                    .catch(function(error) {
                        hideElement('loadingIndicator');
                        showError('Error loading assignees: ' + error.message);
                        console.error("Error loading assignees:", error);
                        
                        // Add fallback for assignee selection
                        const assigneeSelect = document.getElementById('assigneeSelector');
                        if (assigneeSelect) {
                            // Clear existing options
                            assigneeSelect.innerHTML = '';
                            
                            // Add default option
                            const defaultOption = document.createElement('option');
                            defaultOption.value = '';
                            defaultOption.text = '-- Select an assignee --';
                            assigneeSelect.appendChild(defaultOption);
                            
                            // Add current user as fallback
                            if (msalInstance && msalInstance.getAccount()) {
                                const currentUser = msalInstance.getAccount();
                                const option = document.createElement('option');
                                option.value = currentUser.accountIdentifier || currentUser.userName;
                                option.text = currentUser.name || currentUser.userName;
                                assigneeSelect.appendChild(option);
                                console.log("[DEBUG] Added current user as fallback assignee:", option.text);
                                
                                updateStatus('Using current user as assignee (fallback mode)');
                            }
                        }
                    });
            }
        }

        // Create a new Planner task
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
            
            // First, get the buckets for the selected plan
            getAccessToken()
                .then(function(accessToken) {
                    return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${selectedPlanId}/buckets`, {
                        headers: {
                            'Authorization': 'Bearer ' + accessToken
                        }
                    });
                })
                .then(function(response) {
                    if (!response.ok) {
                        throw new Error('Failed to fetch buckets: ' + response.status);
                    }
                    return response.json();
                })
                .then(function(bucketsData) {
                    console.log("[DEBUG] Buckets data received:", bucketsData);
                    
                    // Use the first bucket or create a task without a bucket
                    let bucketId = null;
                    if (bucketsData.value && bucketsData.value.length > 0) {
                        bucketId = bucketsData.value[0].id;
                    }
                    
                    // Create the task
                    const taskDetails = {
                        planId: selectedPlanId,
                        title: taskTitle,
                        assignments: {}
                    };
                    
                    // Add bucket if available
                    if (bucketId) {
                        taskDetails.bucketId = bucketId;
                    }
                    
                    // Add description if available
                    if (taskDescription) {
                        taskDetails.details = {
                            description: taskDescription
                        };
                    }
                    
                    // Add assignee if selected
                    if (selectedAssigneeId) {
                        taskDetails.assignments[selectedAssigneeId] = {
                            '@odata.type': '#microsoft.graph.plannerAssignment',
                            'orderHint': ' !'
                        };
                    }
                    
                    return getAccessToken().then(accessToken => {
                        return fetch('https://graph.microsoft.com/v1.0/planner/tasks', {
                            method: 'POST',
                            headers: {
                                'Authorization': 'Bearer ' + accessToken,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify(taskDetails)
                        });
                    });
                })
                .then(function(response) {
                    if (!response.ok) {
                        return response.text().then(text => {
                            throw new Error('Failed to create task: ' + response.status + ' - ' + text);
                        });
                    }
                    return response.json();
                })
                .then(function(taskData) {
                    console.log("[DEBUG] Task created:", taskData);
                    hideElement('loadingIndicator');
                    updateStatus('Task created successfully!');
                    
                    // Clear form or reset values
                    document.getElementById('taskTitle').value = '';
                    document.getElementById('taskDescription').value = '';
                    document.getElementById('assigneeSelector').innerHTML = '';
                    
                    // Show success message
                    showSuccess('Task "' + taskTitle + '" created successfully in Planner!');
                    
                    // Add link to view the task if possible
                    if (taskData.id) {
                        const viewTaskLink = document.createElement('a');
                        viewTaskLink.href = `https://tasks.office.com/`;
                        viewTaskLink.target = '_blank';
                        viewTaskLink.textContent = 'View your tasks in Planner';
                        viewTaskLink.style.display = 'block';
                        viewTaskLink.style.margin = '10px 0';
                        document.body.appendChild(viewTaskLink);
                    }
                })
                .catch(function(error) {
                    hideElement('loadingIndicator');
                    showError('Error creating task: ' + error.message);
                    console.error("Error creating task:", error);
                    
                    // If task creation fails, offer manual creation option
                    if (error.message.includes('403') || error.message.includes('401')) {
                        const manualCreationDiv = document.createElement('div');
                        manualCreationDiv.style.margin = '20px 0';
                        manualCreationDiv.style.padding = '10px';
                        manualCreationDiv.style.border = '1px solid #e0e0e0';
                        manualCreationDiv.style.backgroundColor = '#f9f9f9';
                        
                        manualCreationDiv.innerHTML = `
                            <h3>Task Details for Manual Creation</h3>
                            <p>You can create this task manually in Planner:</p>
                            <p><strong>Title:</strong> ${taskTitle}</p>
                            <p><strong>Description:</strong> ${taskDescription}</p>
                            <a href="https://tasks.office.com/" target="_blank" class="ms-Button ms-Button--primary">
                                Open Microsoft Planner
                            </a>
                        `;
                        
                        document.body.appendChild(manualCreationDiv);
                    }
                });
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
