// Complete taskpane.js with full functionality
// This version combines working email loading with Planner integration

// Use Office.initialize instead of Office.onReady to ensure proper loading sequence
Office.initialize = function (reason) {
    // This function is called when the Office.js library is fully loaded
    console.log("[DEBUG] Office.initialize called with reason:", reason);
    
    // Set up the UI once Office.js is fully initialized
    $(document).ready(function() {
        console.log("[DEBUG] Document ready event fired");
        
        // MSAL configuration for authentication
        const msalConfig = {
            auth: {
                clientId: '60ca32af-6d83-4369-8a0a-dce7bb909d9d',
                authority: 'https://login.microsoftonline.com/common',
                redirectUri: 'https://ahmed-aboismael.github.io/Autoplanner/taskpane.html',
                navigateToLoginRequestUrl: false
            },
            cache: {
                cacheLocation: 'localStorage',
                storeAuthStateInCookie: true
            }
        };

        // Create the MSAL application object
        let msalInstance;
        try {
            msalInstance = new Msal.UserAgentApplication(msalConfig);
            console.log("[DEBUG] MSAL initialized successfully");
            updateStatus("MSAL initialized successfully");
        } catch (error) {
            console.error("Failed to initialize MSAL:", error);
            showError("Failed to initialize MSAL: " + error.message);
        }

        // Configure the request for Microsoft Graph scopes
        const requestObj = {
            scopes: [
                'User.Read',
                'Group.Read.All',
                'Tasks.ReadWrite',
                'Mail.Read'
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
            updateStatus('Authenticating to access your Planner plans...');
            // Auto-authenticate since we don't have an explicit auth button
            authenticateWithPopup();
        }

        // Authenticate using popup for Outlook desktop
        function authenticateWithPopup() {
            updateStatus('Authenticating...');
            
            if (!msalInstance) {
                showError('Authentication library not initialized');
                return;
            }
            
            console.log("[DEBUG] Starting popup authentication");
            msalInstance.loginPopup(requestObj)
                .then(function(loginResponse) {
                    updateStatus('Authentication successful. Loading plans...');
                    console.log("[DEBUG] Login response:", loginResponse);
                    loadPlannerPlans();
                })
                .catch(function(error) {
                    showError('Authentication error: ' + error.message);
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
                            return fetch(`https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and planner/plans/any(p:p/id eq '${selectedPlanId}')`, {
                                headers: {
                                    'Authorization': 'Bearer ' + accessToken
                                }
                            });
                        });
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Groups API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch group: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(groupsData) {
                        console.log("[DEBUG] Groups data received:", groupsData);
                        
                        if (groupsData.value && groupsData.value.length > 0) {
                            const groupId = groupsData.value[0].id;
                            
                            // Get members of the group
                            return getAccessToken().then(accessToken => {
                                return fetch(`https://graph.microsoft.com/v1.0/groups/${groupId}/members`, {
                                    headers: {
                                        'Authorization': 'Bearer ' + accessToken
                                    }
                                });
                            });
                        } else {
                            throw new Error('No group found for this plan');
                        }
                    })
                    .then(function(response) {
                        console.log("[DEBUG] Members API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch members: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(membersData) {
                        console.log("[DEBUG] Members data received:", membersData);
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
                        if (membersData.value && membersData.value.length > 0) {
                            membersData.value.forEach(function(member) {
                                const option = document.createElement('option');
                                option.value = member.id;
                                option.text = member.displayName;
                                assigneeSelect.appendChild(option);
                                console.log("[DEBUG] Added member:", member.displayName);
                            });
                            updateStatus('Assignees loaded successfully: ' + membersData.value.length + ' members found');
                        } else {
                            updateStatus('No members found in this plan');
                            console.log("[DEBUG] No members found in data:", membersData);
                        }
                    })
                    .catch(function(error) {
                        hideElement('loadingIndicator');
                        showError('Error loading assignees: ' + error.message);
                        console.error("Error loading assignees:", error);
                    });
            }
        }

        // Create Planner task
        function createPlannerTask() {
            const plannerSelect = document.getElementById('planSelector');
            const assigneeSelect = document.getElementById('assigneeSelector');
            const titleInput = document.getElementById('taskTitle');
            const descriptionTextarea = document.getElementById('taskDescription');
            const dueDateInput = document.getElementById('dueDate');
            
            if (!plannerSelect || !assigneeSelect || !titleInput || !descriptionTextarea || !dueDateInput) {
                showError('One or more form elements not found');
                return;
            }
            
            const planId = plannerSelect.value;
            const assigneeId = assigneeSelect.value;
            const title = titleInput.value;
            const description = descriptionTextarea.value;
            const dueDate = dueDateInput.value;
            
            console.log("[DEBUG] Creating task with:", { planId, assigneeId, title, dueDate });
            
            if (!planId) {
                showError('Please select a plan');
                return;
            }
            
            if (!title) {
                showError('Please enter a task title');
                return;
            }
            
            updateStatus('Creating task...');
            showElement('loadingIndicator');
            
            // Prepare task data
            const taskData = {
                planId: planId,
                title: title,
                details: {
                    description: description
                }
            };
            
            if (dueDate) {
                const dueDateObj = new Date(dueDate);
                taskData.dueDateTime = dueDateObj.toISOString();
            }
            
            console.log("[DEBUG] Task data:", taskData);
            
            getAccessToken()
                .then(function(accessToken) {
                    // Create task in Planner
                    return fetch('https://graph.microsoft.com/v1.0/planner/tasks', {
                        method: 'POST',
                        headers: {
                            'Authorization': 'Bearer ' + accessToken,
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(taskData)
                    });
                })
                .then(function(response) {
                    console.log("[DEBUG] Create task API response status:", response.status);
                    if (!response.ok) {
                        return response.text().then(text => {
                            console.log("[DEBUG] Error response body:", text);
                            throw new Error('Failed to create task: ' + response.status + ' - ' + text);
                        });
                    }
                    return response.json();
                })
                .then(function(taskData) {
                    console.log("[DEBUG] Task created successfully:", taskData);
                    
                    // If assignee is selected, assign the task
                    if (assigneeId) {
                        return getAccessToken().then(accessToken => {
                            const assignmentData = {
                                assignments: {
                                    [assigneeId]: {
                                        "@odata.type": "#microsoft.graph.plannerAssignment",
                                        "orderHint": " !"
                                    }
                                }
                            };
                            
                            return fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${taskData.id}/assignments`, {
                                method: 'PATCH',
                                headers: {
                                    'Authorization': 'Bearer ' + accessToken,
                                    'Content-Type': 'application/json',
                                    'If-Match': taskData['@odata.etag']
                                },
                                body: JSON.stringify(assignmentData)
                            }).then(() => taskData);
                        });
                    }
                    
                    return taskData;
                })
                .then(function(data) {
                    hideElement('loadingIndicator');
                    updateStatus('Task created successfully!');
                    
                    // Clear form
                    titleInput.value = '';
                    descriptionTextarea.value = '';
                    dueDateInput.value = '';
                    
                    // Show success message
                    showSuccess('Task created successfully in Planner!');
                })
                .catch(function(error) {
                    hideElement('loadingIndicator');
                    showError('Error creating task: ' + error.message);
                    console.error("Error creating task:", error);
                });
        }

        // Helper function to show success message
        function showSuccess(message) {
            const statusMsg = document.getElementById('statusMessage');
            if (statusMsg) {
                statusMsg.textContent = message;
                statusMsg.style.color = '#4CAF50';
                statusMsg.style.fontWeight = 'bold';
                
                // Reset after 3 seconds
                setTimeout(function() {
                    statusMsg.textContent = '';
                    statusMsg.style.color = '';
                    statusMsg.style.fontWeight = '';
                }, 3000);
            }
        }

        // Show element by ID
        function showElement(id) {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = 'block';
                console.log(`[DEBUG] Element ${id} shown`);
            } else {
                console.log(`[DEBUG] Cannot show element ${id} - not found`);
            }
        }

        // Hide element by ID
        function hideElement(id) {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = 'none';
                console.log(`[DEBUG] Element ${id} hidden`);
            } else {
                console.log(`[DEBUG] Cannot hide element ${id} - not found`);
            }
        }

        // Helper function to update status
        function updateStatus(message) {
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = message;
                console.log("[DEBUG] Status updated:", message);
            } else {
                console.log("[DEBUG] Status element not available. Message:", message);
            }
        }

        // Helper function to show error
        function showError(message) {
            console.error("ERROR:", message);
            updateStatus('Error: ' + message);
            
            const statusMsg = document.getElementById('statusMessage');
            if (statusMsg) {
                statusMsg.textContent = 'Error: ' + message;
                statusMsg.style.color = '#F44336';
                statusMsg.style.fontWeight = 'bold';
                
                // Reset after 5 seconds
                setTimeout(function() {
                    statusMsg.textContent = '';
                    statusMsg.style.color = '';
                    statusMsg.style.fontWeight = '';
                }, 5000);
            }
        }
    });
};
