// Fixed taskpane.js with correct element IDs matching the HTML
// This version ensures proper Office.js initialization and DOM element targeting

// Wait for DOM to be ready
$(document).ready(function() {
    'use strict';

    // Debug flag - set to true for detailed console logging
    const DEBUG = true;

    // Status display element
    let statusElement = document.getElementById('statusMessage');
    
    // Set initial status
    updateStatus('Initializing add-in...');

    // IMPORTANT: Wait for Office.js to be fully loaded before doing anything
    Office.onReady(function(info) {
        debugLog("Office.onReady called with info:", info);
        
        if (info.host === Office.HostType.Outlook) {
            updateStatus('Office.js initialized in Outlook. Preparing authentication...');
            
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
                debugLog("MSAL initialized successfully");
            } catch (error) {
                showError("Failed to initialize MSAL: " + error.message);
                debugLog("MSAL initialization error:", error);
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
            
            // Setup UI elements and event handlers
            setupUIElements();
            
            // Load email data
            loadEmailData();
            
            // Check if user is already signed in
            if (msalInstance && msalInstance.getAccount()) {
                const account = msalInstance.getAccount();
                updateStatus('User already signed in as: ' + account.userName + '. Loading plans...');
                debugLog("User account:", account);
                loadPlannerPlans();
            } else {
                updateStatus('Please sign in to access your Planner plans');
                // Auto-authenticate since we don't have an explicit auth button
                authenticateWithPopup();
            }

            // Setup UI elements and event handlers
            function setupUIElements() {
                debugLog("Setting up UI elements");
                
                // Set up create task button
                const createButton = document.getElementById('createTaskButton');
                if (createButton) {
                    createButton.onclick = createPlannerTask;
                    debugLog("Create task button handler set");
                } else {
                    showError("createTaskButton not found in HTML");
                }
                
                // Set up planner selection dropdown
                const plannerSelect = document.getElementById('planSelector');
                if (plannerSelect) {
                    plannerSelect.onchange = onPlannerSelectionChange;
                    debugLog("Planner selection handler set");
                } else {
                    showError("planSelector not found in HTML");
                }
                
                // Check other required elements
                checkElementExists('taskTitle', 'Task title input');
                checkElementExists('taskDescription', 'Task description textarea');
                checkElementExists('assigneeSelector', 'Assignee selection dropdown');
                checkElementExists('dueDate', 'Due date input');
                checkElementExists('content', 'Content section');
                checkElementExists('loadingIndicator', 'Loading indicator');
            }

            // Authenticate using popup for Outlook desktop
            function authenticateWithPopup() {
                updateStatus('Authenticating...');
                
                if (!msalInstance) {
                    showError('Authentication library not initialized');
                    return;
                }
                
                debugLog("Starting popup authentication");
                msalInstance.loginPopup(requestObj)
                    .then(function(loginResponse) {
                        updateStatus('Authentication successful. Loading plans...');
                        debugLog("Login response:", loginResponse);
                        loadPlannerPlans();
                    })
                    .catch(function(error) {
                        showError('Authentication error: ' + error.message);
                        debugLog("Authentication error:", error);
                    });
            }

            // Get access token for Microsoft Graph API
            function getAccessToken() {
                debugLog("Getting access token");
                
                if (!msalInstance) {
                    return Promise.reject(new Error('Authentication library not initialized'));
                }
                
                return msalInstance.acquireTokenSilent(requestObj)
                    .then(function(tokenResponse) {
                        debugLog("Token acquired silently");
                        return tokenResponse.accessToken;
                    })
                    .catch(function(error) {
                        debugLog("Error in silent token acquisition:", error);
                        if (error.name === "InteractionRequiredAuthError") {
                            debugLog("Interaction required, using popup");
                            return msalInstance.acquireTokenPopup(requestObj)
                                .then(function(tokenResponse) {
                                    debugLog("Token acquired via popup");
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
                        debugLog("Access token obtained, length:", accessToken.length);
                        // Get all plans the user has access to
                        updateStatus('Fetching plans from Microsoft Graph API...');
                        return fetch('https://graph.microsoft.com/v1.0/me/planner/plans', {
                            headers: {
                                'Authorization': 'Bearer ' + accessToken
                            }
                        });
                    })
                    .then(function(response) {
                        debugLog("Plans API response status:", response.status);
                        if (!response.ok) {
                            throw new Error('Failed to fetch plans: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(data) {
                        debugLog("Plans data received:", data);
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
                                debugLog("Added plan:", plan.title);
                            });
                        } else {
                            updateStatus('No plans found. Please create a plan in Microsoft Planner first.');
                            debugLog("No plans found in data:", data);
                        }
                    })
                    .catch(function(error) {
                        hideElement('loadingIndicator');
                        showError('Error loading plans: ' + error.message);
                        debugLog("Error loading plans:", error);
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
                debugLog("Plan selection changed to:", selectedPlanId);
                
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
                            debugLog("Plan details API response status:", response.status);
                            if (!response.ok) {
                                throw new Error('Failed to fetch plan details: ' + response.status);
                            }
                            return response.json();
                        })
                        .then(function(planDetails) {
                            debugLog("Plan details received:", planDetails);
                            
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
                            debugLog("Groups API response status:", response.status);
                            if (!response.ok) {
                                throw new Error('Failed to fetch group: ' + response.status);
                            }
                            return response.json();
                        })
                        .then(function(groupsData) {
                            debugLog("Groups data received:", groupsData);
                            
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
                            debugLog("Members API response status:", response.status);
                            if (!response.ok) {
                                throw new Error('Failed to fetch members: ' + response.status);
                            }
                            return response.json();
                        })
                        .then(function(membersData) {
                            debugLog("Members data received:", membersData);
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
                                    debugLog("Added member:", member.displayName);
                                });
                                updateStatus('Assignees loaded successfully: ' + membersData.value.length + ' members found');
                            } else {
                                updateStatus('No members found in this plan');
                                debugLog("No members found in data:", membersData);
                            }
                        })
                        .catch(function(error) {
                            hideElement('loadingIndicator');
                            showError('Error loading assignees: ' + error.message);
                            debugLog("Error loading assignees:", error);
                        });
                }
            }

            // Load email data
            function loadEmailData() {
                updateStatus('Loading email data...');
                
                try {
                    debugLog("Getting email subject");
                    Office.context.mailbox.item.subject.getAsync(function(result) {
                        debugLog("Subject result:", result);
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            const titleInput = document.getElementById('taskTitle');
                            if (titleInput) {
                                titleInput.value = result.value;
                                debugLog("Email subject loaded:", result.value);
                            } else {
                                showError('taskTitle element not found');
                            }
                        } else {
                            showError('Error getting email subject: ' + result.error.message);
                            debugLog("Error getting email subject:", result.error);
                        }
                    });
                    
                    debugLog("Getting email body");
                    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
                        debugLog("Body result:", result);
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            const descriptionTextarea = document.getElementById('taskDescription');
                            if (descriptionTextarea) {
                                descriptionTextarea.value = result.value;
                                debugLog("Email body loaded, length:", result.value.length);
                            } else {
                                showError('taskDescription element not found');
                            }
                        } else {
                            showError('Error getting email body: ' + result.error.message);
                            debugLog("Error getting email body:", result.error);
                        }
                    });
                    
                    updateStatus('Email data loading initiated');
                } catch (error) {
                    showError('Error accessing email data: ' + error.message);
                    debugLog("Exception in loadEmailData:", error);
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
                
                debugLog("Creating task with:", { planId, assigneeId, title, dueDate });
                
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
                
                debugLog("Task data:", taskData);
                
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
                        debugLog("Create task API response status:", response.status);
                        if (!response.ok) {
                            return response.text().then(text => {
                                debugLog("Error response body:", text);
                                throw new Error('Failed to create task: ' + response.status + ' - ' + text);
                            });
                        }
                        return response.json();
                    })
                    .then(function(taskData) {
                        debugLog("Task created successfully:", taskData);
                        
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
                        debugLog("Error creating task:", error);
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

            // Check if element exists and log result
            function checkElementExists(id, description) {
                const element = document.getElementById(id);
                if (element) {
                    debugLog(`${description} (${id}) found`);
                    return true;
                } else {
                    showError(`${description} (${id}) not found in HTML`);
                    return false;
                }
            }

            // Show element by ID
            function showElement(id) {
                const element = document.getElementById(id);
                if (element) {
                    element.style.display = 'block';
                    debugLog(`Element ${id} shown`);
                } else {
                    debugLog(`Cannot show element ${id} - not found`);
                }
            }

            // Hide element by ID
            function hideElement(id) {
                const element = document.getElementById(id);
                if (element) {
                    element.style.display = 'none';
                    debugLog(`Element ${id} hidden`);
                } else {
                    debugLog(`Cannot hide element ${id} - not found`);
                }
            }

            // Helper function to update status
            function updateStatus(message) {
                if (statusElement) {
                    statusElement.textContent = message;
                    debugLog("Status updated:", message);
                } else {
                    debugLog("Status element not available. Message:", message);
                }
            }

            // Helper function to show error
            function showError(message) {
                debugLog("ERROR:", message);
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

            // Helper function for debug logging
            function debugLog(message, obj) {
                if (DEBUG) {
                    if (obj !== undefined) {
                        console.log("[DEBUG] " + message, obj);
                    } else {
                        console.log("[DEBUG] " + message);
                    }
                }
            }
        } else {
            console.error('This add-in only works in Outlook');
        }
    });
});
