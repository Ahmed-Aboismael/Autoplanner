// Enhanced debug version of taskpane.js with detailed logging for data loading issues
// This version includes extensive error checking and visual feedback

// Add jQuery reference at the top of the file
// This script should be loaded after jQuery in the taskpane.html
$(document).ready(function() {
    'use strict';

    // Debug flag - set to true for detailed console logging
    const DEBUG = true;

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

    // Status display element
    let statusElement = document.getElementById('status');
    if (!statusElement) {
        debugLog("Status element not found, creating one");
        statusElement = document.createElement('div');
        statusElement.id = 'status';
        statusElement.style.margin = '10px 0';
        statusElement.style.padding = '5px';
        statusElement.style.border = '1px solid #ccc';
        statusElement.style.backgroundColor = '#f8f8f8';
        document.body.insertBefore(statusElement, document.body.firstChild);
    }
    
    // Error display element
    let errorElement = document.getElementById('error');
    if (!errorElement) {
        debugLog("Error element not found, creating one");
        errorElement = document.createElement('div');
        errorElement.id = 'error';
        errorElement.style.margin = '10px 0';
        errorElement.style.padding = '5px';
        errorElement.style.border = '1px solid #f88';
        errorElement.style.backgroundColor = '#fee';
        errorElement.style.color = '#c00';
        errorElement.style.display = 'none';
        document.body.insertBefore(errorElement, statusElement.nextSibling);
    }
    
    // Set initial status
    updateStatus('Initializing add-in...');

    // Initialize Office.js
    Office.onReady(function(info) {
        debugLog("Office.onReady called with info:", info);
        
        if (info.host === Office.HostType.Outlook) {
            updateStatus('Office.js initialized in Outlook. Preparing authentication...');
            
            // Check if UI elements exist and create event handlers
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
                showElement('authenticate-section');
                hideElement('task-form-section');
            }
        } else {
            showError('This add-in only works in Outlook');
        }
    });

    // Setup UI elements and event handlers
    function setupUIElements() {
        debugLog("Setting up UI elements");
        
        // Check authenticate button
        const authButton = document.getElementById('authenticate-button');
        if (authButton) {
            authButton.onclick = authenticateWithPopup;
            debugLog("Auth button handler set");
        } else {
            showError("authenticate-button not found in HTML");
        }
        
        // Check create task button
        const createButton = document.getElementById('create-task-button');
        if (createButton) {
            createButton.onclick = createPlannerTask;
            debugLog("Create task button handler set");
        } else {
            showError("create-task-button not found in HTML");
        }
        
        // Check planner selection dropdown
        const plannerSelect = document.getElementById('planner-selection');
        if (plannerSelect) {
            plannerSelect.onchange = onPlannerSelectionChange;
            debugLog("Planner selection handler set");
        } else {
            showError("planner-selection not found in HTML");
        }
        
        // Check other required elements
        checkElementExists('task-title', 'Task title input');
        checkElementExists('task-description', 'Task description textarea');
        checkElementExists('bucket-selection', 'Bucket selection dropdown');
        checkElementExists('due-date', 'Due date input');
        checkElementExists('task-form-section', 'Task form section');
        checkElementExists('authenticate-section', 'Authentication section');
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
                hideElement('authenticate-section');
                showElement('task-form-section');
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
                
                const plannerSelect = document.getElementById('planner-selection');
                if (!plannerSelect) {
                    throw new Error('planner-selection element not found');
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
                    
                    // Show the form
                    showElement('task-form-section');
                } else {
                    updateStatus('No plans found. Please create a plan in Microsoft Planner first.');
                    debugLog("No plans found in data:", data);
                }
            })
            .catch(function(error) {
                showError('Error loading plans: ' + error.message);
                debugLog("Error loading plans:", error);
                showElement('authenticate-section');
            });
    }

    // Handle plan selection change
    function onPlannerSelectionChange() {
        const plannerSelect = document.getElementById('planner-selection');
        if (!plannerSelect) {
            showError('planner-selection element not found');
            return;
        }
        
        const selectedPlanId = plannerSelect.value;
        debugLog("Plan selection changed to:", selectedPlanId);
        
        if (selectedPlanId) {
            updateStatus('Loading buckets for selected plan...');
            
            getAccessToken()
                .then(function(accessToken) {
                    // Get buckets for the selected plan
                    return fetch(`https://graph.microsoft.com/v1.0/planner/plans/${selectedPlanId}/buckets`, {
                        headers: {
                            'Authorization': 'Bearer ' + accessToken
                        }
                    });
                })
                .then(function(response) {
                    debugLog("Buckets API response status:", response.status);
                    if (!response.ok) {
                        throw new Error('Failed to fetch buckets: ' + response.status);
                    }
                    return response.json();
                })
                .then(function(data) {
                    debugLog("Buckets data received:", data);
                    
                    const bucketSelect = document.getElementById('bucket-selection');
                    if (!bucketSelect) {
                        throw new Error('bucket-selection element not found');
                    }
                    
                    // Clear existing options
                    bucketSelect.innerHTML = '';
                    
                    // Add default option
                    const defaultOption = document.createElement('option');
                    defaultOption.value = '';
                    defaultOption.text = '-- Select a bucket --';
                    bucketSelect.appendChild(defaultOption);
                    
                    // Add buckets to dropdown
                    if (data.value && data.value.length > 0) {
                        data.value.forEach(function(bucket) {
                            const option = document.createElement('option');
                            option.value = bucket.id;
                            option.text = bucket.name;
                            bucketSelect.appendChild(option);
                            debugLog("Added bucket:", bucket.name);
                        });
                        updateStatus('Buckets loaded successfully: ' + data.value.length + ' buckets found');
                    } else {
                        updateStatus('No buckets found in this plan');
                        debugLog("No buckets found in data:", data);
                    }
                })
                .catch(function(error) {
                    showError('Error loading buckets: ' + error.message);
                    debugLog("Error loading buckets:", error);
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
                    const titleInput = document.getElementById('task-title');
                    if (titleInput) {
                        titleInput.value = result.value;
                        debugLog("Email subject loaded:", result.value);
                    } else {
                        showError('task-title element not found');
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
                    const descriptionTextarea = document.getElementById('task-description');
                    if (descriptionTextarea) {
                        descriptionTextarea.value = result.value;
                        debugLog("Email body loaded, length:", result.value.length);
                    } else {
                        showError('task-description element not found');
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
        const plannerSelect = document.getElementById('planner-selection');
        const bucketSelect = document.getElementById('bucket-selection');
        const titleInput = document.getElementById('task-title');
        const descriptionTextarea = document.getElementById('task-description');
        const dueDateInput = document.getElementById('due-date');
        
        if (!plannerSelect || !bucketSelect || !titleInput || !descriptionTextarea || !dueDateInput) {
            showError('One or more form elements not found');
            return;
        }
        
        const planId = plannerSelect.value;
        const bucketId = bucketSelect.value;
        const title = titleInput.value;
        const description = descriptionTextarea.value;
        const dueDate = dueDateInput.value;
        
        debugLog("Creating task with:", { planId, bucketId, title, dueDate });
        
        if (!planId) {
            showError('Please select a plan');
            return;
        }
        
        if (!title) {
            showError('Please enter a task title');
            return;
        }
        
        updateStatus('Creating task...');
        
        // Prepare task data
        const taskData = {
            planId: planId,
            title: title,
            description: description
        };
        
        if (bucketId) {
            taskData.bucketId = bucketId;
        }
        
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
            .then(function(data) {
                debugLog("Task created successfully:", data);
                updateStatus('Task created successfully!');
                
                // Clear form
                titleInput.value = '';
                descriptionTextarea.value = '';
                dueDateInput.value = '';
                
                // Show success message
                const successElement = document.getElementById('success-message');
                if (successElement) {
                    successElement.style.display = 'block';
                    setTimeout(function() {
                        successElement.style.display = 'none';
                    }, 3000);
                } else {
                    // Create a temporary success message
                    const tempSuccess = document.createElement('div');
                    tempSuccess.style.margin = '10px 0';
                    tempSuccess.style.padding = '5px';
                    tempSuccess.style.border = '1px solid #8f8';
                    tempSuccess.style.backgroundColor = '#efe';
                    tempSuccess.style.color = '#080';
                    tempSuccess.textContent = 'Task created successfully!';
                    document.body.insertBefore(tempSuccess, statusElement.nextSibling);
                    setTimeout(function() {
                        tempSuccess.remove();
                    }, 3000);
                }
            })
            .catch(function(error) {
                showError('Error creating task: ' + error.message);
                debugLog("Error creating task:", error);
            });
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
        if (errorElement) {
            errorElement.textContent = message;
            errorElement.style.display = 'block';
            setTimeout(function() {
                errorElement.style.display = 'none';
            }, 10000); // Hide after 10 seconds
        }
        updateStatus('Error: ' + message);
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
});
