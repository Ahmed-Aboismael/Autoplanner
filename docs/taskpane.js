// Updated taskpane.js with MSAL v1.x and popup-based authentication for Outlook desktop
// This version uses a compatible authentication approach for Outlook 365 desktop clients

// Add jQuery reference at the top of the file
// This script should be loaded after jQuery in the taskpane.html
$(document).ready(function() {
    'use strict';

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
    const msalInstance = new Msal.UserAgentApplication(msalConfig);

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
    const statusElement = document.getElementById('status');
    
    // Set initial status
    updateStatus('Initializing add-in...');

    // Initialize Office.js
    Office.onReady(function(info) {
        if (info.host === Office.HostType.Outlook) {
            updateStatus('Office.js initialized. Preparing authentication...');
            
            // Initialize UI elements
            $('#authenticate-button').click(authenticateWithPopup);
            $('#create-task-button').click(createPlannerTask);
            $('#planner-selection').change(onPlannerSelectionChange);
            
            // Load email data
            loadEmailData();
            
            // Check if user is already signed in
            if (msalInstance.getAccount()) {
                updateStatus('User already signed in. Loading plans...');
                loadPlannerPlans();
            } else {
                updateStatus('Please sign in to access your Planner plans');
                $('#authenticate-section').show();
                $('#task-form-section').hide();
            }
        } else {
            updateStatus('This add-in only works in Outlook');
        }
    });

    // Authenticate using popup for Outlook desktop
    function authenticateWithPopup() {
        updateStatus('Authenticating...');
        
        msalInstance.loginPopup(requestObj)
            .then(function(loginResponse) {
                updateStatus('Authentication successful. Loading plans...');
                $('#authenticate-section').hide();
                $('#task-form-section').show();
                loadPlannerPlans();
            })
            .catch(function(error) {
                updateStatus('Authentication error: ' + error.message);
                console.error('Authentication error:', error);
            });
    }

    // Get access token for Microsoft Graph API
    function getAccessToken() {
        return msalInstance.acquireTokenSilent(requestObj)
            .then(function(tokenResponse) {
                return tokenResponse.accessToken;
            })
            .catch(function(error) {
                if (error.name === "InteractionRequiredAuthError") {
                    return msalInstance.acquireTokenPopup(requestObj)
                        .then(function(tokenResponse) {
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
                // Get all plans the user has access to
                return fetch('https://graph.microsoft.com/v1.0/me/planner/plans', {
                    headers: {
                        'Authorization': 'Bearer ' + accessToken
                    }
                });
            })
            .then(function(response) {
                if (!response.ok) {
                    throw new Error('Failed to fetch plans: ' + response.status);
                }
                return response.json();
            })
            .then(function(data) {
                if (data.value && data.value.length > 0) {
                    updateStatus('Plans loaded successfully');
                    
                    // Clear existing options
                    $('#planner-selection').empty();
                    
                    // Add default option
                    $('#planner-selection').append($('<option>', {
                        value: '',
                        text: '-- Select a plan --'
                    }));
                    
                    // Add plans to dropdown
                    data.value.forEach(function(plan) {
                        $('#planner-selection').append($('<option>', {
                            value: plan.id,
                            text: plan.title
                        }));
                    });
                    
                    // Show the form
                    $('#task-form-section').show();
                } else {
                    updateStatus('No plans found. Please create a plan in Microsoft Planner first.');
                }
            })
            .catch(function(error) {
                updateStatus('Error loading plans: ' + error.message);
                console.error('Error loading plans:', error);
                $('#authenticate-section').show();
            });
    }

    // Handle plan selection change
    function onPlannerSelectionChange() {
        const selectedPlanId = $('#planner-selection').val();
        
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
                    if (!response.ok) {
                        throw new Error('Failed to fetch buckets: ' + response.status);
                    }
                    return response.json();
                })
                .then(function(data) {
                    // Clear existing options
                    $('#bucket-selection').empty();
                    
                    // Add default option
                    $('#bucket-selection').append($('<option>', {
                        value: '',
                        text: '-- Select a bucket --'
                    }));
                    
                    // Add buckets to dropdown
                    if (data.value && data.value.length > 0) {
                        data.value.forEach(function(bucket) {
                            $('#bucket-selection').append($('<option>', {
                                value: bucket.id,
                                text: bucket.name
                            }));
                        });
                        updateStatus('Buckets loaded successfully');
                    } else {
                        updateStatus('No buckets found in this plan');
                    }
                })
                .catch(function(error) {
                    updateStatus('Error loading buckets: ' + error.message);
                    console.error('Error loading buckets:', error);
                });
        }
    }

    // Load email data
    function loadEmailData() {
        updateStatus('Loading email data...');
        
        Office.context.mailbox.item.subject.getAsync(function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                $('#task-title').val(result.value);
            } else {
                console.error('Error getting email subject:', result.error);
            }
        });
        
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                $('#task-description').val(result.value);
            } else {
                console.error('Error getting email body:', result.error);
            }
        });
        
        updateStatus('Email data loaded');
    }

    // Create Planner task
    function createPlannerTask() {
        const planId = $('#planner-selection').val();
        const bucketId = $('#bucket-selection').val();
        const title = $('#task-title').val();
        const description = $('#task-description').val();
        const dueDate = $('#due-date').val();
        
        if (!planId) {
            updateStatus('Please select a plan');
            return;
        }
        
        if (!title) {
            updateStatus('Please enter a task title');
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
                if (!response.ok) {
                    throw new Error('Failed to create task: ' + response.status);
                }
                return response.json();
            })
            .then(function(data) {
                updateStatus('Task created successfully!');
                
                // Clear form
                $('#task-title').val('');
                $('#task-description').val('');
                $('#due-date').val('');
                
                // Show success message
                $('#success-message').show().delay(3000).fadeOut();
            })
            .catch(function(error) {
                updateStatus('Error creating task: ' + error.message);
                console.error('Error creating task:', error);
            });
    }

    // Helper function to update status
    function updateStatus(message) {
        if (statusElement) {
            statusElement.textContent = message;
        }
        console.log(message);
    }
});
