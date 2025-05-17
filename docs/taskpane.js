// Simplified taskpane.js with minimal functionality
// This version focuses on proper Office.js initialization and basic email data loading

// Use Office.initialize instead of Office.onReady to ensure proper loading sequence
Office.initialize = function (reason) {
    // This function is called when the Office.js library is fully loaded
    console.log("[DEBUG] Office.initialize called with reason:", reason);
    
    // Set up the UI once Office.js is fully initialized
    $(document).ready(function() {
        console.log("[DEBUG] Document ready event fired");
        
        // Set up status message area
        const statusElement = document.getElementById('statusMessage');
        if (statusElement) {
            statusElement.textContent = "Add-in initialized successfully";
        }
        
        // Load email data (subject and body)
        try {
            console.log("[DEBUG] Getting email subject");
            Office.context.mailbox.item.subject.getAsync(function(result) {
                console.log("[DEBUG] Subject result:", result);
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const titleInput = document.getElementById('taskTitle');
                    if (titleInput) {
                        titleInput.value = result.value;
                        console.log("[DEBUG] Email subject loaded:", result.value);
                    } else {
                        console.error("taskTitle element not found");
                    }
                } else {
                    console.error("Error getting email subject:", result.error);
                }
            });
            
            console.log("[DEBUG] Getting email body");
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
                console.log("[DEBUG] Body result:", result);
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const descriptionTextarea = document.getElementById('taskDescription');
                    if (descriptionTextarea) {
                        descriptionTextarea.value = result.value;
                        console.log("[DEBUG] Email body loaded, length:", result.value.length);
                    } else {
                        console.error("taskDescription element not found");
                    }
                } else {
                    console.error("Error getting email body:", result.error);
                }
            });
        } catch (error) {
            console.error("Error accessing email data:", error);
        }
        
        // Set up create task button with minimal functionality
        const createButton = document.getElementById('createTaskButton');
        if (createButton) {
            createButton.onclick = function() {
                const statusMsg = document.getElementById('statusMessage');
                if (statusMsg) {
                    statusMsg.textContent = "Create task button clicked - functionality will be added in next version";
                    statusMsg.style.color = '#4CAF50';
                }
                console.log("[DEBUG] Create task button clicked");
            };
        } else {
            console.error("createTaskButton not found");
        }
    });
};
