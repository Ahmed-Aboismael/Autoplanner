# Proposal: Outlook to Microsoft Planner Add-in

## 1. Introduction

This document outlines the proposal for developing an Outlook add-in that allows users to create tasks in Microsoft Planner directly from their emails. The add-in is envisioned as a micro-SaaS product, available for a one-time purchase on the Microsoft AppSource marketplace.

## 2. User Requirements Summary

Based on our discussions, the key requirements for the add-in are as follows:

*   **Product Type:** Outlook Add-in for creating Microsoft Planner tasks from emails.
*   **Monetization:** One-time purchase.
*   **Target Users:** Public (individual users and teams).
*   **Core Functionality:**
    *   The add-in should activate when an email is selected in Outlook and be accessible via a button/tab.
    *   It should allow the user to select one of their accessible Microsoft Planner plans.
    *   It should allow the user to select an assignee from the members of the currently selected Planner plan.
    *   A calendar picker should be provided for setting the task's due date.
    *   The subject of the selected email will be used as the default title for the Planner task.
    *   The body of the selected email will be used as the default description for the Planner task.
    *   A copy of the selected email should be attached to the created Planner task as an .eml file.
*   **Error Handling:** The add-in should not attempt to process if no email is selected or open. Clear feedback should be provided for other potential issues (e.g., connectivity problems, permission issues, task creation failures).
*   **Technology:** No specific preferences were stated; standard web technologies (JavaScript, HTML, CSS) will be used for Outlook add-in development, leveraging the Office JavaScript API and Microsoft Graph API.
*   **Marketplace Familiarity:** The user is not familiar with the Microsoft AppSource submission process; guidance will be provided.

## 3. Proposed Solution

We propose to build a robust and user-friendly Outlook add-in that seamlessly integrates with Microsoft Planner.

### 3.1. Technical Approach

*   **Outlook Add-in Framework:** The add-in will be developed as a Task Pane add-in for Microsoft Outlook. This will be built using modern web technologies:
    *   **HTML5:** For the structure of the add-in's user interface.
    *   **CSS3:** For styling the user interface to ensure a clean and professional look.
    *   **JavaScript (ES6+):** For the client-side logic, event handling, and interaction with Office.js and Microsoft Graph API.
    *   **Office JavaScript API (Office.js):** To interact with the Outlook application, access details of the currently selected email (subject, body, save as .eml), and manage the add-in's lifecycle.
*   **Microsoft Planner Integration:** Interaction with Microsoft Planner will be achieved using the **Microsoft Graph API**.
    *   **Authentication:** Secure authentication will be implemented using OAuth 2.0 (specifically, the authorization code flow or a suitable flow for Office Add-ins like Single Sign-On (SSO) if applicable and preferred for better user experience) to obtain access tokens for the Microsoft Graph API. This ensures that the add-in only accesses data the user has permitted.
    *   **API Endpoints:** The add-in will utilize various Graph API endpoints to:
        *   List the user's Planner plans (`/me/planner/plans` or `/groups/{group-id}/planner/plans`).
        *   List members of a selected plan (via group members if plans are group-owned: `/groups/{group-id}/members`).
        *   Create a new task in a specified plan (`/planner/tasks`).
        *   Update task details (e.g., assignments, due date, description).
        *   Upload the .eml file as an attachment to the Planner task (`/planner/tasks/{task-id}/details` and managing `references` or using dedicated attachment uploads if available for .eml files directly, potentially by first uploading to OneDrive/SharePoint and linking).
*   **Email Attachment (.eml):** The Office.js API provides methods to save the currently selected email item. We will use `item.saveAsync` or similar to get the raw EML content or an EML file representation which can then be attached to the Planner task via the Graph API.

### 3.2. User Flow

The intended user experience is designed to be intuitive and efficient:

1.  **Select Email:** The user selects an email in their Microsoft Outlook client (Desktop, Web, or new Outlook on Windows).
2.  **Activate Add-in:** The user clicks the add-in button (e.g., in the Outlook ribbon or message action bar).
3.  **Task Pane Opens:** The add-in's task pane appears.
4.  **Authentication (First-time/If Required):** If it's the first time the user is using the add-in or if their session has expired, they will be prompted to sign in with their Microsoft 365 account to grant necessary permissions (e.g., `Mail.Read`, `Tasks.ReadWrite`, `Group.Read.All`, `User.Read`). We will aim for a Single Sign-On (SSO) experience where possible.
5.  **Load Planner Plans:** The add-in fetches and displays a dropdown list of the user's accessible Microsoft Planner plans.
6.  **Select Plan:** The user selects the desired Planner plan from the dropdown.
7.  **Load Plan Members:** Upon plan selection, the add-in fetches and displays a dropdown list of members belonging to that plan (who can be assigned tasks).
8.  **Select Assignee:** The user selects an assignee from the list.
9.  **Set Due Date:** The user selects a due date for the task using an interactive calendar picker.
10. **Review Task Details:**
    *   The **Task Title** field will be pre-populated with the subject of the selected email (user can edit this).
    *   The **Task Description** field will be pre-populated with the body of the selected email (user can edit this).
11. **Create Task:** The user clicks a "Create Task" button.
12. **Processing:** The add-in communicates with the Microsoft Graph API to:
    *   Create the new task in the selected Planner plan with the specified title, description, assignee, and due date.
    *   Retrieve the selected email as an .eml file.
    *   Attach the .eml file to the newly created Planner task.
13. **Feedback:** The user receives a confirmation message in the task pane indicating successful task creation (with a link to the task if possible) or an error message if something went wrong.

### 3.3. UI Description (Conceptual)

The task pane will feature a clean, modern, and intuitive design, consistent with Microsoft Office aesthetics.

*   **Header:** Add-in title (e.g., "Email to Planner Task").
*   **Planner Plan Selection:** A dropdown menu labeled "Select Planner Plan".
*   **Assignee Selection:** A dropdown menu labeled "Assign To" (populated after a plan is selected).
*   **Due Date Selection:** A date input field with an associated calendar picker, labeled "Due Date".
*   **Task Title:** A single-line text input field labeled "Task Title", pre-filled with the email subject.
*   **Task Description:** A multi-line text area labeled "Description", pre-filled with the email body.
*   **Action Button:** A primary button, e.g., "Create Task in Planner".
*   **Status Area:** A space below the button to display success messages, error messages, or loading indicators.

### 3.4. Error Handling and User Feedback

Robust error handling and clear user feedback are crucial:

*   **No Email Selected:** The add-in button might be contextually disabled, or if activated without an email context, it will display a message like "Please select an email to create a task."
*   **Authentication Errors:** Clear prompts for login or re-authentication. Messages for permission-denied errors.
*   **API/Network Errors:** User-friendly messages like "Could not connect to Microsoft Planner. Please check your internet connection and try again." or "Failed to load plans/members."
*   **Task Creation Failures:** Specific messages if possible (e.g., "Invalid due date," "Failed to attach email"), or a general failure message with advice to retry.
*   **Success Notifications:** Clear confirmation upon successful task creation, ideally with a link to the newly created task in Planner.
*   **Loading Indicators:** Visual cues (e.g., spinners) when data is being fetched or processed.

## 4. Estimated Development Timeline

This is a preliminary estimate and may be subject to change based on detailed design specifications and unforeseen complexities.

*   **Phase 1: Detailed Design & UI/UX Prototyping (1-2 Weeks)**
    *   Finalize UI mockups and detailed user flow.
    *   Set up the development environment and initial project structure.
*   **Phase 2: Core Add-in Development (3-4 Weeks)**
    *   Implement the basic Outlook add-in structure and UI components.
    *   Develop email data extraction (subject, body, .eml file generation).
    *   Implement Microsoft Graph API authentication (OAuth 2.0/SSO).
    *   Integrate Planner functionalities: list plans, list members, create tasks, set due dates, assign users.
    *   Implement .eml file attachment to Planner tasks.
*   **Phase 3: Testing, Refinement & Bug Fixing (1-2 Weeks)**
    *   Thorough internal testing across different Outlook clients (Web, Desktop).
    *   Address bugs and refine user experience based on testing.
    *   Implement comprehensive error handling and user feedback mechanisms.
*   **Phase 4: Documentation & AppSource Preparation (1 Week)**
    *   Prepare user documentation (how to install and use).
    *   Prepare materials for AppSource submission (manifest, descriptions, screenshots, privacy policy).

**Total Estimated Time:** Approximately 6-9 weeks.

## 5. AppSource Publishing Overview

Publishing to Microsoft AppSource involves several steps. We will provide guidance throughout this process:

1.  **Partner Center Account:** You will need to have or create a Microsoft Partner Center account.
2.  **Commercial Marketplace Program:** Enroll in the Commercial Marketplace program within Partner Center.
3.  **AppSource Validation Policies:** The add-in must comply with Microsoft's AppSource validation policies. These cover aspects like security, functionality, user experience, privacy, and supportability. We will develop the add-in with these policies in mind.
4.  **Prepare Submission Materials:** This includes:
    *   **Manifest File:** An XML file describing the add-in (ID, version, permissions, capabilities, icons, etc.).
    *   **Store Listing:** Marketing text (title, short description, long description), screenshots, video (optional), icons/logos.
    *   **Privacy Policy:** A publicly accessible URL for your add-in's privacy policy.
    *   **Support Information:** A way for users to get support (e.g., support email, website).
    *   **Terms of Use:** A publicly accessible URL for your add-in's terms of use.
5.  **Submit through Partner Center:** Upload the manifest and all required materials to your Partner Center account.
6.  **Validation Process:** Microsoft will review your submission. This can take several days to a few weeks. They may provide feedback if issues are found, requiring resubmission.
7.  **Approval and Publishing:** Once approved, your add-in will be published to Microsoft AppSource and become available in the Office Store within Outlook.

We will ensure the add-in is developed to meet these requirements and can assist in preparing the necessary documentation and assets for a smooth submission process.

## 6. Monetization

As per your requirement, the add-in will be configured for a **one-time purchase** through the Microsoft AppSource marketplace. Partner Center provides tools to set up pricing and manage transactions.

## 7. Next Steps

We request your review and approval of this proposal. Upon approval, we will proceed with the detailed design and development of the Outlook to Microsoft Planner add-in as outlined.

We are excited about the potential of this add-in and look forward to working with you to bring it to life!
