# Outlook to Planner Add-in: Design Document

## 1. Introduction

This document details the design for the Outlook to Microsoft Planner add-in. It builds upon the approved proposal and focuses on providing a clear architectural overview, user flow, and component design to guide the development process. Special attention is given to structuring information in a way that facilitates step-by-step guidance for creation, testing, and implementation, as per the user's request.

## 2. Overall Architecture

The add-in will consist of two main parts:

1.  **Frontend (Outlook Task Pane Add-in):** This is the user-facing component that runs within Outlook. It will be built using HTML, CSS, and JavaScript.
    *   **Office.js:** The Office JavaScript API will be used to interact with the Outlook application context, such as accessing the currently selected email item (subject, body, saving as .eml) and displaying the task pane.
    *   **User Interface (UI):** HTML will define the structure of the task pane, CSS for styling, and JavaScript for dynamic behavior and user interactions (e.g., handling button clicks, updating dropdowns).
2.  **Backend Services (Microsoft Graph API):** This is not a backend we build, but rather Microsoft's existing API that provides access to Microsoft 365 data, including Planner and user information.
    *   **Microsoft Graph API Client:** The JavaScript frontend will make secure calls to the Microsoft Graph API to perform actions like fetching Planner plans, listing plan members, creating tasks, and attaching files.
    *   **Authentication:** User authentication and authorization to access the Graph API will be handled via OAuth 2.0, aiming for a Single Sign-On (SSO) experience using `getAccessToken()` from Office.js where possible, or otherwise guiding the user through a standard OAuth flow.

**Diagrammatic Overview:**

```
+-----------------------+      +-----------------+      +---------------------------+
|   Outlook Client      |----->| Office.js API   |----->|   Add-in Task Pane (JS)   |
| (Desktop, Web, New)   |      +-----------------+      | (HTML, CSS, JavaScript)   |
+-----------------------+                               +-------------+-------------+
                                                                      |
                                                                      v
                                                        +-------------+-------------+
                                                        | Microsoft Graph API Client  |
                                                        | (JavaScript - Fetch/Axios)  |
                                                        +-------------+-------------+
                                                                      |
                                                                      v
                                                        +---------------------------+
                                                        |   Microsoft Graph API     |
                                                        | (Planner, Users, Mail)    |
                                                        +---------------------------+
```

## 3. Component Breakdown (Task Pane UI)

The task pane will be the primary interface for the user. Key components include:

*   **Authentication Handler:**
    *   Manages the sign-in process for Microsoft Graph API.
    *   Handles token acquisition and refresh.
    *   Displays appropriate UI for sign-in prompts or errors.
*   **Plan Selector:**
    *   Dropdown list (`<select>`).
    *   Populated by fetching the user's Planner plans via Graph API (`GET /me/planner/plans`).
    *   Displays plan names.
    *   Triggers loading of plan members when a plan is selected.
*   **Assignee Selector:**
    *   Dropdown list (`<select>`).
    *   Populated by fetching members of the selected plan (e.g., `GET /planner/plans/{plan-id}/members` or `GET /groups/{group-id}/members` if plan is group-owned).
    *   Displays user names.
*   **Due Date Picker:**
    *   HTML date input (`<input type="date">`) or a JavaScript-based calendar component for better UX.
    *   Allows selection of a due date for the task.
*   **Task Title Input:**
    *   Text input field (`<input type="text">`).
    *   Pre-filled with the selected email's subject (editable).
*   **Task Description Textarea:**
    *   Text area (`<textarea>`).
    *   Pre-filled with the selected email's body (editable, potentially with basic formatting or plain text conversion).
*   **Create Task Button:**
    *   Button (`<button>`).
    *   Initiates the process of creating the task in Planner.
*   **Status/Notification Area:**
    *   A `<div>` or `<p>` element to display success messages, error messages, or loading indicators.

## 4. Data Flow

1.  **Email Selection:** User selects an email in Outlook.
2.  **Add-in Activation:** User opens the add-in task pane.
3.  **Email Data Retrieval (Office.js):**
    *   `Office.context.mailbox.item.subject.getAsync()` -> Task Title (default)
    *   `Office.context.mailbox.item.body.getAsync({coercionType: Office.CoercionType.Text})` -> Task Description (default)
    *   `Office.context.mailbox.item.saveAsync()` -> .eml file content (for attachment)
4.  **Authentication (Graph API):**
    *   If not authenticated, user is prompted to sign in.
    *   Access token is obtained.
5.  **Planner Plan Fetch (Graph API):**
    *   `GET /me/planner/plans` (or similar) using the access token.
    *   Response: List of plans (ID, title).
    *   UI: Populate Plan Selector dropdown.
6.  **Plan Member Fetch (Graph API - on plan selection):**
    *   `GET /planner/plans/{selected-plan-id}/members` (or group members).
    *   Response: List of users (ID, displayName).
    *   UI: Populate Assignee Selector dropdown.
7.  **User Input:** User selects assignee, due date, and potentially modifies title/description.
8.  **Task Creation (Graph API):**
    *   `POST /planner/tasks` with payload:
        *   `planId`: Selected plan ID.
        *   `title`: From Task Title Input.
        *   `assignments`: `{ "{user-id}": { "@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !" } }` (selected assignee ID).
        *   `dueDateTime`: Selected due date.
    *   Response: Created task object (including task ID).
9.  **Task Details Update (Description - Graph API):**
    *   `GET /planner/tasks/{new-task-id}/details` to get ETag.
    *   `PATCH /planner/tasks/{new-task-id}/details` with `If-Match` header (ETag) and payload:
        *   `description`: From Task Description Textarea.
10. **Email Attachment (Graph API):**
    *   The .eml file (obtained in step 3) needs to be attached. The exact mechanism depends on Graph API capabilities for direct .eml upload to Planner task references vs. uploading to OneDrive/SharePoint and then linking.
    *   Preferred: `plannerExternalReference` with `alias` = email subject, `previewPriority` = `!`, `type` = `.eml` and `externalUrl` pointing to the file (if it needs to be hosted first, e.g. on OneDrive).
    *   Alternative: If direct .eml upload to a task is supported as a file attachment (not just a link reference), that would be used. Research indicates `plannerTaskDetails` has a `references` property for `plannerExternalReference` objects. For actual file attachments, it often involves SharePoint/OneDrive.
    *   For .eml, the most straightforward approach for Planner is often to add it as a `plannerExternalReference` if the .eml can be accessed via a URL (e.g., after uploading to OneDrive). If `Office.context.mailbox.item.saveAsync` provides a direct stream or base64 content, we might need to upload this to OneDrive first using Graph API (`PUT /me/drive/root:/{filename}.eml:/content`) and then link the resulting OneDrive item to the Planner task.
11. **Feedback:** UI updated with success or error message.

## 5. Authentication Flow (Microsoft Graph API)

We will prioritize **Single Sign-On (SSO)** for the best user experience:

1.  **Attempt SSO:** When the add-in loads, it will call `Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true })`.
2.  **SSO Success:** If successful, the add-in receives an access token for Microsoft Graph. This token can be exchanged for a Graph API token that can be used to call Planner APIs.
3.  **SSO Fallback (Manual OAuth 2.0):** If SSO fails (e.g., user needs to consent, or admin consent is required and not granted, or SSO is not supported by the Outlook client/environment), the add-in will fall back to a standard OAuth 2.0 Authorization Code Flow.
    *   The add-in will redirect the user (or open a dialog using `Office.context.ui.displayDialogAsync`) to the Microsoft identity platform `/authorize` endpoint.
    *   Required Scopes: `User.Read`, `Mail.Read`, `Group.Read.All` (or `Group.ReadWrite.All` if creating plans/groups), `Tasks.ReadWrite`.
    *   User signs in and consents.
    *   Microsoft identity platform redirects back to a pre-registered redirect URI with an authorization code.
    *   The add-in (or a backend service if using one for token exchange, though for a pure client-side add-in, this exchange happens client-side or via a dialog API callback) exchanges the authorization code for an access token and a refresh token at the `/token` endpoint.
4.  **Token Storage:** Access tokens are short-lived. Refresh tokens (if obtained via manual flow) can be securely stored (e.g., `OfficeRuntime.storage` or browser's local storage if appropriate for the flow, though server-side is better for refresh tokens) to obtain new access tokens without requiring the user to sign in again.
5.  **Making API Calls:** The obtained access token is included in the `Authorization` header of all Microsoft Graph API requests (`Bearer <token>`).

**Permissions (Scopes) Required for Microsoft Graph API:**

*   `User.Read`: To get basic user profile information.
*   `Mail.Read`: To read the selected email's subject and body, and to save the email as .eml.
*   `Group.Read.All`: To list groups the user is a member of (if plans are tied to M365 groups and we need to list plans via groups, or to list group members for assignments).
*   `Tasks.ReadWrite`: To read user's Planner plans and tasks, and to create/update tasks.
*   `Files.ReadWrite.All` (or more specific like `Files.ReadWrite.AppFolder`): If needing to upload the .eml to OneDrive/SharePoint before attaching as a link.

## 6. Detailed User Stories / Use Cases

*   **UC1: Create Task from Email (Happy Path)**
    1.  User selects an email.
    2.  User opens the add-in.
    3.  Add-in loads, authenticates user (SSO or prompts if needed).
    4.  Add-in pre-fills task title (email subject) and description (email body).
    5.  User selects a Planner Plan from a populated dropdown.
    6.  User selects an Assignee from a populated dropdown (members of the selected plan).
    7.  User selects a Due Date using the calendar picker.
    8.  User clicks "Create Task".
    9.  Task is created in Planner with title, description, assignee, due date, and the .eml file attached.
    10. User sees a success message.
*   **UC2: Edit Default Task Details**
    *   Before clicking "Create Task", user modifies the pre-filled task title.
    *   Before clicking "Create Task", user modifies the pre-filled task description.
*   **UC3: No Email Selected**
    *   User opens add-in without an email selected.
    *   Add-in displays a message: "Please select an email to create a task."
*   **UC4: Authentication Failure**
    *   User cancels the sign-in prompt.
    *   Add-in displays a message: "Sign-in is required to access Microsoft Planner."
*   **UC5: API Error (e.g., Cannot Load Plans)**
    *   Graph API call to fetch plans fails.
    *   Add-in displays an error: "Could not load your Planner plans. Please try again."
*   **UC6: Task Creation Fails**
    *   Graph API call to create task fails.
    *   Add-in displays an error: "Failed to create task in Planner. Please try again."

## 7. API Interaction Details (Microsoft Graph v1.0)

*   **List User's Planner Plans:**
    *   `GET https://graph.microsoft.com/v1.0/me/planner/plans`
    *   Response: Array of `plannerPlan` objects.
*   **List Group Members (for Assignees, if plan is group-owned):**
    *   `GET https://graph.microsoft.com/v1.0/groups/{group-id-of-plan}/members?$select=id,displayName,userPrincipalName`
    *   Response: Array of `user` objects.
*   **Create Planner Task:**
    *   `POST https://graph.microsoft.com/v1.0/planner/tasks`
    *   Request Body (example):
        ```json
        {
          "planId": "{plan-id}",
          "bucketId": "{bucket-id}", // Optional, can be auto-assigned by Planner
          "title": "Email Subject",
          "assignments": {
            "{assignee-user-id}": {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              "orderHint": " !"
            }
          },
          "dueDateTime": "YYYY-MM-DDTHH:MM:SS.sssZ"
        }
        ```
*   **Get Planner Task Details (for ETag and to update description/references):**
    *   `GET https://graph.microsoft.com/v1.0/planner/tasks/{task-id}/details`
    *   Response: `plannerTaskDetails` object (includes ETag in `@odata.etag`).
*   **Update Planner Task Details (for description):**
    *   `PATCH https://graph.microsoft.com/v1.0/planner/tasks/{task-id}/details`
    *   Headers: `If-Match: "{ETag-from-GET}"`, `Content-Type: application/json`
    *   Request Body:
        ```json
        {
          "description": "Email body content..."
        }
        ```
*   **Add .eml as Reference to Planner Task Details:**
    *   `PATCH https://graph.microsoft.com/v1.0/planner/tasks/{task-id}/details`
    *   Headers: `If-Match: "{ETag-from-GET-or-previous-PATCH}"`, `Content-Type: application/json`
    *   Request Body (assuming .eml is uploaded to OneDrive and `oneDriveFileUrl` is its webLink):
        ```json
        {
          "references": {
            "https://{tenant-name}.sharepoint.com/personal/{user_path_to_eml_file}.eml": {
                 "@odata.type": "#microsoft.graph.plannerExternalReference",
                 "alias": "Original Email: Email Subject.eml",
                 "type": ".eml",
                 "previewPriority": "!"
            }
          }
        }
        ```
    *   Note: The key for the reference is the URL itself.
*   **Upload .eml to OneDrive (if needed before linking):**
    *   `PUT https://graph.microsoft.com/v1.0/me/drive/root:/OutlookTasksAddinAttachments/{email-subject}.eml:/content`
    *   Headers: `Content-Type: message/rfc822` (or `application/octet-stream`)
    *   Request Body: The raw .eml content.
    *   Response: `driveItem` object which includes `webUrl` for linking.

## 8. Manifest File Structure (Conceptual - `manifest.xml`)

Key elements in the `manifest.xml` will include:

*   `<Id>`: Unique GUID for the add-in.
*   `<Version>`: Add-in version.
*   `<ProviderName>` and `<DisplayName>`.
*   `<Description>`.
*   `<IconUrl>`.
*   `<HighResolutionIconUrl>`.
*   `<SupportUrl>`.
*   `<AppDomains>`: To list domains the add-in interacts with (e.g., `graph.microsoft.com`, authentication domains).
*   `<Hosts>`: Specifies Outlook.
*   `<Requirements>`: Specifies minimum Office API requirement sets.
*   `<FormSettings>` (for Task Pane):
    *   `<SourceLocation>`: URL of the HTML page for the task pane.
*   `<Permissions>`: `ReadWriteMailbox` (or more granular if possible, like `ReadItem`).
*   `<WebApplicationInfo>` (for SSO/Graph API access):
    *   `<Id>`: Application (client) ID from Azure AD app registration.
    *   `<Resource>`: e.g., `api://localhost:3000/{client-id}` (if using localhost for dev).
    *   `<Scopes>`: List of required Graph API scopes (e.g., `profile`, `openid`, `User.Read`, `Mail.Read`, `Tasks.ReadWrite`, `Group.Read.All`, `Files.ReadWrite.All`).

If using the unified manifest (`manifest.json`), the structure will be different but convey similar information.

## 9. Error Handling Strategy

*   **UI Feedback:** All errors will be communicated to the user via the Status/Notification Area in the task pane.
*   **Console Logging:** Detailed error information will be logged to the browser console for debugging.
*   **Specific Error Messages:** Where possible, translate API error codes/messages into user-friendly explanations.
*   **Retry Mechanisms:** For transient network errors, a simple "Try Again" button might be offered.
*   **Graceful Degradation:** If a non-critical feature fails (e.g., pre-filling description), the core functionality (task creation) should still be attempted if possible.

## 10. Step-by-Step Guidance Considerations

Throughout the development and documentation phases, we will focus on creating clear, step-by-step instructions for:

*   **Setting up the Development Environment:** (Node.js, npm, Yeoman generator, code editor).
*   **Creating the Add-in Project:** Using `yo office`.
*   **Understanding the Code Structure:** Explaining key files and folders.
*   **Implementing UI Components:** HTML, CSS, and JS for each part of the task pane.
*   **Integrating Office.js:** How to get email data.
*   **Azure AD App Registration:** Step-by-step for setting up the app registration required for Graph API authentication and SSO, including configuring permissions and redirect URIs.
*   **Implementing Authentication:** Code walkthrough for SSO and fallback OAuth.
*   **Calling Microsoft Graph API:** Examples for each required call (Planner, OneDrive if used).
*   **Testing the Add-in:** Sideloading in Outlook (Web, Desktop), debugging techniques.
*   **Preparing for Publishing:** Creating the manifest, gathering AppSource assets.
*   **AppSource Submission Process:** Guiding through Partner Center.

This design document serves as the foundation for these detailed guides.

## 11. Next Steps

With this design in place, the next phase is the development of the Minimum Viable Product (MVP) based on this architecture and user flow.
