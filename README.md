# INSEE Digital Task Observation System – Full-Stack Prototype

This version converts the original browser-only prototype into a central Node.js/Express application.

## What changed

- Central SQLite database (`data/insee.sqlite`) for users, Area Manager mappings, observations and corrective actions.
- Server-side login using bcrypt password hashing and JWT sessions.
- Admin-only Area Manager Mapping management.
- Centralized mappings shared by all users/devices connected to the server.
- Centralized observation and corrective-action records.
- Automatic email notification when an observation requires corrective action.
- Optional escalation email copied on corrective-action notifications.
- SMTP configuration through `.env`; email credentials are never placed in `index.html`.
- Existing dashboard, observation form, user management and corrective-action screens retained.

## Requirements

- Node.js 18 or newer
- Access to an approved company SMTP service/mailbox

## Installation

1. Extract the ZIP.
2. Open a terminal in this folder.
3. Run:

```bash
npm install
```

4. Copy `.env.example` to `.env`.
5. Set a strong `JWT_SECRET`.
6. Enter the approved SMTP settings.
7. Start the application:

```bash
npm start
```

8. Open:

`http://localhost:3000`

## Demo accounts

- Admin: `admin` / `admin123`
- Safety Manager: `safety` / `safety123`
- Area Manager: `manager` / `manager123`
- Observer: `observer` / `observer123`

**Change these passwords before any real deployment.**

## Microsoft 365 Email Setup

This build uses Microsoft Graph with Microsoft Entra ID application authentication instead of SMTP username/password authentication. Microsoft documents the client-credentials flow for server-to-server applications, and Graph `sendMail` supports the `Mail.Send` application permission. citeturn0search0turn0search3

### 1. Create an App Registration

An authorized Microsoft 365 administrator should open the Microsoft Entra admin center and create an **App registration** for this application. Microsoft documents app registration and credentials here: https://learn.microsoft.com/en-us/graph/auth-register-app-v2. citeturn0search17

Record:
- **Directory (tenant) ID** → `M365_TENANT_ID`
- **Application (client) ID** → `M365_CLIENT_ID`

### 2. Add Microsoft Graph permission

In the App Registration:

`API permissions` → `Add a permission` → `Microsoft Graph` → `Application permissions` → `Mail.Send`

An administrator must grant consent for the tenant. `Mail.Send` is the least-privileged Microsoft Graph application permission listed for `sendMail`. citeturn0search3

### 3. Create a client secret

For this prototype, create a client secret under `Certificates & secrets` and copy the **secret value** immediately. Put it only in the server `.env` file as `M365_CLIENT_SECRET`.

For production, Microsoft recommends stronger credential options such as certificates for confidential applications; keep the credential out of source code and never put it in `index.html`. citeturn0search1turn0search15

### 4. Choose the sending mailbox

Create/use an approved company mailbox such as:

`taskobservation@yourcompany.com`

Set:

```text
M365_SENDER_EMAIL=taskobservation@yourcompany.com
```

The backend sends through Microsoft Graph using:

```text
POST /users/{sender}/sendMail
```

Microsoft documents this endpoint for application permissions. citeturn0search3

### 5. Configure `.env`

Copy `.env.example` to `.env` and enter the values:

```text
PORT=3000
JWT_SECRET=YOUR_LONG_RANDOM_SECRET
M365_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
M365_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
M365_CLIENT_SECRET=YOUR_SECRET_VALUE
M365_SENDER_EMAIL=taskobservation@yourcompany.com
M365_SENDER_NAME=INSEE Digital Task Observation System
```

Do not commit `.env` to Git or distribute it with the website.

### 6. Install and run

```bash
npm install
npm start
```

Then open:

`http://localhost:3000`

### Email flow

```text
INSEE Browser
     |
     | POST observation
     v
Node.js Backend
     |
     | Microsoft Graph access token
     v
Microsoft Entra ID
     |
     v
Microsoft Graph
     |
     | /users/taskobservation@.../sendMail
     v
Microsoft 365 Mailbox
     |
     v
Area Manager / Escalation Manager
```

The browser never receives the client secret. The backend acquires an app-only Graph access token using MSAL Node and calls Microsoft Graph. Microsoft describes client credentials as a server-to-server flow suitable for services running without user interaction. citeturn0search0turn0search8

### Security note

Application `Mail.Send` is powerful. Before production, have your Microsoft 365 administrator apply your organization's Exchange Online application-access controls so the application is restricted to the intended sending mailbox rather than unnecessarily broad mailbox access. Microsoft documents application access controls/RBAC for Exchange Online. citeturn0search14

## Important production hardening

Before company-wide deployment, use HTTPS, a company identity provider such as Microsoft Entra ID where appropriate, a managed database/backup strategy, stronger audit logging, restricted server access, and the organization's approved Microsoft 365/Graph or SMTP authentication method.

The demo credentials and `.env.example` are for development only.
