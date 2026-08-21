# 🇬🇧 Enterprise Project Tracking System (Zoho Projects Clone)
**Enterprise User & Administrator Manual (English Edition)**

---

## 🌟 Table of Contents
1. [Introduction & Getting Started](#1-introduction--getting-started)
2. [Demo Credentials](#2-demo-credentials)
3. [Part 1: User Manual (For Team Members & Managers)](#part-1-user-manual-for-team-members--managers)
   - [3.1 Workspace & Dashboard](#31-workspace--dashboard)
   - [3.2 Task Management](#32-task-management)
   - [3.3 Kanban Board (Drag & Drop)](#33-kanban-board-drag--drop)
   - [3.4 Gantt Chart Timeline](#34-gantt-chart-timeline)
   - [3.5 Milestones & Deliverables](#35-milestones--deliverables)
   - [3.6 Time Tracking & Timesheets](#36-time-tracking--timesheets)
   - [3.7 Comments & File Attachments](#37-comments--file-attachments)
4. [Part 2: Administrator & Executive Manual](#part-2-administrator--executive-manual)
   - [4.1 Executive Dashboard & KPI Analytics](#41-executive-dashboard--kpi-analytics)
   - [4.2 Team Hierarchy & Role Management](#42-team-hierarchy--role-management)
   - [4.3 Password Reset & Credentials Management](#43-password-reset--credentials-management)
   - [4.4 Invoicing & Client Billing System](#44-invoicing--client-billing-system)
   - [4.5 Data Export (Excel & CSV)](#45-data-export-excel--csv)

---

## 1. Introduction & Getting Started
This application is an enterprise-grade project tracking and collaboration platform inspired by **Zoho Projects**. It is designed to empower teams with multi-workspace management, real-time Kanban boards, dynamic Gantt charts, organizational reporting hierarchy, time tracking, and a real-world invoicing engine—all presented in a sleek, responsive Glassmorphic UI with full support for Dark and Light themes.

### 🚀 How to Launch the System:
1. Open Command Prompt or Terminal inside the project root directory.
2. Start the local server by running: `npm start`
3. Open your web browser and navigate to: `http://localhost:3000`

---

## 2. Demo Credentials
When launching the application, you can either click the quick-login demo buttons on the login screen or enter these credentials:

| Role / Access Level | Email Address | Password | Capabilities & Scope |
| :--- | :--- | :--- | :--- |
| **👑 Executive / Admin (CEO)** | `admin@demo.com` | `admin123` | Full access to all features including the Executive Dashboard, Team Hierarchy, Password Reset, Billing/Invoicing, and system configurations. |
| **👔 Project Manager** | `manager@demo.com` | `manager123` | Can create project workspaces, assign tasks, establish phase milestones, and monitor team time entries. |
| **👤 Team Member** | `member@demo.com` | `member123` | Can view assigned tasks, move Kanban cards across columns, log working hours, comment, and upload attachments. |

> [!TIP]
> **Language & Theme Toggles**: 
> - Click the **🇹🇭 TH / 🇬🇧 EN** button in the top navigation bar to switch language interfaces seamlessly.
> - Click the **🌙 / ☀️** icon to toggle between Dark Mode and Light Mode.

---

## Part 1: User Manual (For Team Members & Managers)

### 3.1 Workspace & Dashboard
Upon signing in, you will land on your **Personal Dashboard**, which serves as your central productivity hub:
- **Task Counters**: Displays your total assigned tasks, completed items, and pending deliverables.
- **My Tasks**: Lists immediate action items prioritized by upcoming due dates.
- **Project Progress**: Visualizes completion percentage bars for all workspaces you belong to.
- **Switching Workspaces**: Use the project selector dropdown in the top navbar or click the **📁 Projects** menu in the left sidebar to navigate into a specific project workspace.

### 3.2 Task Management
Inside a project workspace, navigate to the **✅ Tasks** tab to manage action items:
- **Creating Tasks**: Click **`+ Create Task`**. In the modal, define the task title, priority level (`Low`, `Medium`, `High`), assign a team member, set a target due date, link it to a milestone phase, and optionally configure a recurring schedule (`Daily`, `Weekly`, `Monthly`).
- **Editing & Discussing**: Click any task title to open the **Task Details Modal** where you can modify properties, review status, and communicate with colleagues.

### 3.3 Kanban Board (Drag & Drop)
Click the **📋 Board** tab to view tasks organized into three workflow columns: `To Do`, `In Progress`, and `Done`.
- **Drag & Drop Workflow**: Click and hold any task card, then drag and drop it into a different column to update its workflow status instantly. The system persists the change to the database and recalculates project progress metrics automatically.

### 3.4 Gantt Chart Timeline
Click the **📊 Gantt** tab to explore a visual schedule of project phases and deadlines:
- **Zoom Levels**: In the top right corner of the chart, click the view toggles (`Day`, `Week`, or `Month`) to adjust the timeline zoom scale.
- **Task Details**: Hover your mouse over any horizontal timeline bar to inspect the task name, assigned owner, duration, and completion status.

### 3.5 Milestones & Deliverables
Click the **🏁 Milestones** tab to track major project deliverables and phase objectives:
- **Adding Milestones**: Click **`+ Add Milestone`** to define a phase title and deadline date.
- **Completing Deliverables**: Check the circular checkbox next to a milestone to mark it as **Completed**. The progress bar will turn green at 100%, displaying a trophy icon.

### 3.6 Time Tracking & Timesheets
Click the **⏱️ Timesheet** tab to log working hours against specific project tasks for profitability tracking and client billing:
- **Logging Work**: Select the relevant task from the dropdown, input the hours worked, choose the date, and enter descriptive work notes. Click **`Log Time`**.
- **Billing Status**: Logged time entries remain marked as **"Unbilled"** until an Administrator or Project Manager incorporates them into a formal client invoice.

### 3.7 Comments & File Attachments
Click any task title to open the Task Details Modal, which hosts the team's collaboration suite:
- **Discussion Stream (Comments)**: Type messages and click Post to ask questions, report progress, or provide feedback directly beneath the task.
- **File Vault (Attachments)**: Click **`📎 Attach File`** to upload architectural diagrams, specifications, PDF documents, or screenshots, storing them securely within the task context.

---

## Part 2: Administrator & Executive Manual

### 4.1 Executive Dashboard & KPI Analytics
Exclusive to Administrators and Executives (CEO), accessed via **📊 Executive Dashboard** in the left sidebar:
- **KPI Ribbon**: Real-time summary metrics displaying Total Projects, On-Track Projects, Delayed Projects, and Total Budgeted Revenue across the entire organization.
- **Project Health Table**: Evaluates the financial and operational health of every workspace, showing completion percentages, budget utilization, and automated risk flags (`Healthy`, `At Risk`, `Critical`).
- **Team Workload & Performance**: Monitors active tasks versus completed outputs per employee, enabling leadership to balance workloads and prevent staff burnout.

### 4.2 Team Hierarchy & Role Management
Click **👥 Team Management** in the sidebar to administer user accounts and organizational reporting lines:
- **Adding Profiles**: Click **`+ Add Team Member`**, enter their name, email, and assign an initial password.
- **Role Assignment**: Assign one of three access tiers:
  - `Admin`: Full executive access to financial controls, user permissions, and dashboards.
  - `Manager`: Department or project leader capable of creating projects, assigning tasks, and approving timesheets.
  - `Member`: Standard employee access restricted to assigned workspaces and tasks.
- **Reporting Hierarchy (Reports To)**: Select a direct supervisor in the "Manager" dropdown field. This establishes the organizational chain of command, allowing supervisors to delegate tasks and track departmental efficiency.

### 4.3 Password Reset & Credentials Management
Administrators maintain full administrative control over user credentials to assist locked-out personnel and enforce security policies:
- **Method 1: 1-Click Quick Reset**:
  1. In the Team Management table, locate the target user and click **`🔑 Password`** in the actions column.
  2. Click the **`⚡ Quick Set: admin123`** button. The user's password is immediately restored to `admin123`.
  3. Alternatively, type any custom password (minimum 4 characters) in the input box and click Save.
- **Method 2: Profile Update Modal**:
  1. Click **`✏️ Edit`** on any user row.
  2. Enter a new password into the "New Password (leave blank to keep current)" input field and click Save Member.

### 4.4 Invoicing & Client Billing System
An integrated billing engine that transforms project time into client revenue. Access via a project workspace by selecting the **💰 Billing** tab:
- **Automated Invoice Generation**: Click **`+ Create Invoice`**. The system automatically pulls all **Unbilled Timesheets** associated with the project. Check the time entries you wish to bill; the system compiles line items and calculates totals using a standard hourly rate ($50/hr).
- **Discounts & VAT Customization**: Specify discount amounts ($) and tax/VAT percentage rates (e.g., 7%) before saving.
- **Lifecycle Management**: Click the status button on any invoice card to cycle through payment states: `🟡 Unpaid` ➔ `🟢 Paid` ➔ `🔴 Overdue` ➔ `⚪ Cancelled`. Marking an invoice as "Paid" permanently locks all linked timesheet entries as billed.
- **PDF Print & Delivery**: Click an invoice card to open the **Professional Invoice Preview Modal**. This renders an official, branded financial document complete with status watermark badges. Click **`🖨️ Print / Save PDF`** to generate a client-ready document.

### 4.5 Data Export (Excel & CSV)
To support external accounting software and offline data auditing:
- **Exporting Invoices**: In the Billing tab, click **`📥 Export`** to download an `invoices_export.csv` file.
- **Exporting Tasks**: In the Tasks tab, click **`📥 Export`** to download `tasks_export.csv`.
- **Seamless Excel Compatibility**: Every exported CSV file is embedded with a **UTF-8 Byte Order Mark (BOM)** and double-quoted string formatting. When opened in Microsoft Excel or Google Sheets, international characters (including Thai text), client names, and financial figures render flawlessly with correct column alignment.

---
*Engineered by Antigravity AI — Advanced Enterprise Coding Assistant*
