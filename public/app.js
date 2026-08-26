// ============================================================

// ---- Global Fetch Interceptor for Session Security ----
const originalFetch = window.fetch;
window.fetch = function (url, options = {}) {
    const token = localStorage.getItem('ag_session_token');
    if (token) {
        options.headers = options.headers || {};
        options.headers['Authorization'] = `Bearer ${token}`;
    }
    return originalFetch(url, options).then(async response => {
        if (response.status === 401 && !url.toString().includes('/api/login')) {
            localStorage.removeItem('ag_session_token');
            localStorage.removeItem('ag_current_user');
            currentUser = null;
            const loginScreen = document.querySelector('#login-screen');
            if (loginScreen) loginScreen.classList.remove('hidden');
            const container = document.querySelector('#app-container');
            if (container) container.style.display = 'none';
            // Use fallback text if t() helper is not initialized yet
            const msg = (typeof currentLang !== 'undefined' && currentLang === 'th') 
                ? 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' 
                : 'Session expired. Please sign in again.';
            if (typeof showToast === 'function') showToast(msg);
            else alert(msg);
        }
        return response;
    });
};

// ---- MSAL Configuration (Microsoft Login) ----
const MS_CLIENT_ID = 'YOUR_CLIENT_ID_HERE';

let msalInstance = null;
try {
    if (typeof msal !== 'undefined' && MS_CLIENT_ID !== 'YOUR_CLIENT_ID_HERE') {
        msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: MS_CLIENT_ID,
                authority: 'https://login.microsoftonline.com/common',
                redirectUri: window.location.origin
            },
            cache: { cacheLocation: 'sessionStorage' }
        });
    }
} catch (e) { console.warn('MSAL.js init skipped:', e.message); }

// ---- State ----
let currentUser = null;
let currentWorkspaceId = null;
let currentProjectId = null;
let currentTaskDetailId = null;
let ganttZoom = 'day';
let currentLang = localStorage.getItem('ag_lang') || 'en';
let currentTheme = localStorage.getItem('ag_theme') || 'dark';
let allWorkspacesList = [];

function getCurrentWorkspaceRate() {
    const ws = allWorkspacesList.find(w => w.id === currentWorkspaceId);
    return (ws && ws.hourly_rate) ? parseFloat(ws.hourly_rate) : 50;
}

// ---- Translations (English & Thai) ----
const translations = {
    en: {
        // Login
        login_subtitle: "Sign in to your enterprise workspace",
        login_microsoft: "Sign in with Microsoft",
        login_or: "or sign in with email",
        email_label: "Email Address",
        password_label: "Password",
        sign_in_btn: "Sign In",
        demo_creds: "Demo credentials:",
        ms_hint: "Microsoft login requires Azure AD App Registration.",
        // Sidebar & Nav
        workspace_label: "Workspace",
        menu_dashboard: "Dashboard",
        menu_projects: "Projects",
        menu_settings: "Settings",
        menu_exec_dashboard_short: "Exec",
        menu_team_short: "Team",
        user_role: "Manager",
        notif_title: "Notifications",
        notif_mark_read: "Mark all read",
        // Dashboard
        stat_projects: "Projects",
        stat_active_tasks: "Active Tasks",
        stat_hours: "Logged Hours",
        stat_milestones: "Milestones",
        panel_pending_tasks: "Pending Tasks",
        panel_progress: "Progress",
        th_task: "Task",
        th_project: "Project",
        th_start: "Start Date (optional)",
        th_due: "Due",
        th_priority: "Priority",
        th_status: "Status",
        th_date: "Date",
        th_desc: "Description",
        th_user: "User",
        th_hours: "Hours",
        // Projects Screen
        all_projects_title: "All Projects",
        btn_new_project: "New Project",
        // Tabs
        tab_overview: "Overview",
        tab_tasks: "Tasks",
        tab_board: "Board",
        tab_gantt: "Gantt",
        tab_milestones: "Milestones",
        tab_timesheet: "Timesheet",
        tab_billing: "Billing",
        // Overview Tab
        task_breakdown: "Task Breakdown",
        status_todo: "To Do",
        status_in_progress: "In Progress",
        status_done: "Completed",
        // Tasks Tab
        btn_export: "Export",
        btn_add_task: "Add Task",
        // Gantt Tab
        gantt_schedule: "Schedule",
        zoom_day: "Day",
        zoom_week: "Week",
        zoom_month: "Month",
        // Milestones Tab
        btn_add_milestone: "Add Milestone",
        // Timesheet Tab
        btn_log_time: "Log Time",
        // Billing Tab
        invoices_title: "Invoices",
        btn_create_invoice: "Create Invoice",
        summary_title: "Summary",
        bill_unbilled: "Unbilled",
        bill_billed: "Billed",
        bill_rate: "Rate",
        // Settings Screen
        profile_settings_title: "Profile Settings",
        avatar_color_label: "Avatar Color",
        full_name_label: "Full Name",
        email_readonly_label: "Email (read-only)",
        btn_save_changes: "Save Changes",
        // Modals
        modal_new_ws: "New Workspace",
        ws_name_label: "Workspace Name",
        btn_cancel: "Cancel",
        btn_create: "Create",
        modal_new_proj: "New Project",
        proj_name_label: "Project Name",
        proj_desc_label: "Description",
        modal_add_task: "Add Task",
        modal_edit_task: "Edit Task",
        task_name_label: "Task Name",
        milestone_opt_label: "Milestone (optional)",
        recurrence_label: "Recurrence",
        recur_none: "One-time",
        recur_daily: "Daily",
        recur_weekly: "Weekly",
        recur_monthly: "Monthly",
        pri_low: "Low",
        pri_med: "Medium",
        pri_high: "High",
        modal_add_ms: "Add Milestone",
        modal_edit_ms: "Edit Milestone",
        ms_title_label: "Milestone Title",
        target_date_label: "Target Date",
        btn_add: "Add",
        modal_log_time: "Log Time",
        modal_edit_time: "Edit Time Entry",
        notes_label: "Notes",
        btn_log_entry: "Log Entry",
        modal_create_inv: "Create Invoice",
        modal_edit_inv: "Edit Invoice",
        client_name_label: "Client Name",
        btn_generate: "Generate",
        invoice_title: "Invoice",
        btn_print: "Print / PDF",
        attachments_title: "Attachments",
        btn_upload: "Upload",
        comments_title: "Comments",
        btn_post: "Post",
        client_addr_label: "Client Address & Tax ID",
        inv_issue_date: "Issue Date",
        inv_due_date: "Due Date",
        unbilled_timesheets_checklist: "⏱️ Include Unbilled Timesheets",
        inv_line_items: "📦 Line Items Breakdown",
        btn_add_line_item: "➕ Add Item",
        th_rate: "Rate ($)",
        th_amount: "Amount",
        inv_subtotal: "Subtotal:",
        inv_discount: "Discount ($):",
        inv_tax_rate: "Tax / VAT (%):",
        inv_grand_total: "Grand Total:",
        btn_generate_inv: "Generate Invoice",
        btn_quick_bill: "⚡ Bill Unbilled Hours",
        bill_unpaid_rev: "Unpaid Revenue",
        bill_paid_rev: "Paid Revenue",
        // Dynamic Text & Empty States
        no_projects: "No projects in this workspace. Click <strong>New Project</strong> to create one.",
        no_pending_tasks: "No pending tasks",
        no_tasks: "No tasks yet. Click <strong>Add Task</strong> to create one.",
        no_milestones: "No milestones yet.",
        no_timesheets: "No time entries yet.",
        no_invoices: "No invoices yet.",
        no_comments: "No comments yet",
        no_files: "No files attached",
        drag_here: "Drag tasks here",
        no_gantt_tasks: "No tasks with due dates to display.",
        created_on: "Created",
        open_btn: "Open →",
        switch_project: "Switch Project…",
        menu_exec_dashboard: "📊 Executive Dashboard",
        menu_team: "👥 Team & Hierarchy",
        exec_title: "Executive Portfolio Dashboard",
        exec_subtitle: "High-level portfolio health, revenue tracking, and team workload distribution across all projects.",
        btn_refresh: "Refresh",
        kpi_projects: "Total Projects",
        kpi_tasks: "Total Tasks",
        kpi_completion: "Overall Completion",
        kpi_hours: "Total Logged Hours",
        kpi_revenue: "Total Revenue",
        exec_health_title: "Project Health Roster",
        exec_workload_title: "Team Workload Distribution Matrix",
        th_employee: "Employee / Member",
        th_role_dept: "Role & Department",
        th_reports_to: "Reports To (Manager)",
        th_active_tasks: "Active Tasks",
        th_completed_tasks: "Completed",
        th_efficiency: "Efficiency",
        th_logged_hours: "Logged Hours",
        team_title: "Organization & Team Hierarchy",
        team_subtitle: "Define employee roles (Admin, Manager, Member), reporting structures, and assign organizational responsibilities.",
        btn_add_member: "Add Team Member",
        th_title: "Job Title",
        th_dept: "Department",
        th_role: "Role Level",
        th_supervisor: "Supervisor / Manager",
        th_actions: "Actions",
        modal_add_user: "Add Team Member",
        modal_edit_user: "Edit Team Member",
        role_member: "Team Member",
        role_manager: "Supervisor / Manager",
        role_admin: "Admin / Executive",
        btn_save: "Save Member",
        th_assignee: "Assignee",
        unassigned: "Unassigned",
        reset_pass_title: "Reset User Password",
        new_password_label: "New Password",
        btn_quick_pass: "⚡ Quick Set: admin123",
        btn_save_pass: "Save New Password",
        btn_reset_pass: "🔑 Password",
        btn_set_rate: "Set Rate",
        modal_set_rate: "⚙️ Set Standard Hourly Rate",
        set_rate_desc: "Configure the standard billing rate per hour for this workspace. This rate will be used when automatically calculating unbilled timesheets and generating new invoices.",
        hourly_rate_label: "Standard Rate ($ / hr)"
    },
    th: {
        // Login
        login_subtitle: "เข้าสู่ระบบพื้นที่ทำงานองค์กรของคุณ",
        login_microsoft: "เข้าสู่ระบบด้วย Microsoft",
        login_or: "หรือเข้าสู่ระบบด้วยอีเมล",
        email_label: "ที่อยู่อีเมล",
        password_label: "รหัสผ่าน",
        sign_in_btn: "เข้าสู่ระบบ",
        demo_creds: "บัญชีทดลองใช้งาน:",
        ms_hint: "การใช้งาน Microsoft login ต้องตั้งค่า Azure AD App Registration ก่อน",
        // Sidebar & Nav
        workspace_label: "พื้นที่ทำงาน",
        menu_dashboard: "แดชบอร์ด",
        menu_projects: "โครงการทั้งหมด",
        menu_settings: "การตั้งค่า",
        menu_exec_dashboard_short: "บริหาร",
        menu_team_short: "จัดการทีม",
        user_role: "ผู้จัดการ",
        notif_title: "การแจ้งเตือน",
        notif_mark_read: "อ่านทั้งหมดแล้ว",
        // Dashboard
        stat_projects: "จำนวนโครงการ",
        stat_active_tasks: "งานที่กำลังดำเนินการ",
        stat_hours: "ชั่วโมงที่บันทึก",
        stat_milestones: "หมุดหมายสำคัญ",
        panel_pending_tasks: "งานที่รอจัดการ",
        panel_progress: "ความคืบหน้าโครงการ",
        th_task: "ชื่องาน",
        th_project: "โครงการ",
        th_start: "วันเริ่มต้น (ไม่บังคับ)",
        th_due: "กำหนดส่ง",
        th_priority: "ความสำคัญ",
        th_status: "สถานะ",
        th_date: "วันที่",
        th_desc: "รายละเอียด",
        th_user: "ผู้บันทึก",
        th_hours: "ชั่วโมง",
        // Projects Screen
        all_projects_title: "โครงการทั้งหมดในพื้นที่ทำงาน",
        btn_new_project: "สร้างโครงการใหม่",
        // Tabs
        tab_overview: "ภาพรวม",
        tab_tasks: "รายการงาน",
        tab_board: "บอร์ด Kanban",
        tab_gantt: "ปฏิทิน Gantt",
        tab_milestones: "หมุดหมาย (Milestones)",
        tab_timesheet: "บันทึกเวลา",
        tab_billing: "ใบแจ้งหนี้",
        // Overview Tab
        task_breakdown: "สัดส่วนงาน (Task Breakdown)",
        overdue_tasks: "งานที่เลยกำหนด (Overdue)",
        recent_activity: "ความเคลื่อนไหวล่าสุด",
        // Tasks Tab
        btn_export: "ส่งออก Excel/CSV",
        btn_add_task: "เพิ่มงานใหม่",
        all_status: "ทุกสถานะ",
        status_todo: "รอดำเนินการ",
        status_in_progress: "กำลังดำเนินการ",
        status_done: "เสร็จสิ้นแล้ว",
        filter_priority: "ทุกความสำคัญ",
        pri_low: "ต่ำ (Low)",
        pri_med: "ปานกลาง (Medium)",
        pri_high: "สูง (High)",
        recur_none: "ครั้งเดียว (One-time)",
        recur_daily: "ทุกวัน (Daily)",
        recur_weekly: "ทุกสัปดาห์ (Weekly)",
        recur_monthly: "ทุกเดือน (Monthly)",
        recurrence_label: "รอบการทำซ้ำ",
        modal_edit_task: "แก้ไขงาน",
        btn_save_changes: "บันทึกการเปลี่ยนแปลง",
        // Gantt Tab
        gantt_schedule: "ตารางเวลาโครงการ (Gantt Schedule)",
        zoom_day: "วัน (Day)",
        zoom_week: "สัปดาห์ (Week)",
        zoom_month: "เดือน (Month)",
        // Milestones
        btn_add_milestone: "เพิ่มหมุดหมาย",
        modal_add_ms: "เพิ่มหมุดหมายสำคัญ (Milestone)",
        modal_edit_ms: "แก้ไขหมุดหมาย",
        ms_title_label: "ชื่อหมุดหมาย",
        target_date_label: "วันที่เป้าหมาย",
        btn_generate: "ออกใบแจ้งหนี้",
        modal_create_inv: "สร้างใบแจ้งหนี้",
        modal_edit_inv: "แก้ไขใบแจ้งหนี้",
        invoice_title: "ใบแจ้งหนี้ / ใบเสร็จ",
        btn_print: "พิมพ์ / บันทึก PDF",
        attachments_title: "ไฟล์แนบ",
        btn_upload: "อัปโหลดไฟล์",
        comments_title: "ความคิดเห็น / อภิปราย",
        btn_post: "แสดงความคิดเห็น",
        client_addr_label: "ที่อยู่ลูกค้า & เลขผู้เสียภาษี",
        inv_issue_date: "วันที่ออกบิล",
        inv_due_date: "กำหนดชำระเงิน",
        unbilled_timesheets_checklist: "⏱️ เลือกรายการบันทึกเวลาที่ยังไม่เรียกเก็บ",
        inv_line_items: "📦 รายการค่าบริการ / ค่าใช้จ่าย",
        btn_add_line_item: "➕ เพิ่มรายการ",
        th_rate: "อัตรา (฿/$)",
        th_amount: "จำนวนเงิน",
        inv_subtotal: "ยอดรวมก่อนภาษี:",
        inv_discount: "ส่วนลด:",
        inv_tax_rate: "ภาษีมูลค่าเพิ่ม VAT (%):",
        inv_grand_total: "ยอดสุทธิต้องชำระ:",
        btn_generate_inv: "สร้างใบแจ้งหนี้ทันที",
        btn_quick_bill: "⚡ ออกบิลสำหรับเวลาที่ยังไม่เรียกเก็บ",
        bill_unpaid_rev: "ยอดรอรับชำระ",
        bill_paid_rev: "ยอดชำระแล้ว",
        // Dynamic Text & Empty States
        no_projects: "ยังไม่มีโครงการในพื้นที่ทำงานนี้ คลิก <strong>สร้างโครงการใหม่</strong> เพื่อเริ่มต้น",
        no_pending_tasks: "ไม่มีงานค้างที่ต้องจัดการ",
        no_tasks: "ยังไม่มีงานในโครงการนี้ คลิก <strong>เพิ่มงานใหม่</strong> เพื่อเริ่มสั่งงาน",
        no_milestones: "ยังไม่มีหมุดหมายในโครงการนี้",
        no_timesheets: "ยังไม่มีการบันทึกเวลาทำงาน",
        no_invoices: "ยังไม่มีใบแจ้งหนี้ในโครงการนี้",
        no_comments: "ยังไม่มีความคิดเห็น เริ่มสนทนาได้เลย",
        no_files: "ยังไม่มีไฟล์แนบในงานนี้",
        drag_here: "ลากการ์ดงานมาวางที่นี่",
        no_gantt_tasks: "ไม่มีงานที่มีกำหนดส่งสำหรับแสดงบนปฏิทิน Gantt",
        created_on: "สร้างเมื่อ",
        open_btn: "เข้าชมโครงการ →",
        switch_project: "สลับโครงการ…",
        menu_exec_dashboard: "📊 แดชบอร์ดผู้บริหาร",
        menu_team: "👥 จัดการทีมและโครงสร้าง",
        exec_title: "แดชบอร์ดภาพรวมผู้บริหาร (Executive Portfolio)",
        exec_subtitle: "ติดตามสถานะสุขภาพโครงการ รายได้ และการกระจายภาระงานของทีมในทุกโครงการแบบเรียลไทม์",
        btn_refresh: "รีเฟรชข้อมูล",
        kpi_projects: "จำนวนโครงการรวม",
        kpi_tasks: "จำนวนงานทั้งหมด",
        kpi_completion: "ความคืบหน้ารวม",
        kpi_hours: "ชั่วโมงทำงานรวม",
        kpi_revenue: "รายได้โดยรวม",
        exec_health_title: "สรุปสถานะโครงการ (Project Health)",
        exec_workload_title: "ตารางกระจายภาระงานของทีม (Workload Matrix)",
        th_employee: "พนักงาน / สมาชิกทีม",
        th_role_dept: "ระดับบทบาท & แผนก",
        th_reports_to: "รายงานต่อ (หัวหน้างาน)",
        th_active_tasks: "งานที่กำลังทำ",
        th_completed_tasks: "เสร็จสิ้นแล้ว",
        th_efficiency: "ประสิทธิภาพ",
        th_logged_hours: "ชั่วโมงทำงาน",
        team_title: "โครงสร้างองค์กรและจัดการทีม (Team & Hierarchy)",
        team_subtitle: "กำหนดระดับสิทธิ์ (แอดมิน, หัวหน้างาน, ทีมงาน) และสายบังคับบัญชา เพื่อใช้ในการสั่งและมอบหมายงาน",
        btn_add_member: "เพิ่มสมาชิกทีม",
        th_title: "ตำแหน่งงาน",
        th_dept: "ฝ่าย / แผนก",
        th_role: "ระดับสิทธิ์",
        th_supervisor: "หัวหน้าผู้ควบคุม (Supervisor)",
        th_actions: "จัดการ",
        modal_add_user: "เพิ่มสมาชิกทีม",
        modal_edit_user: "แก้ไขข้อมูลสมาชิก",
        role_member: "ทีมงานทั่วไป (Team Member)",
        role_manager: "หัวหน้างาน (Supervisor / Manager)",
        role_admin: "แอดมิน / ผู้บริหาร (Admin)",
        btn_save: "บันทึกข้อมูล",
        th_assignee: "ผู้รับผิดชอบงาน",
        unassigned: "ไม่ระบุ (Unassigned)",
        reset_pass_title: "รีเซ็ตและแก้ไขรหัสผ่านผู้ใช้งาน",
        new_password_label: "รหัสผ่านใหม่",
        btn_quick_pass: "⚡ ตั้งค่าด่วนเป็น: admin123",
        btn_save_pass: "บันทึกรหัสผ่านใหม่",
        btn_reset_pass: "🔑 รหัสผ่าน",
        btn_set_rate: "ตั้งค่าเรทราคา",
        modal_set_rate: "⚙️ ตั้งค่าเรทค่าบริการมาตรฐาน (Standard Rate)",
        set_rate_desc: "กำหนดอัตราค่าบริการมาตรฐานต่อชั่วโมงของพื้นที่ทำงานนี้ ระบบจะใช้ตัวเลขนี้ในการคำนวณยอดเงินของชั่วโมงทำงานที่ยังไม่เบิก และสร้างใบแจ้งหนี้อัตโนมัติ",
        hourly_rate_label: "เรทค่าบริการมาตรฐาน ($ / ชั่วโมง)"
    }
};

function t(key) {
    const dict = translations[currentLang] || translations.en;
    return dict[key] || translations.en[key] || key;
}

// ---- Theme Switcher ----
function toggleTheme() {
    currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', currentTheme);
    localStorage.setItem('ag_theme', currentTheme);
    updateThemeUI();
    showToast(currentTheme === 'dark' ? '🌙 Dark Theme enabled' : '☀️ Light Theme enabled');
}

function updateThemeUI() {
    const icon = currentTheme === 'dark' ? '🌙' : '☀️';
    const text = currentTheme === 'dark' ? 'Dark' : 'Light';
    document.querySelectorAll('.theme-toggle-btn').forEach(btn => {
        const iconEl = btn.querySelector('.theme-icon');
        const textEl = btn.querySelector('.theme-text');
        if (iconEl) iconEl.textContent = icon;
        if (textEl) textEl.textContent = text;
    });
}

// ---- Language Switcher ----
function toggleLanguage() {
    currentLang = currentLang === 'en' ? 'th' : 'en';
    localStorage.setItem('ag_lang', currentLang);
    applyLanguage();
    showToast(currentLang === 'th' ? '🌐 เปลี่ยนเป็นภาษาไทยเรียบร้อย' : '🌐 Switched to English');
}

function applyLanguage() {
    const dict = translations[currentLang] || translations.en;
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (dict[key]) el.innerHTML = dict[key];
    });
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        if (dict[key]) el.setAttribute('placeholder', dict[key]);
    });
    document.querySelectorAll('.lang-toggle-btn').forEach(btn => {
        const iconEl = btn.querySelector('.lang-icon');
        const textEl = btn.querySelector('.lang-text');
        if (iconEl) iconEl.textContent = '🌐';
        if (textEl) textEl.textContent = currentLang === 'en' ? 'EN' : 'TH';
    });
    // Refresh dynamic lists
    refreshCurrentView();
}

function refreshCurrentView() {
    if (!currentUser) return;
    const activeScreen = document.querySelector('.app-screen.active');
    if (!activeScreen) return;
    const id = activeScreen.id;
    if (id === 'screen-dashboard') loadDashboard();
    else if (id === 'screen-projects') loadProjects();
    else if (id === 'screen-executive-dashboard') loadExecutiveDashboard();
    else if (id === 'screen-team-management') loadTeamManagement();
    else if (id === 'screen-project-workspace') {
        const activeTab = document.querySelector('.tab-btn.active');
        if (activeTab) {
            const tab = activeTab.dataset.tab;
            if (tab === 'dashboard') loadProjectOverview();
            else if (tab === 'tasks') loadProjectTasks();
            else if (tab === 'board') loadKanban();
            else if (tab === 'gantt') renderGantt();
            else if (tab === 'milestones') loadMilestones();
            else if (tab === 'timesheet') loadTimesheets();
            else if (tab === 'billing') loadBilling();
        }
    }
}

// Initialize Theme and Language on startup
updateThemeUI();
applyLanguage();
if (document.querySelector('#app-container')) {
    document.querySelector('#app-container').style.display = 'none';
}

// ---- Helpers ----
const API = (path) => path;
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// Escapes user-generated text before it is interpolated into innerHTML, to prevent stored XSS
function esc(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// Lightweight loading indicator shown while a fetch is in flight
function loadingState() {
    return `<div class="loading-state"><div class="loading-spinner"></div></div>`;
}
function loadingStateRow(colspan) {
    return `<tr><td colspan="${colspan}">${loadingState()}</td></tr>`;
}

// Reusable empty-state block: icon + title + optional description + optional CTA button
function emptyState(icon, title, desc, ctaLabel, ctaOnclick) {
    return `<div class="empty-state">
        <div class="empty-state-icon">${icon}</div>
        <div class="empty-state-title">${esc(title)}</div>
        ${desc ? `<div class="empty-state-desc">${esc(desc)}</div>` : ''}
        ${ctaLabel && ctaOnclick ? `<button type="button" class="btn btn-primary empty-state-cta" onclick="${ctaOnclick}">${esc(ctaLabel)}</button>` : ''}
    </div>`;
}

// Same content, wrapped for use as a single spanning table row
function emptyStateRow(colspan, icon, title, desc, ctaLabel, ctaOnclick) {
    return `<tr><td colspan="${colspan}">${emptyState(icon, title, desc, ctaLabel, ctaOnclick)}</td></tr>`;
}

// Escapes text used as a single-quoted JS string literal inside an inline onclick="" HTML attribute
function escJsAttr(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function toggleMobileWorkspaceSelect(show) {
    const el = $('#nav-workspace-selector-wrapper');
    if (el) {
        if (window.innerWidth <= 900) {
            el.style.display = show ? 'inline-block' : 'none';
        } else {
            el.style.display = 'none';
        }
    }
}

function showToast(msg) {
    const tEl = $('#toast');
    $('#toast-message').textContent = msg;
    tEl.classList.add('active');
    setTimeout(() => tEl.classList.remove('active'), 3200);
}

function toggleMobileSidebar(force) {
    const sb = $('#app-sidebar');
    const ov = $('#sidebar-overlay');
    if (!sb || !ov) return;
    const isOpen = sb.classList.contains('mobile-open');
    const shouldOpen = typeof force === 'boolean' ? force : !isOpen;
    if (shouldOpen) {
        sb.classList.add('mobile-open');
        ov.classList.add('active');
    } else {
        sb.classList.remove('mobile-open');
        ov.classList.remove('active');
    }
}

function formatDate(d) {
    if (!d) return '—';
    const loc = currentLang === 'th' ? 'th-TH' : 'en-US';
    return new Date(d).toLocaleDateString(loc, { month: 'short', day: 'numeric', year: 'numeric' });
}

function priorityBadge(p) {
    const cls = { high: 'badge-high', medium: 'badge-medium', low: 'badge-low' };
    const label = p === 'high' ? t('pri_high') : p === 'medium' ? t('pri_med') : t('pri_low');
    return `<span class="badge ${cls[p] || 'badge-low'}">${label}</span>`;
}

function statusBadge(s) {
    const map = {
        todo: ['badge-todo', t('status_todo')],
        in_progress: ['badge-progress', t('status_in_progress')],
        completed: ['badge-done', t('status_done')]
    };
    const [cls, label] = map[s] || map.todo;
    return `<span class="badge ${cls}">${label}</span>`;
}

function getInitials(name) {
    if (!name) return '??';
    return name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
}

// ---- Microsoft Login ----
async function loginWithMicrosoft() {
    if (!msalInstance) {
        showToast(currentLang === 'th' ? 'กรุณาตั้งค่า Client ID ของ Azure AD ใน app.js เพื่อใช้ Microsoft Login' : 'Microsoft Login requires Azure AD configuration in app.js');
        return;
    }
    try {
        const resp = await msalInstance.loginPopup({ scopes: ['User.Read'] });
        const account = resp.account;
        const res = await fetch(API('/api/login/microsoft'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: account.name, email: account.username })
        });
        const data = await res.json();
        if (data.success) {
            currentUser = data.user;
            localStorage.setItem('ag_session_token', data.token);
            onLoginSuccess();
        } else {
            showToast('Microsoft login failed');
        }
    } catch (e) {
        showToast('Microsoft login cancelled or failed.');
        console.error('MSAL Error:', e);
    }
}

// ---- Email/Password Login ----
$('#login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = $('#login-email').value.trim();
    const password = $('#login-password').value.trim();
    
    // Frontend validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        showToast(currentLang === 'th' ? 'กรุณากรอกรูปแบบอีเมลให้ถูกต้อง' : 'Please enter a valid email address');
        return;
    }
    
    try {
        const res = await fetch(API('/api/login'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, password })
        });
        const data = await res.json();
        if (data.success) {
            currentUser = data.user;
            localStorage.setItem('ag_session_token', data.token);
            onLoginSuccess();
        } else {
            showToast(currentLang === 'th' ? 'อีเมลหรือรหัสผ่านไม่ถูกต้อง ลองใช้ admin@demo.com / admin123' : 'Invalid credentials. Try admin@demo.com / admin123');
        }
    } catch (err) {
        showToast('Cannot connect to server.');
    }
});

function onLoginSuccess(isRestore = false) {
    localStorage.setItem('ag_current_user', JSON.stringify(currentUser));
    $('#login-screen').classList.add('hidden');
    const appContainer = $('#app-container');
    if (appContainer) appContainer.style.display = 'flex';
    updateUserUI();
    loadWorkspaces();
    updateNotificationBadge();

    if (isRestore) {
        const targetScreen = localStorage.getItem('ag_last_screen') || 'dashboard';
        const targetMenus = Array.from([...$$('.menu-item'), ...$$('.mobile-nav-item')]).filter(m => m.dataset.screen === targetScreen);
        $$('.menu-item').forEach(m => m.classList.remove('active'));
        $$('.mobile-nav-item').forEach(m => m.classList.remove('active'));
        targetMenus.forEach(m => m.classList.add('active'));
        
        switchScreen(`screen-${targetScreen}`);
        $('#nav-breadcrumb').style.display = 'none';
        const titles = { 
            dashboard: t('menu_dashboard'), 
            'executive-dashboard': t('menu_exec_dashboard'), 
            projects: t('menu_projects'), 
            'team-management': t('menu_team'), 
            settings: t('menu_settings') 
        };
        if ($('#current-screen-title')) $('#current-screen-title').textContent = titles[targetScreen] || targetScreen;
        toggleMobileWorkspaceSelect(targetScreen === 'projects');
        if (targetScreen === 'projects') {
            loadProjects();
        } else if (targetScreen === 'dashboard') {
            loadDashboard();
        } else if (targetScreen === 'executive-dashboard') {
            loadExecutiveDashboard();
        } else if (targetScreen === 'team-management') {
            loadTeamManagement();
        } else if (targetScreen === 'settings') {
            loadSettings();
        }
    } else {
        const navItems = [...$$('.menu-item'), ...$$('.mobile-nav-item')];
        navItems.forEach(m => m.classList.remove('active'));
        navItems.forEach(m => {
            if (m.dataset.screen === 'dashboard') m.classList.add('active');
        });
        switchScreen('screen-dashboard');
        if ($('#nav-breadcrumb')) $('#nav-breadcrumb').style.display = 'none';
        if ($('#current-screen-title')) $('#current-screen-title').textContent = t('menu_dashboard');
        toggleMobileWorkspaceSelect(false);
        showToast(currentLang === 'th' ? `ยินดีต้อนรับกลับมา, ${currentUser.name}!` : `Welcome back, ${currentUser.name}!`);
    }
}

function updateUserUI() {
    const initials = getInitials(currentUser.name);
    const color = currentUser.color || currentUser.avatar_color || '#2684FF';
    if (currentUser.avatar_url && currentUser.avatar_url.trim() !== '') {
        $('#sidebar-avatar').innerHTML = `<img src="${esc(currentUser.avatar_url)}" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">`;
        $('#sidebar-avatar').style.backgroundColor = 'transparent';
    } else {
        $('#sidebar-avatar').innerHTML = '';
        $('#sidebar-avatar').textContent = initials;
        $('#sidebar-avatar').style.backgroundColor = color;
    }
    $('#sidebar-username').textContent = currentUser.name;
    const roleEl = document.querySelector('.sidebar-user .user-role');
    if (roleEl && currentUser.title) {
        roleEl.textContent = `${currentUser.title} (${currentUser.role ? currentUser.role.toUpperCase() : 'MEMBER'})`;
        roleEl.removeAttribute('data-i18n');
    } else if (roleEl && currentUser.role) {
        const roleNames = { admin: t('role_admin'), manager: t('role_manager'), member: t('role_member') };
        roleEl.textContent = roleNames[currentUser.role] || currentUser.role;
        roleEl.removeAttribute('data-i18n');
    }
}

// ---- Logout ----
$('#logout-btn').addEventListener('click', async () => {
    try {
        await fetch(API('/api/logout'), { method: 'POST' });
    } catch (e) {}
    currentUser = null;
    currentWorkspaceId = null;
    currentProjectId = null;
    localStorage.removeItem('ag_current_user');
    localStorage.removeItem('ag_session_token');
    localStorage.removeItem('ag_last_screen');
    if ($('#app-container')) $('#app-container').style.display = 'none';
    $('#login-screen').classList.remove('hidden');
    showToast(currentLang === 'th' ? 'ออกจากระบบเรียบร้อย' : 'Signed out');
});

// ---- Navigation ----
const navItems = [...$$('.menu-item'), ...$$('.mobile-nav-item')];
navItems.forEach(item => {
    item.addEventListener('click', () => {
        const screen = item.dataset.screen;
        if (!screen) return;
        if (window.innerWidth <= 900) toggleMobileSidebar(false);
        navItems.forEach(m => {
            if (m.dataset.screen === screen) {
                m.classList.add('active');
            } else {
                m.classList.remove('active');
            }
        });
        switchScreen(`screen-${screen}`);
        
        $('#nav-breadcrumb').style.display = 'none';
        const titles = { 
            dashboard: t('menu_dashboard'), 
            'executive-dashboard': t('menu_exec_dashboard'), 
            projects: t('menu_projects'), 
            'team-management': t('menu_team'), 
            settings: t('menu_settings') 
        };
        $('#current-screen-title').textContent = titles[screen] || screen;
        
        toggleMobileWorkspaceSelect(screen === 'projects');

        if (screen === 'projects') {
            loadProjects();
        } else if (screen === 'dashboard') {
            loadDashboard();
        } else if (screen === 'executive-dashboard') {
            loadExecutiveDashboard();
        } else if (screen === 'team-management') {
            loadTeamManagement();
        } else if (screen === 'settings') {
            loadSettings();
        }
    });
});

function navigateToProjects() {
    const navItems = [...$$('.menu-item'), ...$$('.mobile-nav-item')];
    navItems.forEach(m => m.classList.remove('active'));
    navItems.forEach(m => {
        if (m.dataset.screen === 'projects') m.classList.add('active');
    });
    switchScreen('screen-projects');
    $('#nav-breadcrumb').style.display = 'none';
    $('#current-screen-title').textContent = t('menu_projects');
    toggleMobileWorkspaceSelect(true);
    loadProjects();
}

function switchScreen(screenId) {
    if (screenId) localStorage.setItem('ag_last_screen', screenId.replace('screen-', ''));
    $$('.app-screen').forEach(s => s.classList.remove('active'));
    const el = $(`#${screenId}`);
    if (el) el.classList.add('active');
}

// ---- Workspace Tabs ----
$$('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $$('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        $$('.workspace-pane').forEach(p => p.classList.remove('active'));
        const pane = $(`#pane-${btn.dataset.tab}`);
        if (pane) pane.classList.add('active');
        const tab = btn.dataset.tab;
        if (tab === 'tasks') loadProjectTasks();
        else if (tab === 'board') loadKanban();
        else if (tab === 'gantt') renderGantt();
        else if (tab === 'milestones') loadMilestones();
        else if (tab === 'timesheet') loadTimesheets();
        else if (tab === 'billing') loadBilling();
        else if (tab === 'dashboard') loadProjectOverview();
    });
});

// ---- Gantt View Toggles ----
$$('.gantt-view-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $$('.gantt-view-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        ganttZoom = btn.dataset.zoom;
        renderGantt();
    });
});

// ---- Workspace Selector ----
$('#workspace-select').addEventListener('change', async (e) => {
    selectWorkspace(e.target.value);
});

if ($('#mobile-navbar-workspace-select')) {
    $('#mobile-navbar-workspace-select').addEventListener('change', (e) => {
        const val = e.target.value;
        if (val === 'CREATE_NEW_WS') {
            e.target.value = currentWorkspaceId; // revert select
            openModal('new-workspace-modal');
        } else {
            selectWorkspace(val);
        }
    });
}

function selectWorkspace(id) {
    currentWorkspaceId = id;
    currentProjectId = null;
    const sel = $('#workspace-select');
    if (sel) sel.value = id;
    const mobSel = $('#mobile-navbar-workspace-select');
    if (mobSel) mobSel.value = id;
    if (window.innerWidth <= 900) toggleMobileSidebar(false);
    renderWorkspacesList();
    navigateToProjects();
}
window.selectWorkspace = selectWorkspace;

function renderWorkspacesList() {
    const listEl = $('#sidebar-workspaces-list');
    if (!listEl) return;
    listEl.innerHTML = allWorkspacesList.map(w => `
        <a class="workspace-list-item ${w.id === currentWorkspaceId ? 'active' : ''}" onclick="selectWorkspace('${w.id}')">
            <span>${esc(w.name)}</span>
        </a>
    `).join('');
}

// ---- Data Loading ----
async function loadWorkspaces() {
    try {
        const res = await fetch(API('/api/workspaces'));
        const workspaces = await res.json();
        allWorkspacesList = workspaces;
        const sel = $('#workspace-select');
        sel.innerHTML = workspaces.map(w => `<option value="${w.id}">${esc(w.name)}</option>`).join('');
        if (currentWorkspaceId) {
            sel.value = currentWorkspaceId;
        } else if (workspaces.length > 0) {
            currentWorkspaceId = workspaces[0].id;
            sel.value = currentWorkspaceId;
        }
        const mobSel = $('#mobile-navbar-workspace-select');
        if (mobSel) {
            mobSel.innerHTML = workspaces.map(w => `<option value="${w.id}">${esc(w.name)}</option>`).join('') +
                `<option value="CREATE_NEW_WS" style="color:var(--accent);font-weight:600;">➕ ${currentLang === 'th' ? 'สร้างพื้นที่ทำงานใหม่...' : 'Create Workspace...'}</option>`;
            if (currentWorkspaceId) mobSel.value = currentWorkspaceId;
        }
        renderWorkspacesList();
        loadTeamManagement();
        loadDashboard();
    } catch (e) { showToast('Failed to load workspaces'); }
}

let allDashboardTasks = [];
let allDashboardProjects = [];
let currentProjectTasksList = [];
let dashboardTaskFilterMode = 'all';

async function loadDashboard() {
    if ($('#dashboard-tasks-tbody')) $('#dashboard-tasks-tbody').innerHTML = loadingStateRow(6);
    if ($('#dashboard-projects-progress')) $('#dashboard-projects-progress').innerHTML = loadingState();
    try {
        const [projectsRes, tasksArr, tsRes, msRes] = await Promise.all([
            fetch(API(`/api/projects?workspace_id=${currentWorkspaceId}`)).then(r => r.json()),
            fetch(API(`/api/tasks?workspace_id=${currentWorkspaceId}`)).then(r => r.json()),
            loadAllTimesheets(),
            loadAllMilestones()
        ]);
        allDashboardProjects = projectsRes || [];
        allDashboardTasks = tasksArr || [];

        // Stats
        $('#stat-projects').textContent = allDashboardProjects.length;
        const activeTasks = allDashboardTasks.filter(t => t.status !== 'completed');
        $('#stat-tasks').textContent = activeTasks.length;
        const totalHours = tsRes.reduce((s, t) => s + (t.hours || 0), 0);
        $('#stat-hours').textContent = totalHours.toFixed(1);
        $('#stat-milestones').textContent = msRes.length;

        renderDashboardTasksTable();

        // Projects progress
        const progressEl = $('#dashboard-projects-progress');
        progressEl.innerHTML = allDashboardProjects.map(p => {
            const pTasks = allDashboardTasks.filter(t => t.project_id === p.id);
            const done = pTasks.filter(t => t.status === 'completed').length;
            const pct = pTasks.length > 0 ? Math.round((done / pTasks.length) * 100) : 0;
            return `<div class="progress-item">
                <div class="progress-item-header"><span class="progress-item-title">${esc(p.name)}</span><span>${pct}%</span></div>
                <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct}%"></div></div>
            </div>`;
        }).join('') || emptyState('📁', currentLang === 'th' ? 'ยังไม่มีโครงการในพื้นที่ทำงานนี้' : 'No projects in this workspace');

    } catch (e) { console.error('Dashboard load error', e); }
}

function renderDashboardTasksTable() {
    const tbody = $('#dashboard-tasks-tbody');
    if (!tbody) return;
    let filtered = [...allDashboardTasks];
    if (dashboardTaskFilterMode === 'active') {
        filtered = filtered.filter(t => t.status !== 'completed');
    } else if (dashboardTaskFilterMode === 'completed') {
        filtered = filtered.filter(t => t.status === 'completed');
    }
    const countBadge = $('#dashboard-tasks-count');
    if (countBadge) countBadge.textContent = filtered.length;

    tbody.innerHTML = filtered.map(t => {
        const proj = allDashboardProjects.find(p => p.id === t.project_id);
        const assigneeName = t.assignee_name && t.assignee_name !== 'Unassigned' ? t.assignee_name : (currentLang === 'th' ? 'ยังไม่ระบุ' : 'Unassigned');
        const assigneeAvatar = t.assignee_avatar || '#8b949e';
        const initials = getInitials(assigneeName);
        return `<tr>
            <td>
                <span class="task-title-cell" style="font-weight:600; cursor:pointer; color:var(--text-primary);" onclick="openTaskDetail('${t.id}')">${esc(t.name)}</span>
                <span class="project-subtitle-cell" style="font-size:0.75rem; color:var(--text-muted); display:block;">${proj ? esc(proj.name) : ''}</span>
            </td>
            <td><span style="font-weight:500;">${proj ? esc(proj.name) : '—'}</span></td>
            <td>
                <div style="display:flex;align-items:center;gap:8px;">
                    <span style="width:24px;height:24px;border-radius:50%;background:${assigneeAvatar};color:#fff;font-weight:700;font-size:0.7rem;display:flex;align-items:center;justify-content:center;">${esc(initials)}</span>
                    <span style="font-size:0.85rem;">${esc(assigneeName)}</span>
                </div>
            </td>
            <td>${formatDate(t.due)}</td>
            <td>${priorityBadge(t.priority)}</td>
            <td>${statusBadge(t.status)}</td>
        </tr>`;
    }).join('') || emptyStateRow(6, '📋', currentLang === 'th' ? 'ไม่พบรายการงาน' : 'No tasks found');
}

function filterDashboardTasks(mode, btnEl) {
    dashboardTaskFilterMode = mode;
    $$('.dashboard-task-filter').forEach(b => b.classList.remove('active'));
    if (btnEl) btnEl.classList.add('active');
    renderDashboardTasksTable();
}

async function loadAllTasks() {
    try {
        const tasks = await fetch(API('/api/tasks')).then(r => r.json());
        return tasks || [];
    } catch (e) { return []; }
}

async function loadAllTimesheets() {
    try {
        const timesheets = await fetch(API(`/api/timesheets?workspace_id=${currentWorkspaceId}`)).then(r => r.json());
        return timesheets || [];
    } catch (e) { return []; }
}

async function loadAllMilestones() {
    try {
        const milestones = await fetch(API(`/api/milestones?workspace_id=${currentWorkspaceId}`)).then(r => r.json());
        return milestones || [];
    } catch (e) { return []; }
}

// ---- Executive Portfolio Dashboard ----
async function loadExecutiveDashboard() {
    if (!currentWorkspaceId) return;
    if ($('#exec-project-health-grid')) $('#exec-project-health-grid').innerHTML = loadingState();
    if ($('#exec-workload-tbody')) $('#exec-workload-tbody').innerHTML = loadingStateRow(7);
    try {
        const res = await fetch(API(`/api/workspaces/${currentWorkspaceId}/executive-summary`));
        if (!res.ok) throw new Error('Failed to fetch summary');
        const data = await res.json();
        
        // Populate KPIs
        const kpis = data.kpis || data.kpi || {};
        const completion = kpis.overall_completion_percentage ?? kpis.completion_percent ?? 0;
        const hours = kpis.total_logged_hours ?? kpis.total_hours ?? 0;
        const revenue = kpis.total_revenue ?? kpis.revenue ?? 0;
        $('#kpi-projects').textContent = kpis.total_projects || 0;
        $('#kpi-tasks').textContent = kpis.total_tasks || 0;
        $('#kpi-completion').textContent = `${completion}%`;
        $('#kpi-hours').textContent = `${parseFloat(hours).toFixed(1)} hrs`;
        $('#kpi-revenue').textContent = `$${parseFloat(revenue).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
        
        // Populate Project Health Roster
        const healthGrid = $('#exec-project-health-grid');
        const projects = data.project_health || [];
        if (projects.length === 0) {
            healthGrid.innerHTML = `<div style="grid-column:1/-1;">${emptyState('📁', currentLang === 'th' ? 'ยังไม่มีโครงการ' : 'No projects yet')}</div>`;
        } else {
            healthGrid.innerHTML = projects.map(p => {
                const statusBadge = p.status_badge || (p.overdue_tasks > 0 ? 'warning' : p.progress_percent === 100 ? 'success' : 'progress');
                const badgeCls = statusBadge === 'success' ? 'badge-done' : statusBadge === 'warning' ? 'badge-medium' : 'badge-progress';
                const healthStatus = p.health_status || (p.overdue_tasks > 0 ? 'Needs Attention' : p.progress_percent === 100 ? 'Completed' : 'Optimal');
                const rev = p.revenue ?? p.total_revenue ?? 0;
                const compPct = p.completion_percentage ?? p.progress_percent ?? 0;
                const loggedH = p.logged_hours ?? p.total_hours ?? 0;
                const teamSz = p.team_size || 3;
                return `<div class="exec-project-card" onclick="openProject('${p.id}', '${escJsAttr(p.name||'')}')">
                    <div class="exec-proj-header">
                        <div>
                            <div class="exec-proj-title">${esc(p.name) || 'Untitled'}</div>
                            <span class="badge ${badgeCls}" style="margin-top:2px;">${healthStatus}</span>
                        </div>
                        <div style="text-align:right;">
                            <div style="font-size:1.1rem; font-weight:800; color:var(--accent);">$${parseFloat(rev).toLocaleString()}</div>
                            <span style="font-size:0.72rem; color:var(--text-muted);">Revenue</span>
                        </div>
                    </div>
                    <div>
                        <div style="display:flex; justify-content:space-between; font-size:0.82rem; margin-bottom:6px;">
                            <span style="color:var(--text-secondary);">Task Completion</span>
                            <span style="font-weight:700;">${p.completed_tasks || 0} / ${p.total_tasks || 0} (${compPct}%)</span>
                        </div>
                        <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${compPct}%"></div></div>
                    </div>
                    <div class="exec-proj-meta">
                        <span>⏱️ ${parseFloat(loggedH).toFixed(1)} hrs logged</span>
                        <span>👥 ${teamSz} members</span>
                    </div>
                </div>`;
            }).join('');
        }

        // Populate Team Workload Matrix
        const workloadTbody = $('#exec-workload-tbody');
        const matrix = (data.workload_matrix || []).filter(m => m.role !== 'admin' && (m.title || '').toLowerCase() !== 'executive director');
        if (matrix.length === 0) {
            workloadTbody.innerHTML = emptyStateRow(7, '👥', currentLang === 'th' ? 'ยังไม่มีข้อมูลภาระงานของทีม' : 'No team workload data');
        } else {
            workloadTbody.innerHTML = matrix.map(m => {
                const initials = getInitials(m.name);
                const roleCls = m.role === 'admin' ? 'badge-admin' : m.role === 'manager' ? 'badge-manager' : 'badge-member';
                const roleName = m.role === 'admin' ? t('role_admin') : m.role === 'manager' ? t('role_manager') : t('role_member');
                return `<tr>
                    <td>
                        <div style="display:flex; align-items:center; gap:10px;">
                            ${m.avatar_url && m.avatar_url.trim() !== '' ?
                                `<img src="${esc(m.avatar_url)}" style="width:28px;height:28px;border-radius:50%;object-fit:cover;">` :
                                `<div class="avatar-xs" style="width:28px;height:28px;border-radius:50%;background:${m.avatar_color || '#2684FF'};color:#fff;font-weight:700;font-size:0.75rem;display:flex;align-items:center;justify-content:center;">${esc(initials)}</div>`
                            }
                            <div>
                                <div style="font-weight:600;color:var(--text-primary);">${esc(m.name)}</div>
                                <div style="font-size:0.75rem;color:var(--text-muted);">${esc(m.email)}</div>
                            </div>
                        </div>
                    </td>
                    <td>
                        <div style="font-weight:500;">${esc(m.title) || 'Team Member'}</div>
                        <span class="badge ${roleCls}" style="margin-top:2px;">${roleName} • ${esc(m.department) || 'General'}</span>
                    </td>
                    <td><span style="font-size:0.85rem;color:var(--text-secondary);">${esc(m.manager_name) || '—'}</span></td>
                    <td><span class="badge badge-progress" style="font-size:0.85rem;">${m.active_tasks}</span></td>
                    <td><span style="font-weight:600;color:var(--success);">${m.completed_tasks}</span> / ${m.total_tasks}</td>
                    <td>
                        <div style="display:flex;align-items:center;gap:8px;">
                            <div class="progress-bar-bg" style="width:60px;"><div class="progress-bar-fill" style="width:${m.efficiency}%"></div></div>
                            <span style="font-size:0.8rem;font-weight:600;">${m.efficiency}%</span>
                        </div>
                    </td>
                    <td><span style="font-weight:600;color:var(--info);">${parseFloat(m.logged_hours || 0).toFixed(1)} hrs</span></td>
                </tr>`;
            }).join('');
        }
    } catch (e) {
        console.error('Executive Dashboard load error:', e);
        showToast('Failed to load Executive Dashboard');
    }
}

// ---- Team Management & Hierarchy ----
let allUsersList = [];
async function loadTeamManagement() {
    const tbody = $('#team-roster-tbody');
    if (tbody) tbody.innerHTML = loadingStateRow(6);
    try {
        const res = await fetch(API('/api/users'));
        if (!res.ok) throw new Error('Failed to fetch users');
        allUsersList = await res.json();

        // Also populate assignee select dropdowns across the app
        populateAssigneeDropdowns(allUsersList);

        if (allUsersList.length === 0) {
            tbody.innerHTML = emptyStateRow(6, '👥', currentLang === 'th' ? 'ยังไม่มีสมาชิกในทีม' : 'No team members yet', currentLang === 'th' ? 'เริ่มต้นด้วยการเพิ่มสมาชิกคนแรกของคุณ' : 'Get started by adding your first team member', currentLang === 'th' ? '+ เพิ่มสมาชิก' : '+ Add Team Member', 'openCreateUserModal()');
            return;
        }
        tbody.innerHTML = allUsersList.map(u => {
            const initials = getInitials(u.name);
            const roleCls = u.role === 'admin' ? 'badge-admin' : u.role === 'manager' ? 'badge-manager' : 'badge-member';
            const roleName = u.role === 'admin' ? t('role_admin') : u.role === 'manager' ? t('role_manager') : t('role_member');
            const manager = allUsersList.find(m => m.id === u.manager_id);
            const managerName = manager ? manager.name : 'Top Level / CEO';
            const isSelf = currentUser && u.id === currentUser.id;
            const isMember = currentUser && currentUser.role === 'member';
            const isManager = currentUser && currentUser.role === 'manager';
            const isProtectedFromMember = isMember && !isSelf;
            const isProtectedFromManager = isManager && (u.role === 'admin' || (u.role === 'manager' && !isSelf));
            const isLocked = isProtectedFromMember || isProtectedFromManager;
            return `<tr>
                <td>
                    <div style="display:flex; align-items:center; gap:10px;">
                        ${u.avatar_url && u.avatar_url.trim() !== '' ?
                            `<img src="${esc(u.avatar_url)}" style="width:32px;height:32px;border-radius:50%;object-fit:cover;">` :
                            `<div style="width:32px;height:32px;border-radius:50%;background:${u.avatar_color || '#2684FF'};color:#fff;font-weight:700;font-size:0.8rem;display:flex;align-items:center;justify-content:center;">${esc(initials)}</div>`
                        }
                        <div>
                            <div style="font-weight:600;color:var(--text-primary);">${esc(u.name)} ${isSelf ? '<span style="color:var(--accent);font-size:0.75rem;">(You)</span>' : ''}</div>
                            <div style="font-size:0.75rem;color:var(--text-muted);">${esc(u.email)}</div>
                        </div>
                    </div>
                </td>
                <td style="font-weight:500;">${esc(u.title) || 'Team Member'}</td>
                <td><span style="color:var(--text-secondary);">${esc(u.department) || 'General'}</span></td>
                <td><span class="badge ${roleCls}">${roleName}</span></td>
                <td style="font-size:0.85rem;color:var(--text-secondary);">👤 ${esc(managerName)}</td>
                <td style="text-align:right;">
                    ${isLocked ? `
                        <span style="font-size:0.78rem; color:var(--text-muted); background:var(--bg-elevated); padding:4px 10px; border-radius:4px; border:1px solid var(--border-color);">🔒 Protected</span>
                    ` : `
                        <button class="btn btn-secondary" style="padding:4px 10px;font-size:0.78rem;" onclick="openResetPasswordModal('${u.id}')">${t('btn_reset_pass') || '🔑 Password'}</button>
                        <button class="btn btn-secondary" style="padding:4px 10px;font-size:0.78rem;margin-left:6px;" onclick="openEditUserModal('${u.id}')">✏️ Edit</button>
                        ${!isSelf ? `<button class="btn btn-danger" style="padding:4px 10px;font-size:0.78rem;margin-left:6px;" onclick="deleteUser('${u.id}', '${escJsAttr(u.name)}')">🗑️</button>` : ''}
                    `}
                </td>
            </tr>`;
        }).join('');
    } catch (e) {
        console.error('Team management load error:', e);
        showToast('Failed to load team list');
    }
}

function populateAssigneeDropdowns(users) {
    const opts = `<option value="">${t('unassigned')}</option>` + users.map(u => `<option value="${u.id}" data-name="${esc(u.name)}" data-avatar="${u.avatar_color || '#2684FF'}">${esc(u.name)} (${esc(u.title || u.role || 'Member')})</option>`).join('');
    const taskAssignee = $('#task-assignee');
    if (taskAssignee) taskAssignee.innerHTML = opts;
}

function openCreateUserModal() {
    $('#user-edit-id').value = '';
    $('#user-form').reset();
    $('#user-modal-title').textContent = t('modal_add_user');
    $('#user-password-group').style.display = 'block';
    $('#user-password').required = true;
    $('#user-password').value = 'admin123';
    const passLabel = document.querySelector('label[for="user-password"]');
    if (passLabel) passLabel.textContent = t('password_label');
    
    // Restrict role options if Manager
    const roleSel = $('#user-role');
    if (roleSel) {
        if (currentUser && currentUser.role === 'manager') {
            roleSel.innerHTML = `<option value="member">${t('role_member') || 'Team Member'}</option>`;
        } else {
            roleSel.innerHTML = `
                <option value="member">${t('role_member') || 'Team Member'}</option>
                <option value="manager">${t('role_manager') || 'Manager / Lead'}</option>
                <option value="admin">${t('role_admin') || 'Administrator'}</option>
            `;
        }
    }

    // Populate manager dropdown
    const managerSel = $('#user-manager');
    managerSel.innerHTML = `<option value="">Top Level / CEO</option>` + allUsersList.map(u => `<option value="${u.id}">${esc(u.name)} (${esc(u.title || u.role)})</option>`).join('');
    openModal('user-modal');
}

function openEditUserModal(userId) {
    const u = allUsersList.find(x => x.id === userId);
    if (!u) return;

    if (currentUser && currentUser.role === 'manager' && u.role === 'admin') {
        showToast(currentLang === 'th' ? 'หัวหน้าทีมไม่สามารถแก้ไขข้อมูลหรือสิทธิ์ของ Admin ได้' : 'Managers cannot edit Admin accounts');
        return;
    }

    $('#user-edit-id').value = u.id;
    $('#user-name').value = u.name || '';
    $('#user-email').value = u.email || '';
    $('#user-title').value = u.title || '';
    $('#user-dept').value = u.department || 'Engineering';

    const roleSel = $('#user-role');
    if (roleSel) {
        if (currentUser && currentUser.role === 'manager') {
            roleSel.innerHTML = `<option value="member">${t('role_member') || 'Team Member'}</option>`;
        } else {
            roleSel.innerHTML = `
                <option value="member">${t('role_member') || 'Team Member'}</option>
                <option value="manager">${t('role_manager') || 'Manager / Lead'}</option>
                <option value="admin">${t('role_admin') || 'Administrator'}</option>
            `;
        }
    }
    $('#user-role').value = u.role || 'member';
    
    $('#user-modal-title').textContent = t('modal_edit_user');
    $('#user-password-group').style.display = 'block';
    $('#user-password').required = false;
    $('#user-password').value = '';
    const passLabel = document.querySelector('label[for="user-password"]');
    if (passLabel) passLabel.textContent = currentLang === 'th' ? 'รหัสผ่านใหม่ (เว้นว่างไว้หากไม่เปลี่ยน)' : 'New Password (leave blank to keep current)';
    
    const managerSel = $('#user-manager');
    managerSel.innerHTML = `<option value="">Top Level / CEO</option>` + allUsersList.filter(x => x.id !== u.id).map(x => `<option value="${x.id}" ${x.id === u.manager_id ? 'selected' : ''}>${esc(x.name)} (${esc(x.title || x.role)})</option>`).join('');
    openModal('user-modal');
}

function openResetPasswordModal(userId) {
    const u = allUsersList.find(x => x.id === userId);
    if (!u) return;

    if (currentUser && currentUser.role === 'manager' && u.role === 'admin') {
        showToast(currentLang === 'th' ? 'หัวหน้าทีมไม่สามารถเปลี่ยนรหัสผ่านของ Admin ได้' : 'Managers cannot reset Admin passwords');
        return;
    }

    $('#reset-user-id').value = u.id;
    $('#reset-new-password').value = '';
    const desc = currentLang === 'th' ? `กำหนดรหัสผ่านใหม่สำหรับการเข้าสู่ระบบของ <b>${esc(u.name)}</b> (${esc(u.email)})` : `Set a new login password for <b>${esc(u.name)}</b> (${esc(u.email)})`;
    $('#reset-user-desc').innerHTML = desc;
    openModal('reset-password-modal');
}

const resetPassForm = $('#reset-password-form');
if (resetPassForm) {
    resetPassForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const userId = $('#reset-user-id').value;
        const password = $('#reset-new-password').value.trim();
        if (!password || password.length < 6) {
            showToast(currentLang === 'th' ? 'รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร' : 'Password must be at least 6 characters');
            return;
        }
        try {
            const res = await fetch(API(`/api/users/${userId}/password`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ password })
            });
            const data = await res.json();
            if (data.success) {
                closeModal('reset-password-modal');
                showToast(currentLang === 'th' ? '🔑 รีเซ็ตและแก้ไขรหัสผ่านเรียบร้อยแล้ว' : '🔑 Password reset successfully');
            } else {
                showToast(data.error || 'Failed to update password');
            }
        } catch (err) {
            showToast('Server error during password reset');
        }
    });
}

$('#user-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const editId = $('#user-edit-id').value;
    const name = $('#user-name').value.trim();
    const email = $('#user-email').value.trim();
    const title = $('#user-title').value.trim();
    const department = $('#user-dept').value;
    const role = $('#user-role').value;
    const manager_id = $('#user-manager').value;

    // Email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        showToast(currentLang === 'th' ? 'กรุณากรอกรูปแบบอีเมลให้ถูกต้อง' : 'Please enter a valid email address');
        return;
    }
    
    try {
        if (!editId) {
            // Create user validation
            const password = $('#user-password').value || 'admin123';
            if (password.length < 6) {
                showToast(currentLang === 'th' ? 'รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร' : 'Password must be at least 6 characters');
                return;
            }
            
            const colors = ['#2684FF', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#3b82f6'];
            const avatar_color = colors[Math.floor(Math.random() * colors.length)];
            const res = await fetch(API('/api/users'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, email, password, avatar_color })
            });
            const data = await res.json();
            if (data.success && data.user) {
                // Update hierarchy for new user
                await fetch(API(`/api/users/${data.user.id}/hierarchy`), {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ role, manager_id, department, title })
                });
                showToast('Team member added successfully');
            } else {
                showToast(data.error || 'Failed to add user');
                return;
            }
        } else {
            // Update existing user validation
            const password = $('#user-password').value.trim();
            if (password.length > 0 && password.length < 6) {
                showToast(currentLang === 'th' ? 'รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร' : 'Password must be at least 6 characters');
                return;
            }
            
            const res = await fetch(API(`/api/users/${editId}`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, email, password, role, manager_id, department, title })
            });
            const data = await res.json();
            if (res.ok && !data.error) {
                showToast(currentLang === 'th' ? 'บันทึกการแก้ไขข้อมูลสิทธิ์และรหัสผ่านเรียบร้อย' : 'Member updated successfully');
                if (currentUser && currentUser.id === editId) {
                    currentUser.name = name;
                    currentUser.email = email;
                    currentUser.role = role;
                    currentUser.title = title;
                    currentUser.department = department;
                    updateUserUI();
                }
            } else {
                showToast(data.error || 'Failed to update user');
                return;
            }
        }
        closeModal('user-modal');
        loadTeamManagement();
    } catch (err) {
        showToast('Error saving user');
    }
});

async function deleteUser(id, name) {
    const target = allUsersList.find(x => x.id === id);
    if (currentUser && currentUser.role === 'manager' && target && target.role === 'admin') {
        showToast(currentLang === 'th' ? 'หัวหน้าทีมไม่สามารถลบบัญชี Admin ได้' : 'Managers cannot delete Admin accounts');
        return;
    }
    if (!(await confirmDialog(`Are you sure you want to remove ${name} from the team?`))) return;
    try {
        const res = await fetch(API(`/api/users/${id}`), { method: 'DELETE' });
        if (res.ok) {
            showToast('Member removed');
            loadTeamManagement();
        } else {
            showToast('Failed to remove member');
        }
    } catch (e) {
        showToast('Error removing member');
    }
}

// ---- Projects ----
async function loadProjects() {
    const grid = $('#projects-grid-list');
    if (grid) grid.innerHTML = loadingState();
    try {
        const projects = await fetch(API(`/api/projects?workspace_id=${currentWorkspaceId}`)).then(r => r.json());
        grid.innerHTML = projects.map(p => `
            <div class="project-card" onclick="openProject('${p.id}', '${escJsAttr(p.name)}')">
                <div class="project-card-title">${esc(p.name)}</div>
                <div class="project-card-desc">${esc(p.desc) || 'No description'}</div>
                <div class="project-card-meta">
                    <span>${t('created_on')} ${formatDate(p.created_at)}</span>
                    <span style="color:var(--accent); font-weight:600;">${t('open_btn')}</span>
                </div>
            </div>
        `).join('') || `<div style="grid-column:1/-1;">${emptyState('📁', currentLang === 'th' ? 'ยังไม่มีโครงการในพื้นที่ทำงานนี้' : 'No projects in this workspace', currentLang === 'th' ? 'เริ่มต้นด้วยการสร้างโปรเจกต์แรกของคุณ' : 'Get started by creating your first project', currentLang === 'th' ? '+ สร้างโครงการใหม่' : '+ New Project', "openModal('new-project-modal')")}</div>`;
    } catch (e) { showToast('Failed to load projects'); }
}

async function openProject(projectId, projectName) {
    currentProjectId = projectId;
    switchScreen('screen-project-workspace');
    $$('.menu-item').forEach(m => m.classList.remove('active'));
    
    // Show breadcrumb
    $('#nav-breadcrumb').style.display = 'inline-flex';
    $('#breadcrumb-project-name').textContent = projectName || 'Project';
    $('#current-screen-title').textContent = '';
    
    toggleMobileWorkspaceSelect(false);
    
    // Reset tabs to overview
    $$('.tab-btn').forEach(b => b.classList.remove('active'));
    $$('.tab-btn')[0].classList.add('active');
    $$('.workspace-pane').forEach(p => p.classList.remove('active'));
    $('#pane-dashboard').classList.add('active');
    loadProjectOverview();
}

async function loadProjectOverview() {
    if (!currentProjectId) return;
    try {
        const [tasks, milestones] = await Promise.all([
            fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json()),
            fetch(API(`/api/milestones?project_id=${currentProjectId}`)).then(r => r.json())
        ]);
        const todo = tasks.filter(t => t.status === 'todo').length;
        const prog = tasks.filter(t => t.status === 'in_progress').length;
        const done = tasks.filter(t => t.status === 'completed').length;
        $('#w-todo-count').textContent = todo;
        $('#w-progress-count').textContent = prog;
        $('#w-done-count').textContent = done;
        // Milestones
        const msEl = $('#w-milestone-progress-list');
        msEl.innerHTML = milestones.map(m => {
            const mTasks = tasks.filter(t => t.milestone_id === m.id);
            const mDone = mTasks.filter(t => t.status === 'completed').length;
            const pct = m.completed ? 100 : (mTasks.length > 0 ? Math.round((mDone / mTasks.length) * 100) : 0);
            return `<div class="progress-item">
                <div class="progress-item-header"><span class="progress-item-title">${m.completed ? '✓ ' : ''}${esc(m.name)}</span><span>${pct}%</span></div>
                <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct}%"></div></div>
            </div>`;
        }).join('') || emptyState('🎯', t('no_milestones'));
    } catch (e) {}
}

// ---- Tasks ----
async function loadProjectTasks() {
    if (!currentProjectId) return;
    const container = $('#w-task-list-container');
    if (container) container.innerHTML = loadingState();
    try {
        const tasks = await fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json());
        currentProjectTasksList = tasks || [];
        renderFilteredProjectTasks();

        const milestones = await fetch(API(`/api/milestones?project_id=${currentProjectId}`)).then(r => r.json());
        const msSel = $('#task-milestone');
        msSel.innerHTML = '<option value="">None</option>' + milestones.map(m => `<option value="${m.id}">${esc(m.name)}</option>`).join('');
    } catch (e) { showToast('Failed to load tasks'); }
}

function renderFilteredProjectTasks() {
    const container = $('#w-task-list-container');
    if (!container) return;

    const query = ($('#task-search-input')?.value || '').trim().toLowerCase();
    const statusFilter = $('#task-filter-status')?.value || '';
    const priorityFilter = $('#task-filter-priority')?.value || '';

    let tasks = currentProjectTasksList;
    if (query) tasks = tasks.filter(t => (t.name || '').toLowerCase().includes(query));
    if (statusFilter) tasks = tasks.filter(t => t.status === statusFilter);
    if (priorityFilter) tasks = tasks.filter(t => t.priority === priorityFilter);

    if (tasks.length === 0 && currentProjectTasksList.length > 0) {
        container.innerHTML = emptyState('🔍', currentLang === 'th' ? 'ไม่พบงานที่ตรงกับตัวกรอง' : 'No tasks match your filters', currentLang === 'th' ? 'ลองเปลี่ยนคำค้นหาหรือตัวกรอง' : 'Try a different search term or filter', currentLang === 'th' ? 'ล้างตัวกรอง' : 'Clear Filters', 'clearProjectTaskFilters()');
        return;
    }

    container.innerHTML = tasks.map(t => `
        <div class="task-row">
            <div class="task-row-left" style="display:flex;align-items:center;gap:14px;min-width:0;">
                <div class="task-checkbox ${t.status === 'completed' ? 'checked' : ''}" onclick="toggleTaskStatus('${t.id}', '${t.status}')"></div>
                <div class="task-row-details">
                    <h4 class="${t.status === 'completed' ? 'completed' : ''}" onclick="openTaskDetail('${t.id}')">${esc(t.name)}</h4>
                    <div class="task-row-meta">
                        <span>📅 ${formatDate(t.due)}</span>
                        ${t.recurring_pattern && t.recurring_pattern !== 'none' ? `<span style="color:var(--info);">↻ ${esc(t.recurring_pattern)}</span>` : ''}
                    </div>
                </div>
            </div>
            <div class="task-row-right" style="display:flex;align-items:center;gap:12px;flex-shrink:0;">
                ${t.assignee_name ? `<span class="assignee-pill" title="Assignee"><span class="assignee-avatar-xs" style="background:${t.assignee_avatar || '#2684FF'}">${esc(getInitials(t.assignee_name))}</span>${esc(t.assignee_name)}</span>` : ''}
                ${priorityBadge(t.priority)}
                ${statusBadge(t.status)}
            </div>
        </div>
    `).join('') || emptyState('📋', currentLang === 'th' ? 'ยังไม่มีงานในโครงการนี้' : 'No tasks yet', currentLang === 'th' ? 'เริ่มต้นด้วยการเพิ่มงานแรกของโปรเจกต์นี้' : 'Get started by adding your first task', currentLang === 'th' ? '+ เพิ่มงานใหม่' : '+ Add Task', 'openAddTaskModal()');
}

function clearProjectTaskFilters() {
    if ($('#task-search-input')) $('#task-search-input').value = '';
    if ($('#task-filter-status')) $('#task-filter-status').value = '';
    if ($('#task-filter-priority')) $('#task-filter-priority').value = '';
    renderFilteredProjectTasks();
}

async function toggleTaskStatus(taskId, currentStatus) {
    const newStatus = currentStatus === 'completed' ? 'todo' : currentStatus === 'in_progress' ? 'completed' : 'in_progress';
    try {
        await fetch(API(`/api/tasks/${taskId}/status`), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status: newStatus })
        });
        showToast(`Task → ${newStatus.replace('_', ' ')}`);
        const activeTab = $('.tab-btn.active');
        if (activeTab) {
            const tab = activeTab.dataset.tab;
            if (tab === 'tasks') loadProjectTasks();
            else if (tab === 'board') loadKanban();
            else if (tab === 'dashboard') loadProjectOverview();
        }
    } catch (e) { showToast('Failed to update task'); }
}

// ---- Kanban ----
async function loadKanban() {
    if (!currentProjectId) return;
    ['#kanban-column-todo', '#kanban-column-progress', '#kanban-column-completed'].forEach(sel => {
        if ($(sel)) $(sel).innerHTML = loadingState();
    });
    try {
        const tasks = await fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json());
        const todo = tasks.filter(t => t.status === 'todo');
        const progress = tasks.filter(t => t.status === 'in_progress');
        const done = tasks.filter(t => t.status === 'completed');

        $('#kanban-todo-count').textContent = todo.length;
        $('#kanban-progress-count').textContent = progress.length;
        $('#kanban-completed-count').textContent = done.length;

        const renderCards = (arr) => arr.map(t => `
            <div class="kanban-card" draggable="true" data-task-id="${t.id}"
                 ondragstart="handleDragStart(event)" ondragend="handleDragEnd(event)">
                <div class="kanban-card-title" onclick="openTaskDetail('${t.id}')">${esc(t.name)}</div>
                <div class="kanban-card-meta" style="display:flex;justify-content:space-between;align-items:center;margin-top:12px;font-size:0.75rem;">
                    <div style="display:flex;align-items:center;gap:6px;">
                        ${priorityBadge(t.priority)}
                        ${t.assignee_name ? `<span class="assignee-pill" title="Assignee"><span class="assignee-avatar-xs" style="background:${t.assignee_avatar || '#2684FF'}">${esc(getInitials(t.assignee_name))}</span>${esc(t.assignee_name.split(' ')[0])}</span>` : ''}
                    </div>
                    <span style="color:var(--text-secondary)">📅 ${formatDate(t.due)}</span>
                </div>
            </div>
        `).join('');

        $('#kanban-column-todo').innerHTML = renderCards(todo) || `<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:0.82rem;">${t('drag_here')}</div>`;
        $('#kanban-column-progress').innerHTML = renderCards(progress) || `<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:0.82rem;">${t('drag_here')}</div>`;
        $('#kanban-column-completed').innerHTML = renderCards(done) || `<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:0.82rem;">${t('drag_here')}</div>`;
    } catch (e) { showToast('Failed to load board'); }
}

function handleDragStart(e) {
    e.target.classList.add('dragging');
    e.dataTransfer.setData('text/plain', e.target.dataset.taskId);
}
function handleDragEnd(e) { e.target.classList.remove('dragging'); }
function allowDrop(e) { e.preventDefault(); }
async function handleDrop(e) {
    e.preventDefault();
    const taskId = e.dataTransfer.getData('text/plain');
    const col = e.target.closest('.kanban-column');
    if (!col) return;
    const status = col.dataset.status;
    try {
        await fetch(API(`/api/tasks/${taskId}/status`), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status })
        });
        loadKanban();
        showToast(currentLang === 'th' ? `ย้ายงานไปยัง ${status}` : `Task moved to ${status.replace('_', ' ')}`);
    } catch (e) { showToast('Drop failed'); }
}

// ---- Gantt ----
async function renderGantt() {
    if (!currentProjectId) return;
    if ($('#gantt-workspace-render')) $('#gantt-workspace-render').innerHTML = loadingState();
    try {
        const tasks = await fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json());
        const tasksWithDue = tasks.filter(t => t.due);
        if (tasksWithDue.length === 0) {
            $('#gantt-workspace-render').innerHTML = emptyState('📊', t('no_gantt_tasks'));
            return;
        }

        const dueDates = tasksWithDue.map(t => new Date(t.due));
        const startDates = tasksWithDue
            .filter(t => t.start_date && !isNaN(new Date(t.start_date)))
            .map(t => new Date(t.start_date));
        const allDates = dueDates.concat(startDates);
        const minDate = new Date(Math.min(...allDates));
        const maxDate = new Date(Math.max(...allDates));
        minDate.setDate(minDate.getDate() - 3);
        maxDate.setDate(maxDate.getDate() + 7);

        const colWidth = ganttZoom === 'day' ? 46 : ganttZoom === 'week' ? 28 : 16;
        const daysBetween = Math.ceil((maxDate - minDate) / (1000 * 60 * 60 * 24)) + 1;
        const totalWidth = daysBetween * colWidth;

        let taskColHTML = `<div class="gantt-task-cell-header">${t('th_task')}</div>`;
        tasksWithDue.forEach(t => {
            taskColHTML += `<div class="gantt-task-cell" title="${esc(t.name)}">${esc(t.name)}</div>`;
        });

        let headerHTML = '';
        for (let d = new Date(minDate); d <= maxDate; d.setDate(d.getDate() + 1)) {
            const label = ganttZoom === 'day' ? d.getDate() : ganttZoom === 'week' ? `W${Math.ceil(d.getDate() / 7)}` : d.toLocaleDateString(currentLang === 'th' ? 'th-TH' : 'en-US', { month: 'short' });
            const isToday = d.toDateString() === new Date().toDateString();
            headerHTML += `<div class="gantt-timeline-col" style="width:${colWidth}px;min-width:${colWidth}px;${isToday ? 'background:var(--accent-subtle);color:var(--accent);font-weight:700;' : ''}">${label}</div>`;
        }

        let rowsHTML = '';
        tasksWithDue.forEach(t => {
            let gridCols = '';
            for (let d = new Date(minDate); d <= maxDate; d.setDate(d.getDate() + 1)) {
                gridCols += `<div class="gantt-grid-col" style="width:${colWidth}px;min-width:${colWidth}px;"></div>`;
            }
            const dueDate = new Date(t.due);
            const hasStart = t.start_date && !isNaN(new Date(t.start_date)) && new Date(t.start_date) <= dueDate;

            let barLeft, barWidth, barTitle;
            if (hasStart) {
                const startDate = new Date(t.start_date);
                const startOffset = Math.max(0, Math.round((startDate - minDate) / (1000 * 60 * 60 * 24)));
                const durationDays = Math.round((dueDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
                barLeft = startOffset * colWidth;
                barWidth = Math.max(colWidth, durationDays * colWidth);
                barTitle = `${esc(t.name)} (${formatDate(t.start_date)} → ${formatDate(t.due)})`;
            } else {
                const startOffset = Math.max(0, Math.ceil((dueDate - minDate) / (1000 * 60 * 60 * 24)) - 2);
                barLeft = startOffset * colWidth;
                barWidth = colWidth * 3;
                barTitle = esc(t.name);
            }
            const barColor = t.status === 'completed' ? 'linear-gradient(135deg,#10b981,#059669)' : t.status === 'in_progress' ? 'linear-gradient(135deg,var(--accent),var(--accent-hover))' : 'linear-gradient(135deg,#64748b,#475569)';
            rowsHTML += `<div class="gantt-grid-row" style="width:${totalWidth}px;">${gridCols}<div class="gantt-bar-item" style="left:${barLeft}px;width:${barWidth}px;background:${barColor};" title="${barTitle}">${esc(t.name)}</div></div>`;
        });

        $('#gantt-workspace-render').innerHTML = `
            <div class="gantt-tasks-column">${taskColHTML}</div>
            <div class="gantt-timeline-wrapper">
                <div class="gantt-timeline-header" style="width:${totalWidth}px;">${headerHTML}</div>
                ${rowsHTML}
            </div>
        `;
    } catch (e) { showToast('Gantt rendering failed'); }
}

// ---- Milestones ----
async function loadMilestones() {
    if (!currentProjectId) return;
    const container = $('#w-milestones-container');
    if (container) container.innerHTML = loadingState();
    try {
        const [milestones, tasks] = await Promise.all([
            fetch(API(`/api/milestones?project_id=${currentProjectId}`)).then(r => r.json()),
            fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json())
        ]);
        container.innerHTML = milestones.map(m => {
            const mTasks = tasks.filter(t => t.milestone_id === m.id);
            const done = mTasks.filter(t => t.status === 'completed').length;
            const pct = m.completed ? 100 : (mTasks.length > 0 ? Math.round((done / mTasks.length) * 100) : 0);
            return `<div class="milestone-card">
                <div class="milestone-card-left" style="display:flex;align-items:center;gap:14px;">
                    <div class="milestone-checkbox ${m.completed ? 'completed' : ''}" onclick="toggleMilestone('${m.id}', ${m.completed ? 0 : 1})"></div>
                    <div class="milestone-info">
                        <h3 class="${m.completed ? 'completed' : ''}" style="font-weight:600;font-size:1rem;">${esc(m.name)}</h3>
                        <div class="milestone-date" style="font-size:0.8rem;color:var(--text-secondary);margin-top:4px;">🎯 ${formatDate(m.date)}</div>
                    </div>
                </div>
                <div class="milestone-progress" style="display:flex;align-items:center;gap:12px;">
                    <div class="milestone-progress-bar"><div class="milestone-progress-fill" style="width:${pct}%"></div></div>
                    <span style="font-size:0.82rem;color:var(--text-secondary);min-width:38px;text-align:right;font-weight:600;">${pct}%</span>
                    <button type="button" class="btn-icon" title="Edit Milestone" onclick="openEditMilestoneModal('${m.id}')">✏️</button>
                </div>
            </div>`;
        }).join('') || emptyState('🎯', t('no_milestones'), null, currentLang === 'th' ? '+ เพิ่มหมุดหมาย' : '+ Add Milestone', "openAddMilestoneModal()");
    } catch (e) { showToast('Failed to load milestones'); }
}

async function toggleMilestone(id, completed) {
    try {
        await fetch(API(`/api/milestones/${id}/status`), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ completed })
        });
        loadMilestones();
        showToast(completed ? 'Milestone completed! 🎉' : 'Milestone reopened');
    } catch (e) {}
}

// ---- Timesheets ----
async function loadTimesheets() {
    if (!currentProjectId) return;
    const tbody = $('#w-timesheet-tbody');
    if (tbody) tbody.innerHTML = loadingStateRow(7);
    try {
        const [timesheets, tasks] = await Promise.all([
            fetch(API(`/api/timesheets?project_id=${currentProjectId}`)).then(r => r.json()),
            fetch(API(`/api/tasks?project_id=${currentProjectId}`)).then(r => r.json())
        ]);
        tbody.innerHTML = timesheets.map(ts => {
            const task = tasks.find(t => t.id === ts.task_id);
            const taskLabel = task ? task.name : (ts.task_id && !ts.task_id.startsWith('t_') ? ts.task_id : (currentLang === 'th' ? 'งานทั่วไป / ทั่วไป' : 'General Work'));
            return `<tr>
                <td>${formatDate(ts.date)}</td>
                <td style="font-weight:500;">${esc(taskLabel)}</td>
                <td>${esc(ts.notes) || '—'}</td>
                <td>${esc(ts.user) || esc(currentUser.name)}</td>
                <td>${ts.billed ? `<span class="badge badge-done">${t('bill_billed')}</span>` : `<span class="badge badge-todo">${t('bill_unbilled')}</span>`}</td>
                <td style="font-weight:700;color:var(--accent);">${ts.hours}h</td>
                <td>${ts.billed ? '' : `<button type="button" class="btn-icon" title="Edit Time Entry" onclick="openEditTimesheetModal('${ts.id}')">✏️</button>`}</td>
            </tr>`;
        }).join('') || emptyStateRow(7, '⏱️', t('no_timesheets'), null, currentLang === 'th' ? '+ บันทึกเวลา' : '+ Log Time', "openAddTimesheetModal()");
        
        const taskSel = $('#timesheet-task');
        taskSel.innerHTML = tasks.map(t => `<option value="${t.id}">${esc(t.name)}</option>`).join('');
    } catch (e) { showToast('Failed to load timesheets'); }
}

// ---- Billing / Invoices ----
let currentInvoiceLineItems = [];
let availableUnbilledTimesheets = [];

async function loadBilling() {
    if (!currentProjectId) return;
    const listEl = $('#w-invoice-list-container');
    if (listEl) listEl.innerHTML = loadingState();
    try {
        const [invoices, timesheets] = await Promise.all([
            fetch(API(`/api/invoices?project_id=${currentProjectId}`)).then(r => r.json()),
            fetch(API(`/api/timesheets?project_id=${currentProjectId}`)).then(r => r.json())
        ]);
        listEl.innerHTML = invoices.map(inv => {
            const status = inv.status || 'unpaid';
            const badgeClass = `badge badge-${status}`;
            const statusText = status === 'paid' ? (currentLang === 'th' ? '🟢 ชำระแล้ว' : '🟢 Paid') :
                               status === 'overdue' ? (currentLang === 'th' ? '🔴 เกินกำหนด' : '🔴 Overdue') :
                               status === 'cancelled' ? (currentLang === 'th' ? '⚪ ยกเลิก' : '⚪ Cancelled') :
                               (currentLang === 'th' ? '🟡 รอชำระ' : '🟡 Unpaid');
            return `
            <div class="invoice-item" style="display:flex;justify-content:space-between;align-items:center;padding:16px;background:var(--bg-surface);border:1px solid var(--border-color);border-radius:var(--radius-md);margin-bottom:12px;">
                <div style="cursor:pointer;flex-grow:1;" onclick="showInvoicePreview('${inv.id}')">
                    <div style="display:flex;align-items:center;gap:12px;">
                        <span style="font-weight:700;font-size:1.05rem;color:var(--accent);">${esc(inv.invoice_number)}</span>
                        <span class="${badgeClass}" style="font-size:0.75rem;">${statusText}</span>
                    </div>
                    <div style="font-weight:600;font-size:0.92rem;color:var(--text-primary);margin-top:4px;">${esc(inv.client)}</div>
                    <div style="font-size:0.8rem;color:var(--text-secondary);margin-top:2px;">
                        ${inv.client_address ? esc(inv.client_address) + ' — ' : ''}Date: ${formatDate(inv.date)}${inv.due_date ? ' | Due: ' + formatDate(inv.due_date) : ''}
                    </div>
                </div>
                <div style="display:flex;align-items:center;gap:16px;flex-shrink:0;">
                    <div style="text-align:right;">
                        <div style="font-weight:800;font-size:1.2rem;color:${status === 'paid' ? 'var(--success)' : 'var(--text-primary)'};">$${(inv.total_amount || inv.amount || 0).toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                        <div style="font-size:0.75rem;color:var(--text-muted);">${inv.hours || 0} hrs billed</div>
                    </div>
                    <div style="display:flex;gap:6px;" onclick="event.stopPropagation();">
                        <button class="btn-icon" title="View / Print PDF" onclick="showInvoicePreview('${inv.id}')">📄</button>
                        <button class="btn-icon" title="Edit Invoice" onclick="openEditInvoiceModal('${inv.id}')">✏️</button>
                        <button class="btn-icon" title="Toggle Paid Status" onclick="changeInvoiceStatus('${inv.id}', '${status === 'paid' ? 'unpaid' : 'paid'}')">${status === 'paid' ? '↩️' : '✅'}</button>
                        <button class="btn-icon" title="Delete Invoice" style="color:var(--danger);" onclick="deleteInvoice('${inv.id}')">🗑️</button>
                    </div>
                </div>
            </div>
        `;
        }).join('') || emptyState('🧾', t('no_invoices'), null, currentLang === 'th' ? '+ สร้างใบแจ้งหนี้' : '+ New Invoice', 'openCreateInvoiceModal()');

        const unbilled = timesheets.filter(ts => !ts.billed);
        const billed = timesheets.filter(ts => ts.billed);
        const unbilledH = unbilled.reduce((s, t) => s + (t.hours || 0), 0);
        
        const unpaidRev = invoices.filter(i => (i.status || 'unpaid') !== 'paid' && (i.status || 'unpaid') !== 'cancelled').reduce((s, i) => s + (i.total_amount || i.amount || 0), 0);
        const paidRev = invoices.filter(i => i.status === 'paid').reduce((s, i) => s + (i.total_amount || i.amount || 0), 0);

        const rate = getCurrentWorkspaceRate();
        if ($('#bill-unpaid-amount')) $('#bill-unpaid-amount').textContent = `$${unpaidRev.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;
        if ($('#bill-paid-amount')) $('#bill-paid-amount').textContent = `$${paidRev.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;
        if ($('#bill-unbilled-hours')) $('#bill-unbilled-hours').textContent = `${unbilledH.toFixed(1)} hrs ($${(unbilledH * rate).toFixed(2)})`;
        if ($('#bill-rate-display')) $('#bill-rate-display').textContent = `$${rate.toFixed(2)} / hr`;
    } catch (e) { showToast('Failed to load billing'); }
}

async function openCreateInvoiceModal() {
    if (!currentProjectId) return;
    $('#new-invoice-form').reset();
    $('#invoice-edit-id').value = '';
    $('#invoice-modal-title').textContent = t('modal_create_inv');
    $('#invoice-modal-submit-btn').textContent = t('btn_generate_inv');
    if ($('#inv-unbilled-section')) $('#inv-unbilled-section').style.display = '';
    const today = new Date().toISOString().split('T')[0];
    const due = new Date(Date.now() + 15 * 86400000).toISOString().split('T')[0];
    if ($('#invoice-date')) $('#invoice-date').value = today;
    if ($('#invoice-due-date')) $('#invoice-due-date').value = due;
    if ($('#invoice-discount')) $('#invoice-discount').value = 0;
    if ($('#invoice-tax-rate')) $('#invoice-tax-rate').value = 7;
    if ($('#invoice-issuer-name')) $('#invoice-issuer-name').value = localStorage.getItem('ag_issuer_name') || 'AG PROJECTS CO., LTD.';
    if ($('#invoice-issuer-addr')) $('#invoice-issuer-addr').value = localStorage.getItem('ag_issuer_addr') || '123 Tech Tower, Silom Road, Bangrak, Bangkok 10500 | Tax ID: 0105550001234';

    try {
        const rate = getCurrentWorkspaceRate();
        const ts = await fetch(API(`/api/invoices/unbilled-timesheets?project_id=${currentProjectId}`)).then(r => r.json());
        availableUnbilledTimesheets = ts;
        const listEl = $('#inv-unbilled-list');
        if ($('#inv-unbilled-count')) $('#inv-unbilled-count').textContent = `${ts.length} unbilled`;
        
        if (ts.length === 0) {
            listEl.innerHTML = `<div style="color:var(--text-muted);padding:10px;text-align:center;">No unbilled time entries available for this project.</div>`;
        } else {
            listEl.innerHTML = ts.map((t, idx) => `
                <label class="unbilled-item-row" style="cursor:pointer;">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <input type="checkbox" class="unbilled-ts-cb" value="${t.id}" data-idx="${idx}" checked onchange="syncTimesheetsToLineItems()">
                        <div>
                            <div style="font-weight:600;color:var(--text-primary);">${esc(t.task_name)} (${t.hours} hrs)</div>
                            <div style="font-size:0.75rem;color:var(--text-secondary);">${esc(t.notes) || 'Time log'} — ${esc(t.user)} on ${formatDate(t.date)}</div>
                        </div>
                    </div>
                    <div style="font-weight:700;color:var(--success);">$${((t.hours || 0) * rate).toFixed(2)}</div>
                </label>
            `).join('');
        }
        syncTimesheetsToLineItems();
        openModal('new-invoice-modal');
    } catch (e) { showToast('Failed to load unbilled timesheets'); }
}

async function openEditInvoiceModal(invoiceId) {
    try {
        const invoices = await fetch(API(`/api/invoices?project_id=${currentProjectId}`)).then(r => r.json());
        const inv = invoices.find(x => x.id === invoiceId);
        if (!inv) { showToast('Invoice not found'); return; }

        $('#new-invoice-form').reset();
        $('#invoice-edit-id').value = inv.id;
        $('#invoice-modal-title').textContent = t('modal_edit_inv') || (currentLang === 'th' ? 'แก้ไขใบแจ้งหนี้' : 'Edit Invoice');
        $('#invoice-modal-submit-btn').textContent = t('btn_save_changes');
        if ($('#inv-unbilled-section')) $('#inv-unbilled-section').style.display = 'none';

        $('#invoice-client').value = inv.client || '';
        $('#invoice-client-addr').value = inv.client_address || '';
        $('#invoice-date').value = inv.date || '';
        $('#invoice-due-date').value = inv.due_date || '';
        $('#invoice-discount').value = inv.discount || 0;
        $('#invoice-tax-rate').value = inv.tax_rate !== undefined && inv.tax_rate !== null ? String(inv.tax_rate) : '7';
        $('#invoice-notes').value = inv.notes || '';
        $('#invoice-issuer-name').value = localStorage.getItem('ag_issuer_name') || 'AG PROJECTS CO., LTD.';
        $('#invoice-issuer-addr').value = localStorage.getItem('ag_issuer_addr') || '123 Tech Tower, Silom Road, Bangrak, Bangkok 10500 | Tax ID: 0105550001234';

        let items = [];
        try { items = typeof inv.items === 'string' ? JSON.parse(inv.items) : (inv.items || []); } catch (e) { items = []; }
        currentInvoiceLineItems = items.length > 0 ? items : [{ desc: '', hours: 1, rate: 0, amount: 0 }];
        renderInvoiceLineItems();

        openModal('new-invoice-modal');
    } catch (e) { showToast('Failed to load invoice for editing'); }
}

function syncTimesheetsToLineItems() {
    const checkedBoxes = $$('.unbilled-ts-cb:checked');
    currentInvoiceLineItems = [];
    const rate = getCurrentWorkspaceRate();
    checkedBoxes.forEach(cb => {
        const t = availableUnbilledTimesheets[cb.dataset.idx];
        if (t) {
            currentInvoiceLineItems.push({
                desc: `${t.task_name}: ${t.notes || 'Professional services'} (${t.user}, ${formatDate(t.date)})`,
                hours: t.hours || 0,
                rate: rate,
                amount: (t.hours || 0) * rate,
                timesheet_id: t.id
            });
        }
    });
    if (currentInvoiceLineItems.length === 0) {
        currentInvoiceLineItems.push({ desc: 'Professional Project Management Services', hours: 1, rate: rate, amount: rate });
    }
    renderInvoiceLineItems();
}

function addInvoiceLineItem() {
    currentInvoiceLineItems.push({ desc: 'Custom Project Milestone / Consultation', hours: 1, rate: 100, amount: 100 });
    renderInvoiceLineItems();
}

function removeInvoiceLineItem(idx) {
    currentInvoiceLineItems.splice(idx, 1);
    if (currentInvoiceLineItems.length === 0) {
        currentInvoiceLineItems.push({ desc: '', hours: 1, rate: 0, amount: 0 });
    }
    renderInvoiceLineItems();
}

function renderInvoiceLineItems() {
    const tbody = $('#inv-items-tbody');
    if (!tbody) return;
    tbody.innerHTML = currentInvoiceLineItems.map((item, idx) => `
        <tr>
            <td style="padding:6px;"><input type="text" class="inv-item-input" value="${esc(item.desc)}" placeholder="Item description..." oninput="updateLineItem(${idx}, 'desc', this.value)"></td>
            <td style="padding:6px;"><input type="number" class="inv-item-input" style="text-align:right;" min="0" step="0.5" value="${item.hours}" oninput="updateLineItem(${idx}, 'hours', this.value)"></td>
            <td style="padding:6px;"><input type="number" class="inv-item-input" style="text-align:right;" min="0" step="1" value="${item.rate}" oninput="updateLineItem(${idx}, 'rate', this.value)"></td>
            <td style="padding:6px;text-align:right;font-weight:600;color:var(--text-primary);">$${(item.amount || 0).toFixed(2)}</td>
            <td style="padding:6px;text-align:center;"><button type="button" class="btn-icon" style="color:var(--danger);font-size:1.1rem;" onclick="removeInvoiceLineItem(${idx})" title="Remove item">&times;</button></td>
        </tr>
    `).join('');
    calculateInvoiceTotals();
}

function updateLineItem(idx, field, val) {
    if (!currentInvoiceLineItems[idx]) return;
    if (field === 'desc') currentInvoiceLineItems[idx].desc = val;
    else if (field === 'hours') {
        currentInvoiceLineItems[idx].hours = parseFloat(val) || 0;
        currentInvoiceLineItems[idx].amount = currentInvoiceLineItems[idx].hours * currentInvoiceLineItems[idx].rate;
    } else if (field === 'rate') {
        currentInvoiceLineItems[idx].rate = parseFloat(val) || 0;
        currentInvoiceLineItems[idx].amount = currentInvoiceLineItems[idx].hours * currentInvoiceLineItems[idx].rate;
    }
    const rowAmtEl = $$('#inv-items-tbody tr')[idx]?.cells[3];
    if (rowAmtEl) rowAmtEl.textContent = `$${currentInvoiceLineItems[idx].amount.toFixed(2)}`;
    calculateInvoiceTotals();
}

function calculateInvoiceTotals() {
    const subtotal = currentInvoiceLineItems.reduce((s, i) => s + (i.amount || 0), 0);
    const discount = parseFloat($('#invoice-discount')?.value) || 0;
    const taxRate = parseFloat($('#invoice-tax-rate')?.value) || 0;
    const taxable = Math.max(0, subtotal - discount);
    const taxAmount = taxable * (taxRate / 100);
    const grandTotal = taxable + taxAmount;

    if ($('#inv-subtotal-display')) $('#inv-subtotal-display').textContent = `$${subtotal.toFixed(2)}`;
    if ($('#inv-grandtotal-display')) $('#inv-grandtotal-display').textContent = `$${grandTotal.toFixed(2)}`;
}

async function changeInvoiceStatus(id, newStatus) {
    try {
        await fetch(API(`/api/invoices/${id}/status`), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status: newStatus })
        });
        showToast(`Invoice status changed to ${newStatus.toUpperCase()}`);
        loadBilling();
    } catch (e) { showToast('Failed to change status'); }
}

async function changeInvoiceStatusFromPreview() {
    const sel = $('#preview-inv-status');
    if (!sel || !sel.dataset.id) return;
    const newStatus = sel.value;
    try {
        await fetch(API(`/api/invoices/${sel.dataset.id}/status`), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status: newStatus })
        });
        showToast(`Invoice status updated to ${newStatus.toUpperCase()}`);
        loadBilling();
    } catch (e) { showToast('Failed to update status'); }
}

async function deleteInvoice(id) {
    if (!(await confirmDialog(currentLang === 'th' ? 'คุณต้องการลบใบแจ้งหนี้นี้ใช่หรือไม่?' : 'Are you sure you want to delete this invoice?'))) return;
    try {
        await fetch(API(`/api/invoices/${id}`), { method: 'DELETE' });
        showToast(currentLang === 'th' ? 'ลบใบแจ้งหนี้เรียบร้อย' : 'Invoice deleted');
        loadBilling();
    } catch (e) { showToast('Failed to delete invoice'); }
}

async function exportInvoicesToCSV() {
    if (!currentProjectId) return;
    try {
        const res = await fetch(API(`/api/invoices/export?project_id=${currentProjectId}`));
        if (!res.ok) throw new Error('Export failed');
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'invoices_export.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast(currentLang === 'th' ? 'ดาวน์โหลดไฟล์ใบแจ้งหนี้ CSV เรียบร้อย' : 'Invoices exported to CSV');
    } catch (e) {
        console.error('Export error:', e);
        showToast('Failed to export invoices');
    }
}

async function showInvoicePreview(invoiceId) {
    try {
        const invoices = await fetch(API(`/api/invoices?project_id=${currentProjectId}`)).then(r => r.json());
        const inv = invoices.find(i => i.id === invoiceId);
        if (!inv) return;
        
        const sel = $('#preview-inv-status');
        if (sel) {
            sel.value = inv.status || 'unpaid';
            sel.dataset.id = inv.id;
        }

        let items = [];
        try { items = typeof inv.items === 'string' ? JSON.parse(inv.items) : (inv.items || []); } catch (e) { items = []; }
        if (!items || items.length === 0) {
            items = [{ desc: 'Professional Project Management Services', hours: inv.hours || 8, rate: 50, amount: inv.amount || 400 }];
        }

        const subtotal = inv.amount || items.reduce((s, i) => s + (i.amount || 0), 0);
        const discount = inv.discount || 0;
        const taxRate = inv.tax_rate !== undefined ? inv.tax_rate : 7;
        const taxAmount = inv.tax_amount !== undefined ? inv.tax_amount : (subtotal - discount) * (taxRate / 100);
        const totalAmount = inv.total_amount !== undefined ? inv.total_amount : (subtotal - discount + taxAmount);

        const statusText = inv.status === 'paid' ? 'PAID / ชำระเงินแล้ว' : inv.status === 'overdue' ? 'OVERDUE / เกินกำหนด' : inv.status === 'cancelled' ? 'CANCELLED / ยกเลิก' : 'UNPAID / รอการชำระเงิน';
        const statusColor = inv.status === 'paid' ? '#10b981' : inv.status === 'overdue' ? '#ef4444' : inv.status === 'cancelled' ? '#94a3b8' : '#f59e0b';

        const issuerName = localStorage.getItem('ag_issuer_name') || 'AG PROJECTS CO., LTD.';
        const issuerAddr = localStorage.getItem('ag_issuer_addr') || '123 Tech Tower, Silom Road, Bangrak, Bangkok 10500\nTax ID: 0105550001234 (Head Office) | Tel: 02-111-2222';
        const issuerAddrFormatted = issuerAddr.replace(/\n/g, '<br>');

        const area = $('#invoice-render-area');
        area.innerHTML = `
            <div class="invoice-preview-wrapper" style="padding:40px;background:#ffffff;color:#0f172a;border-radius:12px;font-family:'Inter',sans-serif;position:relative;">
                <!-- Stamp Badge -->
                <div style="position:absolute;top:35px;right:40px;border:3px solid ${statusColor};color:${statusColor};padding:6px 16px;border-radius:8px;font-weight:800;font-size:1.1rem;letter-spacing:1px;text-transform:uppercase;transform:rotate(-5deg);opacity:0.9;">
                    ${statusText}
                </div>

                <div style="display:flex;justify-content:space-between;border-bottom:2px solid #cbd5e1;padding-bottom:24px;margin-bottom:28px;">
                    <div>
                        <div style="font-weight:800;font-size:1.5rem;color:#2684FF;letter-spacing:-0.5px;">${issuerName}</div>
                        <div style="font-size:0.85rem;color:#475569;margin-top:4px;line-height:1.5;">
                            ${issuerAddrFormatted}
                        </div>
                    </div>
                    <div style="text-align:right;margin-top:50px;">
                        <h2 style="font-size:1.8rem;color:#0f172a;font-weight:800;margin:0;">TAX INVOICE</h2>
                        <div style="font-weight:700;font-size:1.1rem;color:#2684FF;margin-top:2px;"># ${esc(inv.invoice_number)}</div>
                    </div>
                </div>

                <div style="display:grid;grid-template-columns:1.2fr 0.8fr;gap:24px;margin-bottom:32px;background:#f8fafc;padding:20px;border-radius:8px;border:1px solid #e2e8f0;">
                    <div>
                        <span style="font-weight:700;color:#64748b;font-size:0.75rem;text-transform:uppercase;letter-spacing:0.5px;">BILLED TO / ลูกค้า:</span>
                        <div style="font-size:1.15rem;font-weight:700;color:#0f172a;margin-top:4px;">${esc(inv.client)}</div>
                        <div style="font-size:0.88rem;color:#475569;margin-top:4px;line-height:1.4;">${esc(inv.client_address) || 'Address not specified'}</div>
                    </div>
                    <div style="display:flex;flex-direction:column;gap:8px;font-size:0.88rem;border-left:1px solid #e2e8f0;padding-left:20px;">
                        <div style="display:flex;justify-content:space-between;"><span style="color:#64748b;">Issue Date / วันที่ออก:</span><strong style="color:#0f172a;">${formatDate(inv.date)}</strong></div>
                        <div style="display:flex;justify-content:space-between;"><span style="color:#64748b;">Due Date / ครบกำหนด:</span><strong style="color:${inv.status === 'overdue' ? '#ef4444' : '#0f172a'};">${inv.due_date ? formatDate(inv.due_date) : 'On Receipt'}</strong></div>
                        ${inv.paid_date ? `<div style="display:flex;justify-content:space-between;"><span style="color:#64748b;">Paid Date / วันที่ชำระ:</span><strong style="color:#10b981;">${formatDate(inv.paid_date)}</strong></div>` : ''}
                    </div>
                </div>

                <table class="invoice-table" style="width:100%;border-collapse:collapse;margin-bottom:28px;">
                    <thead style="background:#f1f5f9;border-bottom:2px solid #cbd5e1;">
                        <tr>
                            <th style="text-align:left;padding:12px 14px;color:#334155;font-size:0.82rem;text-transform:uppercase;">Description / รายการ</th>
                            <th style="text-align:right;padding:12px 14px;color:#334155;font-size:0.82rem;text-transform:uppercase;width:15%;">Hours / Qty</th>
                            <th style="text-align:right;padding:12px 14px;color:#334155;font-size:0.82rem;text-transform:uppercase;width:18%;">Rate / หน่วยละ</th>
                            <th style="text-align:right;padding:12px 14px;color:#334155;font-size:0.82rem;text-transform:uppercase;width:20%;">Amount / จำนวนเงิน</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${items.map((it, idx) => `
                            <tr style="border-bottom:1px solid #e2e8f0;background:${idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
                                <td style="padding:14px;font-weight:500;color:#0f172a;font-size:0.92rem;">${esc(it.desc)}</td>
                                <td style="text-align:right;padding:14px;color:#475569;font-size:0.9rem;">${it.hours}</td>
                                <td style="text-align:right;padding:14px;color:#475569;font-size:0.9rem;">$${Number(it.rate).toFixed(2)}</td>
                                <td style="text-align:right;padding:14px;font-weight:700;color:#0f172a;font-size:0.95rem;">$${Number(it.amount).toFixed(2)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>

                <div style="display:grid;grid-template-columns:1.2fr 0.8fr;gap:30px;align-items:start;">
                    <div>
                        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;font-size:0.85rem;color:#334155;">
                            <div style="font-weight:700;color:#0f172a;margin-bottom:6px;font-size:0.88rem;">💳 Payment Terms & Bank Transfer Info:</div>
                            ${inv.notes ? `<div style="margin-bottom:8px;color:#0f172a;">${esc(inv.notes)}</div>` : ''}
                            <div><strong>Bank Name:</strong> Kasikorn Bank (Public Co., Ltd.)</div>
                            <div><strong>Account Name:</strong> AG Projects Co., Ltd.</div>
                            <div><strong>Account Number:</strong> 777-2-34567-8 (Savings / Silom Branch)</div>
                            <div><strong>PromptPay / Tax ID:</strong> 0105550001234</div>
                        </div>
                    </div>
                    <div style="background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:18px;display:flex;flex-direction:column;gap:10px;">
                        <div style="display:flex;justify-content:space-between;font-size:0.9rem;color:#475569;">
                            <span>Subtotal / ยอดรวม:</span>
                            <strong style="color:#0f172a;">$${subtotal.toFixed(2)}</strong>
                        </div>
                        ${discount > 0 ? `
                        <div style="display:flex;justify-content:space-between;font-size:0.9rem;color:#ef4444;">
                            <span>Less Discount / ส่วนลด:</span>
                            <strong>-$${discount.toFixed(2)}</strong>
                        </div>` : ''}
                        <div style="display:flex;justify-content:space-between;font-size:0.9rem;color:#475569;">
                            <span>VAT ${taxRate}% / ภาษีมูลค่าเพิ่ม:</span>
                            <strong style="color:#0f172a;">$${taxAmount.toFixed(2)}</strong>
                        </div>
                        <div style="display:flex;justify-content:space-between;font-size:1.2rem;font-weight:800;color:#2684FF;border-top:2px solid #cbd5e1;padding-top:12px;margin-top:2px;">
                            <span>Grand Total / ยอดสุทธิ:</span>
                            <span>$${totalAmount.toFixed(2)}</span>
                        </div>
                    </div>
                </div>

                <!-- Signature Blocks -->
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:50px;padding-top:30px;border-top:1px dashed #cbd5e1;text-align:center;font-size:0.85rem;color:#475569;">
                    <div>
                        <div style="height:50px;border-bottom:1px solid #94a3b8;width:70%;margin:0 auto 10px;"></div>
                        <div>Authorized Signature / ผู้มีอำนาจลงนาม</div>
                        <div style="font-size:0.75rem;color:#64748b;margin-top:4px;">Date: _____ / _____ / ________</div>
                    </div>
                    <div>
                        <div style="height:50px;border-bottom:1px solid #94a3b8;width:70%;margin:0 auto 10px;"></div>
                        <div>Customer Acceptance / ผู้รับบริการ</div>
                        <div style="font-size:0.75rem;color:#64748b;margin-top:4px;">Date: _____ / _____ / ________</div>
                    </div>
                </div>
            </div>
        `;
        openModal('invoice-preview-modal');
    } catch (e) { showToast('Failed to load invoice'); }
}

function editIssuerAddressFromPreview() {
    const curName = localStorage.getItem('ag_issuer_name') || 'AG PROJECTS CO., LTD.';
    const curAddr = localStorage.getItem('ag_issuer_addr') || '123 Tech Tower, Silom Road, Bangrak, Bangkok 10500\nTax ID: 0105550001234 (Head Office) | Tel: 02-111-2222';
    const newName = prompt(currentLang === 'th' ? 'ชื่อบริษัทผู้ออกบิล (Company Name):' : 'Issuer Company Name:', curName);
    if (newName === null) return;
    const newAddr = prompt(currentLang === 'th' ? 'ที่อยู่บริษัทและเลขประจำตัวผู้เสียภาษี (Address & Tax ID):' : 'Issuer Address & Tax ID:', curAddr);
    if (newAddr === null) return;
    localStorage.setItem('ag_issuer_name', newName.trim() || 'AG PROJECTS CO., LTD.');
    localStorage.setItem('ag_issuer_addr', newAddr.trim() || curAddr);
    const sel = $('#preview-inv-status');
    if (sel && sel.dataset.id) {
        showInvoicePreview(sel.dataset.id);
    }
    showToast(currentLang === 'th' ? 'บันทึกที่อยู่บริษัทผู้ออกบิลเรียบร้อย' : 'Issuer company address updated');
}

// ---- Task Detail Modal ----
async function openTaskDetail(taskId) {
    currentTaskDetailId = taskId;
    try {
        const tasks = await fetch(API(`/api/tasks`)).then(r => r.json());
        const task = tasks.find(t => t.id === taskId);
        if (!task) { showToast(currentLang === 'th' ? 'ไม่พบข้อมูลงาน' : 'Task not found'); return; }

        $('#td-task-name').textContent = task.name;
        const priEl = $('#td-task-priority');
        priEl.className = `badge badge-${task.priority}`;
        priEl.textContent = task.priority === 'high' ? t('pri_high') : task.priority === 'medium' ? t('pri_med') : t('pri_low');
        $('#td-task-due').textContent = formatDate(task.due);
        $('#td-task-recur').textContent = task.recurring_pattern && task.recurring_pattern !== 'none' ? task.recurring_pattern : 'None';

        const assignEl = $('#td-task-assignee');
        if (assignEl) {
            const avatarColor = task.assignee_avatar || '#8b949e';
            const name = task.assignee_name || t('unassigned') || 'Unassigned';
            assignEl.innerHTML = `<span style="display:inline-flex;align-items:center;gap:6px;"><span style="width:20px;height:20px;border-radius:50%;background:${avatarColor};color:#fff;font-size:0.7rem;font-weight:700;display:inline-flex;align-items:center;justify-content:center;">${esc(getInitials(name))}</span> <span>${esc(name)}</span></span>`;
        }

        const comments = await fetch(API(`/api/tasks/${taskId}/comments`)).then(r => r.json());
        const commentsEl = $('#td-comments-list');
        commentsEl.innerHTML = comments.map(c => `
            <div class="comment-card">
                <div class="comment-card-header" style="display:flex;justify-content:space-between;margin-bottom:4px;font-size:0.75rem;"><span class="comment-author">${esc(c.user_name)}</span><span class="comment-time" style="color:var(--text-muted);">${formatDate(c.created_at)}</span></div>
                <div class="comment-content" style="color:var(--text-primary);font-size:0.88rem;line-height:1.4;">${esc(c.content)}</div>
            </div>
        `).join('') || emptyState('💬', t('no_comments'));
        commentsEl.scrollTop = commentsEl.scrollHeight;

        openModal('task-detail-modal');
    } catch (e) { showToast('Failed to load task details'); }
}

async function deleteCurrentTask() {
    if (!currentTaskDetailId) return;
    if (!(await confirmDialog(currentLang === 'th' ? 'คุณแน่ใจหรือไม่ว่าต้องการลบงานนี้?' : 'Are you sure you want to delete this task?'))) return;
    try {
        const res = await fetch(API(`/api/tasks/${currentTaskDetailId}`), { method: 'DELETE' });
        if (res.ok) {
            showToast(currentLang === 'th' ? 'ลบงานเรียบร้อยแล้ว' : 'Task deleted successfully');
            closeModal('task-detail-modal');
            refreshCurrentView();
            loadDashboard();
        } else {
            showToast('Failed to delete task');
        }
    } catch (e) {
        showToast('Error deleting task');
    }
}

function openAddTaskModal() {
    $('#task-edit-id').value = '';
    $('#task-name').value = '';
    $('#task-priority').value = 'medium';
    $('#task-start').value = '';
    $('#task-due').value = '';
    if ($('#task-milestone')) $('#task-milestone').value = '';
    if ($('#task-recurring')) $('#task-recurring').value = 'none';
    if ($('#task-assignee')) $('#task-assignee').value = '';
    $('#task-modal-title').textContent = t('modal_add_task');
    $('#task-modal-submit-btn').textContent = t('btn_add_task');
    openModal('new-task-modal');
}

async function openEditTaskModal(taskId) {
    if (!taskId) return;
    try {
        const tasks = await fetch(API(`/api/tasks`)).then(r => r.json());
        const task = tasks.find(x => x.id === taskId);
        if (!task) { showToast(currentLang === 'th' ? 'ไม่พบข้อมูลงาน' : 'Task not found'); return; }

        closeModal('task-detail-modal');

        $('#task-edit-id').value = task.id;
        $('#task-name').value = task.name || '';
        $('#task-priority').value = task.priority || 'medium';
        $('#task-start').value = task.start_date || '';
        $('#task-due').value = task.due || '';
        if ($('#task-milestone')) {
            const milestones = await fetch(API(`/api/milestones?project_id=${task.project_id}`)).then(r => r.json());
            $('#task-milestone').innerHTML = '<option value="">None</option>' + milestones.map(m => `<option value="${m.id}">${esc(m.name)}</option>`).join('');
            $('#task-milestone').value = task.milestone_id || '';
        }
        if ($('#task-recurring')) $('#task-recurring').value = task.recurring_pattern || 'none';
        if ($('#task-assignee')) $('#task-assignee').value = task.assignee_id || '';

        $('#task-modal-title').textContent = t('modal_edit_task');
        $('#task-modal-submit-btn').textContent = t('btn_save_changes');
        openModal('new-task-modal');
    } catch (e) {
        showToast('Failed to load task for editing');
    }
}

// ---- Settings ----
function loadSettings() {
    if (!currentUser) return;
    $('#settings-fullname').value = currentUser.name;
    $('#settings-email').value = currentUser.email;
    const initials = getInitials(currentUser.name);
    const preview = $('#settings-avatar-preview');
    $('#settings-avatar-url-input').value = currentUser.avatar_url || '';
    if (currentUser.avatar_url && currentUser.avatar_url.trim() !== '') {
        preview.innerHTML = `<img src="${esc(currentUser.avatar_url)}" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">`;
        preview.style.backgroundColor = 'transparent';
        if ($('#btn-remove-photo')) $('#btn-remove-photo').style.display = 'inline-flex';
    } else {
        preview.innerHTML = '';
        preview.textContent = initials;
        preview.style.backgroundColor = currentUser.color || '#2684FF';
        if ($('#btn-remove-photo')) $('#btn-remove-photo').style.display = 'none';
    }
    $$('.color-circle').forEach(c => {
        c.classList.toggle('selected', c.dataset.color === (currentUser.color || '#2684FF'));
    });
}

function handleProfileImageSelect(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size > 3 * 1024 * 1024) {
        showToast(currentLang === 'th' ? '⛔ ขนาดไฟล์รูปภาพต้องไม่เกิน 3 MB' : '⛔ Image size must not exceed 3 MB');
        return;
    }
    const reader = new FileReader();
    reader.onload = function(evt) {
        const dataUrl = evt.target.result;
        $('#settings-avatar-url-input').value = dataUrl;
        const preview = $('#settings-avatar-preview');
        preview.innerHTML = `<img src="${dataUrl}" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">`;
        preview.style.backgroundColor = 'transparent';
        if ($('#btn-remove-photo')) $('#btn-remove-photo').style.display = 'inline-flex';
        showToast(currentLang === 'th' ? '✅ เลือกรูปแล้ว กด Save Changes ด้านล่างเพื่อบันทึก' : '✅ Photo selected! Click Save Changes below.');
    };
    reader.readAsDataURL(file);
}

function removeProfilePhoto() {
    $('#settings-avatar-url-input').value = '';
    const preview = $('#settings-avatar-preview');
    preview.innerHTML = '';
    preview.textContent = getInitials(currentUser.name);
    const selectedColor = $('.color-circle.selected');
    preview.style.backgroundColor = selectedColor ? selectedColor.dataset.color : (currentUser.color || '#2684FF');
    if ($('#btn-remove-photo')) $('#btn-remove-photo').style.display = 'none';
    showToast(currentLang === 'th' ? '🗑️ นำรูปออกแล้ว กด Save Changes เพื่อยืนยัน' : '🗑️ Photo removed. Click Save Changes.');
}

$$('.color-circle').forEach(circle => {
    circle.addEventListener('click', () => {
        $$('.color-circle').forEach(c => c.classList.remove('selected'));
        circle.classList.add('selected');
        const hasPhoto = $('#settings-avatar-url-input').value && $('#settings-avatar-url-input').value.trim() !== '';
        if (!hasPhoto) {
            $('#settings-avatar-preview').style.backgroundColor = circle.dataset.color;
        }
    });
});

$('#settings-profile-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = $('#settings-fullname').value.trim();
    const selectedColor = $('.color-circle.selected');
    const color = selectedColor ? selectedColor.dataset.color : '#2684FF';
    const avatarUrl = $('#settings-avatar-url-input').value || '';
    try {
        await fetch(API('/api/profile/update'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, color, avatar_url: avatarUrl, email: currentUser.email })
        });
        currentUser.name = name;
        currentUser.color = color;
        currentUser.avatar_url = avatarUrl;
        localStorage.setItem('ag_current_user', JSON.stringify(currentUser));
        updateUserUI();
        showToast(currentLang === 'th' ? 'บันทึกโปรไฟล์และรูปถ่ายเรียบร้อยแล้ว!' : 'Profile and photo updated successfully!');
    } catch (e) { showToast('Failed to save profile'); }
});

// ---- Notifications ----
async function updateNotificationBadge() {
    try {
        const { count } = await fetch(API('/api/notifications/unread-count')).then(r => r.json());
        const badge = $('#notif-badge');
        if (!badge) return;
        badge.textContent = count > 99 ? '99+' : count;
        badge.style.display = count > 0 ? 'flex' : 'none';
    } catch (e) {}
}

$('#notif-bell-btn').addEventListener('click', async () => {
    const dropdown = $('#notif-dropdown');
    const isActive = dropdown.classList.contains('active');
    dropdown.classList.toggle('active');
    if (!isActive) {
        try {
            const notifs = await fetch(API('/api/notifications')).then(r => r.json());
            const container = $('#notif-list-container');
            container.innerHTML = notifs.map(n => `
                <div class="notif-item ${n.is_read ? '' : 'unread'}">
                    <div style="font-weight:500;color:var(--text-primary);">${esc(n.message)}</div>
                    <div style="font-size:0.72rem;color:var(--text-muted);margin-top:4px;">🕒 ${formatDate(n.created_at)}</div>
                </div>
            `).join('') || emptyState('🔔', currentLang === 'th' ? 'ยังไม่มีการแจ้งเตือน' : 'No notifications');

            updateNotificationBadge();
        } catch (e) {}
    }
});

document.addEventListener('click', (e) => {
    if (!e.target.closest('#notif-bell-btn') && !e.target.closest('#notif-dropdown')) {
        $('#notif-dropdown').classList.remove('active');
    }
});

async function markAllNotificationsRead() {
    try {
        await fetch(API('/api/notifications/read'), { method: 'PUT' });
        $$('.notif-item').forEach(n => n.classList.remove('unread'));
        $('#notif-badge').style.display = 'none';
        showToast(currentLang === 'th' ? 'อ่านการแจ้งเตือนทั้งหมดแล้ว' : 'All notifications marked as read');
    } catch (e) {}
}

setInterval(() => {
    if (!currentUser) return;
    updateNotificationBadge();
}, 15000);

// ---- Modal Utilities ----
function openModal(id) { $(`#${id}`).classList.add('active'); }
function closeModal(id) { $(`#${id}`).classList.remove('active'); }
function closeModalOnOverlay(e, id) { if (e.target === e.currentTarget) closeModal(id); }

// ---- Styled Confirm Dialog (replaces native confirm()) ----
let _confirmResolver = null;
function confirmDialog(message, title) {
    return new Promise((resolve) => {
        _confirmResolver = resolve;
        $('#confirm-modal-message').textContent = message;
        $('#confirm-modal-title').textContent = title || (currentLang === 'th' ? '⚠️ กรุณายืนยัน' : '⚠️ Please Confirm');
        openModal('confirm-modal');
    });
}
function resolveConfirmDialog(result) {
    closeModal('confirm-modal');
    if (_confirmResolver) {
        _confirmResolver(result);
        _confirmResolver = null;
    }
}

// ---- Form Submissions ----
$('#new-workspace-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = $('#workspace-name').value.trim();
    try {
        await fetch(API('/api/workspaces'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name })
        });
        closeModal('new-workspace-modal');
        $('#workspace-name').value = '';
        loadWorkspaces();
        showToast(currentLang === 'th' ? `สร้างพื้นที่ทำงาน "${name}" สำเร็จ!` : `Workspace "${name}" created!`);
    } catch (e) { showToast('Failed to create workspace'); }
});

$('#new-project-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = $('#project-name').value.trim();
    const desc = $('#project-desc').value.trim();
    try {
        await fetch(API('/api/projects'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ workspace_id: currentWorkspaceId, name, desc })
        });
        closeModal('new-project-modal');
        $('#project-name').value = '';
        $('#project-desc').value = '';
        loadProjects();
        showToast(currentLang === 'th' ? `สร้างโครงการ "${name}" สำเร็จ!` : `Project "${name}" created!`);
    } catch (e) { showToast('Failed to create project'); }
});

$('#new-task-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const editId = $('#task-edit-id').value;
    const name = $('#task-name').value.trim();
    const priority = $('#task-priority').value;
    const start_date = $('#task-start').value;
    const due = $('#task-due').value;
    const milestone_id = $('#task-milestone').value;
    const recurring_pattern = $('#task-recurring').value;
    const assignee_id = $('#task-assignee') ? $('#task-assignee').value : null;
    const selOpt = $('#task-assignee')?.selectedOptions?.[0];
    const assignee_name = selOpt?.dataset?.name || null;
    const assignee_avatar = selOpt?.dataset?.avatar || null;
    if (start_date && due && start_date > due) {
        showToast(currentLang === 'th' ? 'วันเริ่มต้นต้องไม่เกินวันครบกำหนด' : 'Start date cannot be after the due date');
        return;
    }
    const payload = { name, priority, due, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar, start_date };
    try {
        const res = editId
            ? await fetch(API(`/api/tasks/${editId}`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
            : await fetch(API('/api/tasks'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ...payload, project_id: currentProjectId })
            });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            showToast(err.error || (editId ? 'Failed to save task' : 'Failed to add task'));
            return;
        }
        closeModal('new-task-modal');
        $('#task-edit-id').value = '';
        $('#task-name').value = '';
        $('#task-start').value = '';
        refreshCurrentView();
        showToast(editId
            ? (currentLang === 'th' ? `บันทึกงาน "${name}" แล้ว!` : `Task "${name}" saved!`)
            : (currentLang === 'th' ? `เพิ่มงาน "${name}" แล้ว!` : `Task "${name}" added!`));
    } catch (e) { showToast(editId ? 'Failed to save task' : 'Failed to add task'); }
});

function openAddMilestoneModal() {
    $('#new-milestone-form').reset();
    $('#milestone-edit-id').value = '';
    $('#milestone-modal-title').textContent = t('modal_add_ms');
    $('#milestone-modal-submit-btn').textContent = t('btn_add');
    openModal('new-milestone-modal');
}

async function openEditMilestoneModal(milestoneId) {
    try {
        const milestones = await fetch(API(`/api/milestones?project_id=${currentProjectId}`)).then(r => r.json());
        const m = milestones.find(x => x.id === milestoneId);
        if (!m) { showToast('Milestone not found'); return; }

        $('#milestone-edit-id').value = m.id;
        $('#milestone-name').value = m.name || '';
        $('#milestone-date').value = m.date || '';
        $('#milestone-modal-title').textContent = t('modal_edit_ms') || (currentLang === 'th' ? 'แก้ไขหมุดหมาย' : 'Edit Milestone');
        $('#milestone-modal-submit-btn').textContent = t('btn_save_changes');
        openModal('new-milestone-modal');
    } catch (e) { showToast('Failed to load milestone for editing'); }
}

$('#new-milestone-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const editId = $('#milestone-edit-id').value;
    const name = $('#milestone-name').value.trim();
    const date = $('#milestone-date').value;
    try {
        const res = editId
            ? await fetch(API(`/api/milestones/${editId}`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, date })
            })
            : await fetch(API('/api/milestones'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ project_id: currentProjectId, name, date })
            });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.error) {
            showToast(data.error || (editId ? 'Failed to save milestone' : 'Failed to add milestone'));
            return;
        }
        closeModal('new-milestone-modal');
        $('#milestone-edit-id').value = '';
        $('#milestone-name').value = '';
        loadMilestones();
        showToast(editId
            ? (currentLang === 'th' ? `บันทึกหมุดหมาย "${name}" แล้ว!` : `Milestone "${name}" saved!`)
            : (currentLang === 'th' ? `เพิ่มหมุดหมาย "${name}" แล้ว!` : `Milestone "${name}" added!`));
    } catch (e) { showToast(editId ? 'Failed to save milestone' : 'Failed to add milestone'); }
});

function openAddTimesheetModal() {
    $('#new-timesheet-form').reset();
    $('#timesheet-edit-id').value = '';
    $('#timesheet-modal-title').textContent = t('modal_log_time');
    $('#timesheet-modal-submit-btn').textContent = t('btn_log_entry');
    openModal('new-timesheet-modal');
}

async function openEditTimesheetModal(timesheetId) {
    try {
        const timesheets = await fetch(API(`/api/timesheets?project_id=${currentProjectId}`)).then(r => r.json());
        const ts = timesheets.find(x => x.id === timesheetId);
        if (!ts) { showToast('Time entry not found'); return; }
        if (ts.billed) {
            showToast(currentLang === 'th' ? 'แก้ไขไม่ได้ เนื่องจากออกใบแจ้งหนี้ไปแล้ว' : 'Cannot edit — this entry has already been billed on an invoice');
            return;
        }

        $('#timesheet-edit-id').value = ts.id;
        $('#timesheet-task').value = ts.task_id || '';
        $('#timesheet-hours').value = ts.hours || '';
        $('#timesheet-date').value = ts.date || '';
        $('#timesheet-notes').value = ts.notes || '';
        $('#timesheet-modal-title').textContent = t('modal_edit_time') || (currentLang === 'th' ? 'แก้ไขเวลาทำงาน' : 'Edit Time Entry');
        $('#timesheet-modal-submit-btn').textContent = t('btn_save_changes');
        openModal('new-timesheet-modal');
    } catch (e) { showToast('Failed to load time entry for editing'); }
}

$('#new-timesheet-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const editId = $('#timesheet-edit-id').value;
    const task_id = $('#timesheet-task').value;
    const hours = parseFloat($('#timesheet-hours').value);
    const date = $('#timesheet-date').value;
    const notes = $('#timesheet-notes').value.trim();
    try {
        const res = editId
            ? await fetch(API(`/api/timesheets/${editId}`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ task_id, hours, date, notes })
            })
            : await fetch(API('/api/timesheets'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ project_id: currentProjectId, task_id, hours, date, notes, user: currentUser.name })
            });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.error) {
            showToast(data.error || (editId ? 'Failed to save time entry' : 'Failed to log time'));
            return;
        }
        closeModal('new-timesheet-modal');
        $('#timesheet-edit-id').value = '';
        loadTimesheets();
        showToast(editId
            ? (currentLang === 'th' ? 'บันทึกการแก้ไขเรียบร้อย!' : 'Time entry updated!')
            : (currentLang === 'th' ? 'บันทึกเวลาทำงานเรียบร้อย!' : 'Time entry logged!'));
    } catch (e) { showToast(editId ? 'Failed to save time entry' : 'Failed to log time'); }
});

$('#new-invoice-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const editId = $('#invoice-edit-id').value;
    const issuerName = $('#invoice-issuer-name')?.value.trim();
    const issuerAddr = $('#invoice-issuer-addr')?.value.trim();
    if (issuerName) localStorage.setItem('ag_issuer_name', issuerName);
    if (issuerAddr) localStorage.setItem('ag_issuer_addr', issuerAddr);

    const client = $('#invoice-client').value.trim();
    const client_address = $('#invoice-client-addr')?.value.trim() || '';
    const date = $('#invoice-date')?.value || new Date().toISOString().split('T')[0];
    const due_date = $('#invoice-due-date')?.value || '';
    const discount = parseFloat($('#invoice-discount')?.value) || 0;
    const tax_rate = parseFloat($('#invoice-tax-rate')?.value) || 0;
    const notes = $('#invoice-notes').value.trim();

    try {
        let res;
        if (editId) {
            res = await fetch(API(`/api/invoices/${editId}`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ client, client_address, date, due_date, discount, tax_rate, notes, items: currentInvoiceLineItems })
            });
        } else {
            const checkedBoxes = $$('.unbilled-ts-cb:checked');
            const timesheet_ids = Array.from(checkedBoxes).map(cb => cb.value);
            res = await fetch(API('/api/invoices'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    project_id: currentProjectId,
                    client,
                    client_address,
                    date,
                    due_date,
                    discount,
                    tax_rate,
                    notes,
                    items: currentInvoiceLineItems,
                    timesheet_ids
                })
            });
        }
        const data = await res.json();
        if (!res.ok || data.error) {
            showToast(data.error || data.message || (editId ? 'Failed to save invoice' : 'Failed to create invoice'));
        } else {
            closeModal('new-invoice-modal');
            $('#invoice-edit-id').value = '';
            loadBilling();
            showToast(editId
                ? (currentLang === 'th' ? 'บันทึกใบแจ้งหนี้เรียบร้อย!' : 'Invoice saved successfully!')
                : (currentLang === 'th' ? 'ออกใบแจ้งหนี้เรียบร้อย!' : 'Enterprise Invoice generated successfully!'));
        }
    } catch (e) { showToast(editId ? 'Failed to save invoice' : 'Failed to create invoice'); }
});

$('#td-comment-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const content = $('#td-comment-input').value.trim();
    if (!content || !currentTaskDetailId) return;
    try {
        await fetch(API(`/api/tasks/${currentTaskDetailId}/comments`), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ user_name: currentUser.name, content })
        });
        $('#td-comment-input').value = '';
        openTaskDetail(currentTaskDetailId);
        showToast(currentLang === 'th' ? 'โพสต์ความคิดเห็นแล้ว' : 'Comment posted');
    } catch (e) { showToast('Failed to post comment'); }
});

async function exportTasksToCSV() {
    if (!currentProjectId) return;
    try {
        const res = await fetch(API(`/api/export/tasks?project_id=${currentProjectId}`));
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'tasks_export.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast(currentLang === 'th' ? 'ดาวน์โหลดไฟล์ CSV เรียบร้อย' : 'Tasks exported to CSV');
    } catch (e) { showToast('Export failed'); }
}

function openSetRateModal() {
    if (!currentUser) return;
    if (currentUser.role !== 'admin' && currentUser.role !== 'manager') {
        showToast(currentLang === 'th' ? '⛔ เฉพาะ Admin และหัวหน้าทีม (Manager) เท่านั้นที่สามารถเปลี่ยนเรทราคาได้' : '⛔ Permission denied: Only Admin and Manager can change standard billing rates.');
        return;
    }
    const rate = getCurrentWorkspaceRate();
    if ($('#workspace-hourly-rate')) $('#workspace-hourly-rate').value = rate;
    openModal('set-rate-modal');
}

if ($('#set-rate-form')) {
    $('#set-rate-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!currentWorkspaceId) return;
        const newRate = parseFloat($('#workspace-hourly-rate').value) || 50;
        try {
            const res = await fetch(API(`/api/workspaces/${currentWorkspaceId}/rate`), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ hourly_rate: newRate })
            });
            const data = await res.json();
            if (data.success) {
                const ws = allWorkspacesList.find(w => w.id === currentWorkspaceId);
                if (ws) ws.hourly_rate = newRate;
                closeModal('set-rate-modal');
                showToast(currentLang === 'th' ? `⚙️ อัปเดตเรทค่าบริการเป็น $${newRate}/ชม. เรียบร้อยแล้ว` : `⚙️ Standard rate updated to $${newRate}/hr`);
                loadBilling();
            }
        } catch (err) {
            showToast('Failed to update hourly rate');
        }
    });
}

// ---- Auto-Restore Session on Page Reload ----
window.addEventListener('DOMContentLoaded', () => {
    try {
        const savedUser = localStorage.getItem('ag_current_user');
        const savedToken = localStorage.getItem('ag_session_token');
        if (savedUser && savedToken) {
            currentUser = JSON.parse(savedUser);
            onLoginSuccess(true);
        }
    } catch (e) {
        console.warn('Failed to restore session:', e);
    }
});
