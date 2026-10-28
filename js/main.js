/* ==========================================
   8. MAIN ENTRY POINT (Application Bootstrap)
   ========================================== */
document.addEventListener('DOMContentLoaded', () => {
    console.log("Initializing Dammam Valley E-Signature Portal...");

    // 1. Bind modules to global window object for HTML inline event handlers
    window.AuthModule = AuthModule;
    window.NavigationModule = NavigationModule;
    window.DashboardModule = DashboardModule;
    window.RecipientsModule = RecipientsModule;
    window.FilesModule = FilesModule;
    window.SubmissionModule = SubmissionModule;

    // 2. Initialize Azure Authentication & User Session
    if (typeof AuthModule !== 'undefined' && AuthModule.authenticateUser) {
        AuthModule.authenticateUser();
    }

    // 3. Initialize Smart Recipients Search & Chips
    if (typeof RecipientsModule !== 'undefined' && RecipientsModule.init) {
        RecipientsModule.init();
    }

    // 4. Initialize Multiple Files Upload Management
    if (typeof FilesModule !== 'undefined' && FilesModule.init) {
        FilesModule.init();
    }
});