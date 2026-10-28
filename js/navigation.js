/* ==========================================
   3. NAVIGATION MODULE (UI State Management)
   ========================================== */
const NavigationModule = (function() {

    function showSection(sectionId) {
        // Hide all main sections initially
        const newRequestSection = document.getElementById('newRequestSection');
        const dashboardSection = document.getElementById('dashboardSection');
        
        if (newRequestSection) newRequestSection.style.display = 'none';
        if (dashboardSection) dashboardSection.style.display = 'none';
        
        // Remove 'active' class from all sidebar links
        document.querySelectorAll('.nav-links li').forEach(li => {
            li.classList.remove('active');
        });
        
        // Show the targeted section and highlight the corresponding sidebar link
        if (sectionId === 'newRequest') {
            if (newRequestSection) newRequestSection.style.display = 'block';
            const navLink = document.querySelectorAll('.nav-links li')[0];
            if (navLink) navLink.classList.add('active');
            
        } else if (sectionId === 'dashboard') {
            if (dashboardSection) dashboardSection.style.display = 'block';
            const navLink = document.querySelectorAll('.nav-links li')[1];
            if (navLink) navLink.classList.add('active');
            
            // Refresh dashboard data every time the view is opened
            if (typeof DashboardModule !== 'undefined') {
                DashboardModule.loadDashboardData(); 
            }
        }
    }

    // Export the function to be accessible globally (e.g., from HTML onclick events)
    return {
        showSection
    };
})();