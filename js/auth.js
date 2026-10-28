/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL & Teams SSO)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام (مع دعم الدخول الصامت داخل التيمز)
    async function initializeAuth() {
        try {
            // التحقق مما إذا كان النظام يفتح داخل Microsoft Teams
            if (typeof microsoftTeams !== 'undefined') {
                try {
                    await microsoftTeams.app.initialize();
                    const context = await microsoftTeams.app.getContext();
                    const userEmail = context?.user?.userPrincipalName;
                    
                    if (userEmail) {
                        console.log("Teams Environment Detected. Attempting Silent SSO for:", userEmail);
                        
                        // محاولة الدخول الصامت باستخدام إيميل الموظف المأخوذ من التيمز مباشرة
                        try {
                            await msalInstance.ssoSilent({
                                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                                loginHint: userEmail
                            });
                        } catch (ssoError) {
                            console.warn("Teams Silent SSO required interaction:", ssoError);
                        }
                    }
                } catch (e) {
                    console.log("Running in a standard Web Browser.");
                }
            }

            await msalInstance.handleRedirectPromise();
        } catch (error) {
            // نتجاهل أي خطأ وهمي أثناء التهيئة
        } finally {
            checkUserSession();
        }
    }

    // 2. تحديث الواجهة بناءً على حالة الجلسة الحقيقية
    function checkUserSession() {
        const accounts = msalInstance.getAllAccounts();
        const authBtn = document.getElementById('authActionButton');
        const profileHeader = document.getElementById('profileMenuHeader');
        const greeting = document.getElementById('profileGreeting');
        const userNameEl = document.getElementById('headerUserName');
        const menuUserName = document.getElementById('menuUserName');
        const menuUserEmail = document.getElementById('menuUserEmail');

        if (accounts.length > 0) {
            const user = accounts[0];
            currentUserEmail = user.username;
            currentUserName = user.name;

            if (userNameEl) userNameEl.textContent = currentUserName;
            if (greeting) greeting.textContent = "Welcome,";
            if (menuUserName) menuUserName.textContent = currentUserName;
            if (menuUserEmail) menuUserEmail.textContent = currentUserEmail;
            if (profileHeader) profileHeader.style.display = 'flex'; 

            if (authBtn) {
                authBtn.className = 'btn-logout';
                authBtn.innerHTML = '<i class="fa-solid fa-arrow-right-from-bracket"></i> Sign Out';
            }

            const senderNameInput = document.getElementById('senderName');
            const senderEmailInput = document.getElementById('senderEmail');
            if (senderNameInput) senderNameInput.value = currentUserName;
            if (senderEmailInput) senderEmailInput.value = currentUserEmail;

            if (typeof DashboardModule !== 'undefined') {
                DashboardModule.loadDashboardData();
            }
        } else {
            if (userNameEl) userNameEl.textContent = "Sign In Required";
            if (greeting) greeting.textContent = "Hello,";
            if (menuUserName) menuUserName.textContent = '';
            if (menuUserEmail) menuUserEmail.textContent = '';
            if (profileHeader) profileHeader.style.display = 'none';

            if (authBtn) {
                authBtn.className = 'btn-signin';
                authBtn.innerHTML = '<i class="fa-solid fa-arrow-right-to-bracket"></i> Sign In';
            }
            
            const recentContainer = document.getElementById('recentActivityContainer');
            if(recentContainer) {
                 recentContainer.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">Sign in to view recent activity.</p>';
            }
        }
    }

    // 3. معالج النقر على الأزرار
    function handleAuthAction() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            logoutUser();
        } else {
            authenticateUser();
        }
    }

    // 4. تسجيل الدخول اليدوي (للمتصفح العادي)
    async function authenticateUser() {
        try {
            sessionStorage.removeItem("msal.interaction.status");
            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account"
            });
        } catch (error) {
            if (error.errorCode && error.errorCode !== "user_cancelled" && error.errorCode !== "interaction_in_progress") {
                console.warn("MSAL Popup Info:", error.errorCode);
            }
        } finally {
            checkUserSession(); 
        }
    }

    // 5. تسجيل الخروج السلس
    async function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) return;

        try {
            await msalInstance.logoutPopup({
                account: accounts[0]
            });
        } catch (error) {
            // نتجاهل الإنذارات الكاذبة
        } finally {
            sessionStorage.clear();
            checkUserSession();
        }
    }

    // 6. التحكم بالقائمة المنسدلة
    function toggleProfileMenu(event) {
        const menu = document.getElementById('profileMenu');
        if (menu) {
            menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
        }
    }

    document.addEventListener('click', function(event) {
        const profileBtn = document.getElementById('userProfileBtn');
        const profileMenu = document.getElementById('profileMenu');
        if (profileMenu && profileBtn && !profileBtn.contains(event.target)) {
            profileMenu.style.display = 'none';
        }
    });

    window.addEventListener('DOMContentLoaded', () => {
        initializeAuth();
    });

    return {
        authenticateUser,
        toggleProfileMenu,
        logoutUser,
        handleAuthAction,
        msalInstance
    };
})();
