/* ==========================================
   5. RECIPIENTS MODULE (Smart Search & Chips)
   ========================================== */
const RecipientsModule = (function() {

    const searchInput = document.getElementById('recipientSearch');
    const dropdown = document.getElementById('searchResultsDropdown');
    const chipsContainer = document.getElementById('recipientChipsContainer');
    const finalRecipients = document.getElementById('finalRecipients');

    let searchTimeout;

    // Search users via Microsoft Graph API
    async function searchUsers(searchTerm) {
        try {
            const account = AuthModule.msalInstance.getAllAccounts()[0];
            if (!account) return;

            const tokenResponse = await AuthModule.msalInstance.acquireTokenSilent({
                scopes: ["User.ReadBasic.All"], 
                account: account
            });
            
            let url = `${Config.graphApiUrl}/users?$top=20`;
            if (searchTerm.length > 0) {
                url += `&$search="displayName:${searchTerm}" OR "mail:${searchTerm}"&$count=true`;
            }

            const response = await fetch(url, { 
                headers: { 
                    "Authorization": `Bearer ${tokenResponse.accessToken}`,
                    "ConsistencyLevel": "eventual" 
                } 
            });
            
            if (!response.ok) throw new Error("Failed to fetch users");

            const data = await response.json();
            let finalUsers = data.value;
            
            if (searchTerm.length > 0) {
                const term = searchTerm.toLowerCase();
                finalUsers = finalUsers.filter(u => 
                    (u.displayName && u.displayName.toLowerCase().includes(term)) || 
                    (u.mail && u.mail.toLowerCase().includes(term))
                );
            }

            renderDropdown(finalUsers);
        } catch (error) {
            console.error("Search Error:", error);
        }
    }

    // Render dropdown suggestions
    function renderDropdown(users) {
        if (!dropdown) return;
        dropdown.innerHTML = '';
        
        if (!users || users.length === 0) {
            dropdown.style.display = 'none';
            return;
        }
        
        users.forEach(user => {
            const email = user.mail || user.userPrincipalName;
            if (!email) return;

            let div = document.createElement('div');
            div.className = 'dropdown-item';
            div.innerHTML = `<span class="dropdown-item-name">${user.displayName}</span><span class="dropdown-item-email">${email}</span>`;
            
            div.onclick = function() {
                addChip(user.displayName, email);
                dropdown.style.display = 'none';
                if (searchInput) searchInput.value = '';
            };
            dropdown.appendChild(div);
        });
        dropdown.style.display = 'block';
    }

    // Add recipient chip
    function addChip(name, email) {
        if (selectedRecipients.some(r => r.email === email)) return;
        
        selectedRecipients.push({ name: name, email: email });
        updateChipsUI();
    }

    // Remove recipient chip
    window.removeChip = function(email) {
        selectedRecipients = selectedRecipients.filter(r => r.email !== email);
        updateChipsUI();
    };

    // Update chips UI and hidden input
    function updateChipsUI() {
        if (!chipsContainer || !searchInput || !finalRecipients) return;

        document.querySelectorAll('.chip').forEach(el => el.remove());
        
        selectedRecipients.forEach(user => {
            let chip = document.createElement('div');
            chip.className = 'chip';
            chip.innerHTML = `<i class="fa-solid fa-user"></i> ${user.name} <span class="chip-close" onclick="removeChip('${user.email}')">&times;</span>`;
            chipsContainer.insertBefore(chip, searchInput);
        });
        
        finalRecipients.value = JSON.stringify(selectedRecipients);
    }

    // Initialize event listeners
    function init() {
        if (!searchInput) return;

        searchInput.addEventListener('focus', function() {
            searchUsers(this.value);
        });

        searchInput.addEventListener('input', function() {
            clearTimeout(searchTimeout);
            const val = this.value;
            searchTimeout = setTimeout(() => searchUsers(val), 200);
        });

        searchInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                const val = this.value.trim();
                const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
                
                if (emailRegex.test(val)) {
                    addChip(val, val);
                    this.value = '';
                    if (dropdown) dropdown.style.display = 'none';
                } else if (val.length > 0) {
                    alert("Please enter a valid email address (e.g., name@company.com).");
                }
            }
        });

        document.addEventListener('click', function(e) {
            if (dropdown && searchInput && e.target !== searchInput && e.target !== dropdown) {
                dropdown.style.display = 'none';
            }
        });
    }

    return {
        init,
        getSelectedRecipients: () => selectedRecipients,
        clearRecipients: () => {
            selectedRecipients = [];
            updateChipsUI();
        }
    };
})();