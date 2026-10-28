/* ==========================================
   4. DASHBOARD MODULE (SharePoint Integration & Tables)
   ========================================== */
const DashboardModule = (function() {

    // Fetch and render sent/received requests from SharePoint list
    async function loadDashboardData() {
        const sentBody = document.getElementById('sentRequestsBody');
        const receivedBody = document.getElementById('receivedRequestsBody');
        const recentContainer = document.getElementById('recentActivityContainer');
        
        if (!sentBody || !receivedBody) return;

        sentBody.innerHTML = '<tr><td colspan="5">Loading from SharePoint...</td></tr>'; 
        receivedBody.innerHTML = '<tr><td colspan="5">Loading from SharePoint...</td></tr>'; 
        if(recentContainer) recentContainer.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">Loading recent activity...</p>';

        try {
            const account = AuthModule.msalInstance.getAllAccounts()[0];
            if (!account) {
                 if(recentContainer) recentContainer.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">Sign in to view recent activity.</p>';
                 return;
            }

            const tokenResponse = await AuthModule.msalInstance.acquireTokenSilent({
                scopes: ["Sites.Read.All"],
                account: account
            });

            const { siteHostname, sitePath, listId } = Config.sharePointContext;
            const graphUrl = `${Config.graphApiUrl}/sites/${siteHostname}:${sitePath}:/lists/${listId}/items?$expand=fields`;
            
            const response = await fetch(graphUrl, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });

            if (!response.ok) throw new Error("Failed to fetch data from SharePoint");

            const data = await response.json();
            const items = data.value;

            sentBody.innerHTML = '';
            receivedBody.innerHTML = '';
            
            realRequests = []; 

            if (items.length === 0) {
                 sentBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No sent requests found</td></tr>';
                 receivedBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No received requests found</td></tr>';
                 if(recentContainer) recentContainer.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">No recent activity found.</p>';
                 return;
            }

            // --- SORT ITEMS BY DATE (Newest to Oldest) ---
            items.sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));

            let sentCount = 0;
            let receivedCount = 0;
            
            // Array to hold recent activity for rendering later
            let recentActivityList = [];

            items.forEach((item) => {
                const fields = item.fields;
                const request = {
                    id: fields.OrderNumber || fields.AgreementID || fields.OData__x004f_rderNumber || "N/A", 
                    senderName: fields.SenderName || fields.Sender_x0020_Name || "N/A",
                    senderEmail: fields.SenderEmail || fields.Sender_x0020_Email || "N/A",
                    recipientName: fields.RecipientName || fields.Recipient_x0020_Name || fields.RecipientEmail || "N/A",
                    recipientEmail: fields.RecipientEmail || fields.Recipient_x0020_Email || "N/A",
                    fileName: fields.Title || "N/A", 
                    date: new Date(item.createdDateTime).toLocaleDateString(),
                    status: fields.RequestStatus || fields.Request_x0020_Status || "Pending" 
                };

                realRequests.push(request);
                const currentIndex = realRequests.length - 1;

                let statusClass = 'pending';
                if (request.status === 'Completed' || request.status === 'Signed') statusClass = 'signed';
                if (request.status === 'Declined') statusClass = 'declined';
                if (request.status === 'Cancelled') statusClass = 'cancelled'; // Added Cancelled logic
                
                let tr = document.createElement('tr');
                
                // Uniform column layout: [Target Name / Sender] [Document] [Date] [Status] [Details]
                if (request.senderEmail === currentUserEmail) {
                    sentCount++;
                    tr.innerHTML = `
                        <td>${request.recipientName}</td>
                        <td>${request.fileName}</td>
                        <td>${request.date}</td>
                        <td><span class="status ${statusClass}">${request.status}</span></td>
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})">View Details</button></td>
                    `;
                    sentBody.appendChild(tr);
                    
                    // Add to recent activity if relevant
                    recentActivityList.push({ type: 'Sent', data: request, index: currentIndex, statusClass: statusClass });
                } 
                else if (request.recipientEmail === currentUserEmail) {
                    receivedCount++;
                    tr.innerHTML = `
                        <td>${request.senderName}</td>
                        <td>${request.fileName}</td>
                        <td>${request.date}</td>
                        <td><span class="status ${statusClass}">${request.status}</span></td>
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})">View Details</button></td>
                    `;
                    receivedBody.appendChild(tr);
                    
                     // Add to recent activity if relevant
                    recentActivityList.push({ type: 'Received', data: request, index: currentIndex, statusClass: statusClass });
                }
            });

            if (sentCount === 0) sentBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No sent requests found</td></tr>';
            if (receivedCount === 0) receivedBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No received requests found</td></tr>';
            
            // Render Recent Activity (Top 3)
            renderRecentActivity(recentActivityList, recentContainer);
            
        } catch (error) {
            console.error("Dashboard Error:", error);
            sentBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
            receivedBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
            if(recentContainer) recentContainer.innerHTML = '<p style="color:red; font-size: 13px;">Error loading recent activity.</p>';
        }
    }
    
    // Function to render the top 3 recent activities in the New Request page
    function renderRecentActivity(activityList, container) {
        if (!container) return;
        
        if (activityList.length === 0) {
            container.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">No recent activity found.</p>';
            return;
        }

        // Take only the first 3 items (they are already sorted by date)
        const recentItems = activityList.slice(0, 3);
        
        let tableHTML = `
            <table class="requests-table" style="font-size: 13px; margin-bottom: 0;">
                <thead>
                    <tr>
                        <th style="padding: 10px;">Type</th>
                        <th style="padding: 10px;">Contact</th>
                        <th style="padding: 10px;">Document</th>
                        <th style="padding: 10px;">Status</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        recentItems.forEach(item => {
            const contactName = item.type === 'Sent' ? item.data.recipientName : item.data.senderName;
            const typeIcon = item.type === 'Sent' ? '<i class="fa-solid fa-paper-plane" style="color:var(--dv-teal);"></i> Sent' : '<i class="fa-solid fa-inbox" style="color:var(--dv-dark-blue);"></i> Received';
            
            tableHTML += `
                <tr>
                    <td style="padding: 10px;">${typeIcon}</td>
                    <td style="padding: 10px;">${contactName}</td>
                    <td style="padding: 10px;">${item.data.fileName}</td>
                    <td style="padding: 10px;"><span class="status ${item.statusClass}" style="font-size: 11px; padding: 3px 8px;">${item.data.status}</span></td>
                </tr>
            `;
        });
        
        tableHTML += `</tbody></table>`;
        container.innerHTML = tableHTML;
    }

    // Open details modal for a specific request
    function openDynamicModal(index) {
        const request = realRequests[index];
        if (!request) return;
        
        document.getElementById('modalSenderName').textContent = request.senderName;
        document.getElementById('modalSenderEmail').textContent = request.senderEmail;
        
        document.getElementById('modalRecipientName').textContent = request.recipientName;
        document.getElementById('modalRecipientEmail').textContent = request.recipientEmail;
        
        document.getElementById('modalFileName').textContent = request.fileName;
        document.getElementById('modalOrderId').textContent = request.id;
        document.getElementById('modalDate').textContent = request.date;
        
        const statusBadge = document.getElementById('modalStatusBadge');
        statusBadge.textContent = request.status;
        
        if (request.status === 'Signed' || request.status === 'Completed') {
            statusBadge.className = 'status signed';
        } else if (request.status === 'Declined') {
            statusBadge.className = 'status declined'; 
        } else if (request.status === 'Cancelled') {
            statusBadge.className = 'status cancelled'; 
        } else {
            statusBadge.className = 'status pending';
        }
        
        const btnReminder = document.getElementById('btnSendReminder');
        if (request.status === 'Pending' && request.senderEmail === currentUserEmail) {
            btnReminder.style.display = 'block'; 
        } else {
            btnReminder.style.display = 'none';  
        }

        const btnView = document.getElementById('btnViewDocument');
        btnView.onclick = function() {
            // جلب الهوست (مثال: dammamvalleyd.sharepoint.com)
            const siteHostname = Config.sharePointContext.siteHostname;
            // مسار الموقع (مثال: E-SignatureSystem) - نقوم بإزالة أي /sites/ قد تكون مضافة بالخطأ في الإعدادات
            let sitePath = Config.sharePointContext.sitePath.replace('/sites/', '');
            
            const archiveLibraryName = "SignedContractsArchive"; 
            
            // بناء الرابط النظيف
            const fileUrl = `https://${siteHostname}/sites/${sitePath}/${archiveLibraryName}/${encodeURIComponent(request.fileName)}?web=1`;
            
            window.open(fileUrl, '_blank');
        };
        
        document.getElementById('detailsModal').style.display = 'block';
    }

    function closeModal() {
        const modal = document.getElementById('detailsModal');
        if (modal) modal.style.display = 'none';
    }

    window.addEventListener('click', function(event) {
        const modal = document.getElementById('detailsModal');
        if (event.target === modal) {
            modal.style.display = "none";
        }
    });

    return {
        loadDashboardData,
        openDynamicModal,
        closeModal
    };
})();