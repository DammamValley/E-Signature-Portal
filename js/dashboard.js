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
                    status: fields.RequestStatus || fields.Request_x0020_Status || "Pending",
                    // --- إضافة جلب حقل CurrentSignerEmail ---
                    currentSignerEmail: fields.CurrentSignerEmail || null
                };

                let statusClass = 'pending';
                if (request.status === 'Completed' || request.status === 'Signed') statusClass = 'signed';
                if (request.status === 'Declined') statusClass = 'declined';
                if (request.status === 'Cancelled') statusClass = 'cancelled'; 
                
                // --- إصلاح مشكلة الإرسال للنفس ---
                if (request.senderEmail === currentUserEmail) {
                    realRequests.push({...request, displayType: 'sent'});
                    const currentIndex = realRequests.length - 1;
                    sentCount++;
                    
                    let tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${request.recipientName}</td>
                        <td title="${request.fileName}">${request.fileName}</td>
                        <td>${request.date}</td>
                        <td><span class="status ${statusClass}">${request.status}</span></td>
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})"><i class="fa-solid fa-eye"></i> View</button></td>
                    `;
                    sentBody.appendChild(tr);
                    recentActivityList.push({ type: 'Sent', data: request, index: currentIndex, statusClass: statusClass });
                } 
                
                if (request.recipientEmail === currentUserEmail) {
                    realRequests.push({...request, displayType: 'received'});
                    const currentIndex = realRequests.length - 1;
                    receivedCount++;
                    
                    let tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${request.senderName}</td>
                        <td title="${request.fileName}">${request.fileName}</td>
                        <td>${request.date}</td>
                        <td><span class="status ${statusClass}">${request.status}</span></td>
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})"><i class="fa-solid fa-eye"></i> View</button></td>
                    `;
                    receivedBody.appendChild(tr);
                    recentActivityList.push({ type: 'Received', data: request, index: currentIndex, statusClass: statusClass });
                }
            });

            if (sentCount === 0) sentBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No sent requests found</td></tr>';
            if (receivedCount === 0) receivedBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No received requests found</td></tr>';
            
            renderRecentActivity(recentActivityList, recentContainer);
            
        } catch (error) {
            console.error("Dashboard Error:", error);
            sentBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
            receivedBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
            if(recentContainer) recentContainer.innerHTML = '<p style="color:red; font-size: 13px;">Error loading recent activity.</p>';
        }
    }
    
    function renderRecentActivity(activityList, container) {
        if (!container) return;
        
        if (activityList.length === 0) {
            container.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">No recent activity found.</p>';
            return;
        }

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
                    <td style="padding: 10px;" title="${item.data.fileName}">${item.data.fileName}</td>
                    <td style="padding: 10px;"><span class="status ${item.statusClass}" style="font-size: 11px; padding: 3px 8px;">${item.data.status}</span></td>
                </tr>
            `;
        });
        
        tableHTML += `</tbody></table>`;
        container.innerHTML = tableHTML;
    }

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
        
        // --- 1. برمجة زر التذكير ---
        const btnReminder = document.getElementById('btnSendReminder');
        
        if (request.status === 'Pending' && request.senderEmail === currentUserEmail) {
            btnReminder.style.display = 'block';
            
            btnReminder.onclick = async function() {
                const originalHtml = btnReminder.innerHTML;
                btnReminder.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Sending...';
                btnReminder.disabled = true;
                btnReminder.style.opacity = '0.7';

                try {
                    const payload = {
                        actionType: "remind",
                        agreementId: request.id,
                        userEmail: currentUserEmail
                    };

                    const response = await fetch(Config.masterFlowWebhookUrl, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });

                    if (response.ok) {
                        btnReminder.innerHTML = '<i class="fa-solid fa-check"></i> Reminder Sent';
                        btnReminder.style.backgroundColor = '#d4edda';
                        btnReminder.style.color = '#155724';
                        btnReminder.style.borderColor = '#c3e6cb';
                    } else {
                        throw new Error("Failed to send reminder");
                    }
                } catch (error) {
                    console.error("Reminder Error:", error);
                    btnReminder.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Error';
                    btnReminder.style.backgroundColor = '#fee2e2';
                    btnReminder.style.color = '#b91c1c';
                    alert("Failed to send reminder. Please check your connection or Webhook URL.");
                } finally {
                    setTimeout(() => {
                        btnReminder.innerHTML = originalHtml;
                        btnReminder.disabled = false;
                        btnReminder.style.opacity = '1';
                        btnReminder.style.backgroundColor = '';
                        btnReminder.style.color = '';
                        btnReminder.style.borderColor = '';
                    }, 3000);
                }
            };
        } else {
            btnReminder.style.display = 'none';  
        }

        // --- 2. برمجة زر العرض (View Document) ---
        const btnView = document.getElementById('btnViewDocument');
        btnView.onclick = function() {
            const siteHostname = Config.sharePointContext.siteHostname;
            let sitePath = Config.sharePointContext.sitePath.replace('/sites/', '');
            const archiveLibraryName = "SignedContractsArchive"; 
            const fileUrl = `https://${siteHostname}/sites/${sitePath}/${archiveLibraryName}/${encodeURIComponent(request.fileName)}?web=1`;
            window.open(fileUrl, '_blank');
        };

        // --- 3. إنشاء وبرمجة زر التوقيع الذكي (Sign Now) ---
        let btnSignNow = document.getElementById('btnSignNow');
        if (!btnSignNow) {
            btnSignNow = document.createElement('button');
            btnSignNow.id = 'btnSignNow';
            btnSignNow.className = 'btn-primary';
            btnSignNow.style.width = 'auto';
            document.querySelector('.modal-actions').prepend(btnSignNow);
        }

        // تنظيف خصائص الزر السابقة لضمان عدم تداخل الحالات بين الطلبات
        btnSignNow.disabled = false;
        btnSignNow.style.backgroundColor = '';
        btnSignNow.style.color = '';
        btnSignNow.style.cursor = 'pointer';
        btnSignNow.style.border = '';
        btnSignNow.onclick = null;

        if (request.status === 'Pending') {
            btnSignNow.style.display = 'block';

            // التحقق: هل الموظف الحالي هو المطلوب للتوقيع؟
            const isMyTurn = request.currentSignerEmail && request.currentSignerEmail.toLowerCase().includes(currentUserEmail.toLowerCase());

            if (isMyTurn) {
                // حالة 1: دور الموظف الحالي
                btnSignNow.innerHTML = '✍️ Sign Now';
                
                btnSignNow.onclick = async function() {
                    // 💡 خدعة تجاوز مانع النوافذ المنبثقة: نفتح نافذة فارغة فور الضغط
                    const newTab = window.open('about:blank', '_blank');

                    const originalHtml = btnSignNow.innerHTML;
                    btnSignNow.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Loading...';
                    btnSignNow.disabled = true;

                    try {
                        const payload = {
                            actionType: "sign",
                            agreementId: request.id,
                            userEmail: currentUserEmail
                        };

                        const response = await fetch(Config.masterFlowWebhookUrl, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify(payload)
                        });

                        if (response.ok) {
                            const responseData = await response.json();
                            if (responseData.url) {
                                // توجيه النافذة الفارغة إلى رابط أدوبي
                                newTab.location.href = responseData.url; 
                            } else {
                                newTab.close(); // إغلاق النافذة إذا فشل جلب الرابط
                                alert("Could not retrieve the signing URL.");
                            }
                        } else {
                            newTab.close();
                            throw new Error("Failed to get signing URL");
                        }
                    } catch (error) {
                        console.error("Sign Now Error:", error);
                        newTab.close();
                        alert("Error loading the signature page.");
                    } finally {
                        // إعادة الزر لشكله الطبيعي
                        btnSignNow.innerHTML = originalHtml;
                        btnSignNow.disabled = false;
                    }
                };
            } else {
                // حالة 2: قيد الانتظار، ولكن ليس دور الموظف الحالي
                btnSignNow.innerHTML = '⏳ Waiting for others';
                btnSignNow.disabled = true;
                btnSignNow.style.backgroundColor = '#f1f5f9'; // رمادي باهت
                btnSignNow.style.color = '#64748b'; // نص رمادي غامق
                btnSignNow.style.cursor = 'not-allowed';
                btnSignNow.style.border = '1px solid #cbd5e1';
            }
        } else {
            // حالة 3: إذا كان الطلب مكتمل أو ملغي، نخفي الزر تماماً
            btnSignNow.style.display = 'none';
        }
        
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
