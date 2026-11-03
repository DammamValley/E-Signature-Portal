/* ==========================================
   7. SUBMISSION MODULE (Modal, Reordering & API POST)
   ========================================== */
const SubmissionModule = (function() {

    // 1. Open options modal and review files & recipients
    function openOptionsModal() {
        // طبقة حماية ثانية تمنع إكمال الطلب إذا لم يكن هناك حساب مسجل
        const accounts = AuthModule.msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            alert("Security Error: Please sign in first to complete your request.");
            return;
        }

        const documentFiles = FilesModule.getFiles();
        const recipients = RecipientsModule.getSelectedRecipients();
        
        if (recipients.length === 0) {
            alert("Please select at least one recipient.");
            return;
        }
        if (documentFiles.length === 0) {
            alert("Please attach at least one document.");
            return;
        }

        // Render selected files
        const filesList = document.getElementById('summaryFilesList');
        if (filesList) {
            filesList.innerHTML = '';
            for (let i = 0; i < documentFiles.length; i++) {
                filesList.innerHTML += `<li><i class="fa-solid fa-file-lines" style="color:var(--dv-teal); margin-right:5px;"></i> ${documentFiles[i].name}</li>`;
            }
        }

        renderSummaryRecipients();

        // --- NEW LOGIC: Lock routing type if only one recipient is selected ---
        const routeGroup = document.getElementById('routeGroup');
        const routeIndividual = document.getElementById('routeIndividual');
        const singleRecipientNotice = document.getElementById('singleRecipientNotice');

        if (recipients.length === 1) {
            // Lock to Individual
            if (routeIndividual) routeIndividual.checked = true;
            if (routeGroup) routeGroup.disabled = true;
            // Show the warning message
            if (singleRecipientNotice) singleRecipientNotice.style.display = 'block';
        } else {
            // Re-enable and default to Group if multiple recipients
            if (routeGroup) {
                routeGroup.disabled = false;
                routeGroup.checked = true;
            }
            // Hide the warning message
            if (singleRecipientNotice) singleRecipientNotice.style.display = 'none';
        }
        // ----------------------------------------------------------------------

        // Show modal
        const modal = document.getElementById('sendOptionsModal');
        if (modal) modal.style.display = 'block';
    }

    // Close options modal
    function closeOptionsModal() {
        const modal = document.getElementById('sendOptionsModal');
        if (modal) modal.style.display = 'none';
    }

    // Render recipients list inside modal with up/down controls
    function renderSummaryRecipients() {
        const list = document.getElementById('summaryRecipientsList');
        if (!list) return;
        list.innerHTML = '';
        
        const recipients = RecipientsModule.getSelectedRecipients();

        recipients.forEach((rec, index) => {
            const li = document.createElement('li');
            
            const info = document.createElement('span');
            info.innerHTML = `<strong>${rec.name}</strong> <span style="font-size:11px; color:#888;">(${rec.email})</span>`;
            
            const controls = document.createElement('div');
            controls.className = 'reorder-controls';
            
            const upBtn = document.createElement('button');
            upBtn.innerHTML = '<i class="fa-solid fa-chevron-up"></i>';
            upBtn.disabled = index === 0;
            upBtn.onclick = () => moveRecipient(index, -1);
            
            const downBtn = document.createElement('button');
            downBtn.innerHTML = '<i class="fa-solid fa-chevron-down"></i>';
            downBtn.disabled = index === recipients.length - 1;
            downBtn.onclick = () => moveRecipient(index, 1);
            
            controls.appendChild(upBtn);
            controls.appendChild(downBtn);
            
            li.appendChild(info);
            li.appendChild(controls);
            list.appendChild(li);
        });
    }

    // Move recipient position up or down
    function moveRecipient(index, direction) {
        const recipients = RecipientsModule.getSelectedRecipients();
        const temp = recipients[index];
        recipients[index] = recipients[index + direction];
        recipients[index + direction] = temp;
        
        renderSummaryRecipients();
    }

    // Final submission to Power Automate webhook
    async function submitFinalRequest(signaturePlacement) {
        // 1. Format Payload Variables to strictly match Power Automate conditions
        const routingInput = document.querySelector('input[name="modalRouting"]:checked');
        let rawRouting = routingInput ? routingInput.value : 'sequential';
        const routingType = (rawRouting === 'parallel') ? 'Individual' : 'Sequential';
        const placementType = (signaturePlacement === 'manual') ? 'Manual' : 'Auto';

        const documentFiles = FilesModule.getFiles();
        const recipients = RecipientsModule.getSelectedRecipients();
        
        const senderName = document.getElementById('senderName').value || currentUserEmail;
        const senderEmail = document.getElementById('senderEmail').value || currentUserEmail;

        const modalContent = document.querySelector('#sendOptionsModal .modal-content');
        const originalContent = modalContent ? modalContent.innerHTML : '';
        
        // متغير لحماية شاشة النجاح من المسح التلقائي
        let isSuccessScreenActive = false; 

        if (modalContent) {
            // شاشة التحميل
            modalContent.innerHTML = '<div style="text-align:center; padding: 40px;"><h2><i class="fa-solid fa-spinner fa-spin" style="color:var(--dv-teal);"></i> Processing Request...</h2><p>Please wait while we connect to Power Automate and Adobe Sign. Do not close this window.</p></div>';
        }

        try {
            let filesData = [];
            for (let i = 0; i < documentFiles.length; i++) {
                const base64String = await convertFileToBase64(documentFiles[i]);
                filesData.push({
                    fileName: documentFiles[i].name,
                    fileContent: base64String
                });
            }

            const payload = {
                senderName: senderName,
                senderEmail: senderEmail,
                recipients: recipients,
                routingType: routingType,
                signaturePlacement: placementType,
                files: filesData
            };

            const response = await fetch(Config.powerAutomateWebhookUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const responseData = await response.json();
                isSuccessScreenActive = true; // تفعيل حماية الشاشة
                
                let successHtml = '';

                // 2. تصميم شاشة النجاح بناءً على نوع الإعداد
                if (placementType === 'Manual' && responseData && responseData.urls) {
                    // شاشة النجاح للإعداد اليدوي (زر واحد للذهاب للإعداد + زر إنهاء)
                    successHtml = `
                        <div style="text-align:center; padding: 40px;">
                            <i class="fa-solid fa-circle-check" style="color:#00A896; font-size: 50px; margin-bottom:20px;"></i>
                            <h2 style="color:var(--dv-dark-blue); margin-bottom: 10px;">Ready for Placement</h2>
                            <p style="color:var(--text-muted); margin-bottom: 20px;">Your document is successfully generated. Please click the button below to manually place the signature fields in Adobe Sign.</p>
                            
                            <button id="btnGoToPlacement" class="btn-primary" style="display:inline-block; margin-bottom:15px; width: 80%;">
                                <i class="fa-solid fa-pen-nib"></i> Go to Placement
                            </button>
                            <br>
                            <button id="btnSuccessOk" class="btn-secondary" style="width: 150px;">Done</button>
                        </div>
                    `;
                } else {
                    // شاشة النجاح للإعداد التلقائي (رسالة نجاح + زر موافق)
                    successHtml = `
                        <div style="text-align:center; padding: 40px;">
                            <i class="fa-solid fa-circle-check" style="color:#00A896; font-size: 50px; margin-bottom:20px;"></i>
                            <h2 style="color:var(--dv-dark-blue); margin-bottom: 10px;">Request Sent Successfully!</h2>
                            <p style="color:var(--text-muted); margin-bottom: 20px;">Your auto-workflow has been triggered and the recipients will be notified shortly.</p>
                            <button id="btnSuccessOk" class="btn-primary" style="margin-top: 20px; width: 150px;">OK</button>
                        </div>
                    `;
                }

                // عرض شاشة النجاح وربط الأزرار بالوظائف
                if (modalContent) {
                    modalContent.innerHTML = successHtml;
                    
                    // --- وظيفة الزر الجديد: فتح الروابط ---
                    const btnGoToPlacement = document.getElementById('btnGoToPlacement');
                    if (btnGoToPlacement) {
                        btnGoToPlacement.addEventListener('click', () => {
                            if (responseData && responseData.urls) {
                                responseData.urls.forEach(url => {
                                    if (url) window.open(url, '_blank');
                                });
                            }
                        });
                    }

                    // --- وظيفة زر الإنهاء: التنظيف الصامت ---
                    const btnSuccessOk = document.getElementById('btnSuccessOk');
                    if (btnSuccessOk) {
                        btnSuccessOk.addEventListener('click', () => {
                            // 1. إعادة محتوى النافذة المنبثقة لأصلها
                            modalContent.innerHTML = originalContent;
                            
                            // 2. تصفير النموذج وقائمة الملفات والمستلمين
                            const form = document.getElementById('signatureForm');
                            if (form) form.reset();
                            RecipientsModule.clearRecipients();
                            FilesModule.clearFiles();
                            
                            const fileNameDisplay = document.getElementById('fileNameDisplay');
                            if (fileNameDisplay) {
                                fileNameDisplay.innerHTML = 'Drag and drop your files here, or click to browse';
                                fileNameDisplay.style.display = 'block';
                            }
                            
                            // 3. إغلاق النافذة بأمان
                            closeOptionsModal();
                        });
                    }
                }

            } else {
                throw new Error("Server Response Error");
            }

        } catch (error) {
            console.error("Submission Error:", error);
            alert("An error occurred while sending the request. Please check your connection.");
        } finally {
            // إعادة الشاشة لحالتها الأصلية فقط في حال (فشل الطلب)
            if (modalContent && !isSuccessScreenActive) {
                modalContent.innerHTML = originalContent;
            }
        }
    }

    // Helper: Convert file to Base64
    function convertFileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result.split(',')[1]);
            reader.onerror = error => reject(error);
        });
    }

    return {
        openOptionsModal,
        closeOptionsModal,
        submitFinalRequest
    };
})();
