/* ==========================================
   1. GLOBAL CONFIGURATION & STATE
   ========================================== */
const Config = {
    msalConfig: {
        auth: {
            clientId: "05bcce89-b01a-4732-857f-8402f7a649f5", 
            authority: "https://login.microsoftonline.com/58d349d2-2319-41f9-9aff-f3055d0a1880", 
            knownAuthorities: ["login.microsoftonline.com"],
            redirectUri: "https://haider414.github.io/e-signature-portal/" 
        },
        cache: {
            // استخدام الذاكرة المؤقتة يضمن تسجيل الخروج التلقائي عند إغلاق الصفحة
            cacheLocation: "sessionStorage", 
            storeAuthStateInCookie: false    
        }
    },
    graphApiUrl: "https://graph.microsoft.com/v1.0",
    powerAutomateWebhookUrl: "https://default58d349d2231941f99afff3055d0a18.80.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/0a66f6833aa04ca68f94ff8d4b84b180/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZO_h_F9xftNmNbcfPkFnrGaqHKIpDdxQ3jWqmIklIkk",
    sharePointContext: {
        siteHostname: "dammamvalleyd.sharepoint.com", 
        sitePath: "/sites/E-SignatureSystem",
        listId: "4de9b715-9c83-4c02-a4aa-670346ce2254"
    }
};

let currentUserEmail = ""; 
let currentUserName = "";
let realRequests = [];
let selectedRecipients = [];
let dt = new DataTransfer();