/**
 * FNRG Portal — Frontend
 * Login by phone number -> Thursday Calling page (Google Sheets backend).
 *
 * - Admin sees all contacts; users see only contacts assigned to them.
 * - Editable in the portal: W/S, Calling Status (+ Assigned To for admin).
 *   Sessions, No. of Calls and Phone Number are read-only (sheet-only edits).
 * - "Others" status opens a popup asking for mandatory details text.
 * - Each row has Submit (locks to "Submitted" until the row changes again)
 *   and Send (opens WhatsApp chat with that contact).
 */

// PASTE YOUR DEPLOYED APPS SCRIPT WEB APP URL HERE:
const API_URL = "https://script.google.com/macros/s/AKfycbx5LrshnzKQLkXDHP095JGqYfsHEGN7k9zaD5jTR66YBMsq4Zx0WA-nV06pnGpR2ALR/exec";

const STORAGE_KEY = "fnrg_user";
const STATUS_DEFAULT = "Not Done";
const STATUS_OTHERS = "Others";

// --- State ---
let currentUser = null;   // { name, phone, role }
let contacts = [];        // [{ name, phone, ws, sessions, calls, cultivatedBy, assignedTo, status }]
let wsOptions = ["W", "S", "NA"];
let statusOptions = [STATUS_DEFAULT, STATUS_OTHERS];
let activeDate = "";      // name of the sheet's last date column, e.g. "Jul 11"
let userNames = [];       // admin only: names for the Assigned To / Cultivated By dropdowns
let userLimits = {};      // admin only: name -> call limit mapping
let registeredUserPhones = []; // admin only: list of registered user phone numbers to avoid assigning to
let dirtyContacts = new Set(); // set of phones with staged updates for bulk save
let activeCampaign = "Thursday Calling"; // default active campaign
let settings = {}; // settings dictionary for WhatsApp templates

function isAdmin() {
  return !!(currentUser && currentUser.role === "admin");
}

// --- DOM ---
const loginView = document.getElementById("login-view");
const appView = document.getElementById("app-view");
const loginForm = document.getElementById("login-form");
const phoneInput = document.getElementById("phone-input");
const loginBtn = document.getElementById("login-btn");
const loginError = document.getElementById("login-error");
const userNameEl = document.getElementById("user-name");
const roleBadge = document.getElementById("role-badge");
const activeDateEl = document.getElementById("active-date");
const contactsHead = document.getElementById("contacts-head-row");
const contactsBody = document.getElementById("contacts-body");
const syncStatus = document.getElementById("sync-status");
const refreshBtn = document.getElementById("refresh-btn");
const logoutBtn = document.getElementById("logout-btn");
const toastEl = document.getElementById("toast");

const othersModal = document.getElementById("others-modal");
const othersInput = document.getElementById("others-input");
const othersError = document.getElementById("others-error");
const othersConfirm = document.getElementById("others-confirm");
const othersCancel = document.getElementById("others-cancel");

// Stats elements
const statsSummaryBar = document.getElementById("stats-summary-bar");
const statTotal = document.getElementById("stat-total");
const statPositive = document.getElementById("stat-positive");
const statPending = document.getElementById("stat-pending");

// Dashboard & Tabs
const dashboardView = document.getElementById("campaign-dashboard");
const mainContent = document.getElementById("main-content");
const backToDashboardBtn = document.getElementById("back-to-dashboard-btn");
const adminTabs = document.getElementById("admin-tabs");

// Admin Sections
const masterSection = document.getElementById("master-data-section");
const masterContactsBody = document.getElementById("master-contacts-body");
const assignBtn = document.getElementById("assign-btn");
const saveAllBtn = document.getElementById("save-all-btn");

// History Modal elements
const historyModal = document.getElementById("history-modal");
const historyBody = document.getElementById("history-body");
const historyCloseBtn = document.getElementById("history-close-btn");
const historyContactInfo = document.getElementById("history-contact-info");
const historyModalTitle = document.getElementById("history-modal-title");
const historyTh2 = document.getElementById("history-th-2");
const historyTh3 = document.getElementById("history-th-3");
const historyTh4 = document.getElementById("history-th-4");

// Reception elements
const receptionSection = document.getElementById("reception-section");
const receptionSearch = document.getElementById("reception-search");
const receptionLoader = document.getElementById("reception-loader");
const receptionResult = document.getElementById("reception-result");
const receptionNotFound = document.getElementById("reception-not-found");
const receptionNameEl = document.getElementById("reception-name");
const receptionPhoneEl = document.getElementById("reception-phone");
const receptionSessionsEl = document.getElementById("reception-sessions");
const receptionSessionName = document.getElementById("reception-session-name");
const markAttendanceBtn = document.getElementById("mark-attendance-btn");

// Cultivation elements
const cultivationSection = document.getElementById("cultivation-section");
const cultivationTable = document.getElementById("cultivation-table");
const cultivationBody = document.getElementById("cultivation-body");
const cultivationEmpty = document.getElementById("cultivation-empty");

// Add Person modal elements
const adminAddPersonBtn = document.getElementById("admin-add-person-btn");
const receptionAddPersonBtn = document.getElementById("reception-add-person-btn");
const addPersonModal = document.getElementById("add-person-modal");
const addPersonTitle = document.getElementById("add-person-title");
const addPersonClose = document.getElementById("add-person-close");
const addPersonNameInput = document.getElementById("add-person-name");
const addPersonPhoneInput = document.getElementById("add-person-phone");
const addPersonError = document.getElementById("add-person-error");
const addPersonSubmitBtn = document.getElementById("add-person-submit");
const receptionAttendanceBody = document.getElementById("reception-attendance-body");

// Assigned Data summary elements
const assignedDataBtn = document.getElementById("assigned-data-btn");
const assignedDataModal = document.getElementById("assigned-data-modal");
const assignedDataClose = document.getElementById("assigned-data-close");
const assignedDataBody = document.getElementById("assigned-data-body");

// Import/Export elements
const adminImportBtn = document.getElementById("admin-import-btn");
const adminExportBtn = document.getElementById("admin-export-btn");
const importFileInput = document.getElementById("import-file-input");
const importMappingModal = document.getElementById("import-mapping-modal");
const importMappingClose = document.getElementById("import-mapping-close");
const importMappingSubmit = document.getElementById("import-mapping-submit");
const mappingNameSelect = document.getElementById("mapping-name");
const mappingPhoneSelect = document.getElementById("mapping-phone");
const importMappingError = document.getElementById("import-mapping-error");

// Admin Festival section elements
const adminFestivalSection = document.getElementById("admin-festival-section");
const adminFestivalFilter = document.getElementById("admin-festival-filter");
const autoAssignUsersList = document.getElementById("auto-assign-users-list");
const adminFestivalContactsBody = document.getElementById("admin-festival-contacts-body");

// Promotion Message Template elements
const adminMessageBtnCalling = document.getElementById("admin-message-btn-calling");
const adminMessageBtnFestival = document.getElementById("admin-message-btn-festival");
const messageModal = document.getElementById("message-modal");
const messageModalTitle = document.getElementById("message-modal-title");
const messageModalClose = document.getElementById("message-modal-close");
const messageTemplateInput = document.getElementById("message-template-input");
const messagePosterUpload = document.getElementById("message-poster-upload");
const messagePosterPreviewContainer = document.getElementById("message-poster-preview-container");
const messagePosterPreview = document.getElementById("message-poster-preview");
const messagePosterRemoveBtn = document.getElementById("message-poster-remove-btn");
const messageModalError = document.getElementById("message-modal-error");
const saveMessageBtn = document.getElementById("save-message-btn");

const campaignPosterBanner = document.getElementById("campaign-poster-banner");
const campaignPosterImg = document.getElementById("campaign-poster-img");

// ============================================================
// Init
// ============================================================

document.addEventListener("DOMContentLoaded", () => {
  loginForm.addEventListener("submit", onLoginSubmit);
  refreshBtn.addEventListener("click", refreshContacts);
  logoutBtn.addEventListener("click", logout);
  setupOthersModal();
  setupReception();

  // PWA Back Button / popstate routing
  window.addEventListener("popstate", (event) => {
    if (event.state && event.state.page) {
      if (event.state.page === "dashboard") {
        returnToDashboard(true);
      } else {
        selectCampaign(event.state.page, true);
      }
    } else {
      returnToDashboard(true);
    }
  });

  // Add Person click events
  if (adminAddPersonBtn) adminAddPersonBtn.addEventListener("click", () => openAddPersonModal("admin"));
  if (receptionAddPersonBtn) receptionAddPersonBtn.addEventListener("click", () => openAddPersonModal("reception"));
  if (addPersonClose) addPersonClose.addEventListener("click", closeAddPersonModal);
  if (addPersonModal) {
    addPersonModal.addEventListener("click", (e) => {
      if (e.target === addPersonModal) closeAddPersonModal();
    });
  }
  if (addPersonSubmitBtn) addPersonSubmitBtn.addEventListener("click", submitAddPersonForm);

  // Assigned Data click events
  if (assignedDataBtn) assignedDataBtn.addEventListener("click", openAssignedDataModal);
  if (assignedDataClose) {
    assignedDataClose.addEventListener("click", () => assignedDataModal.classList.remove("active"));
  }
  if (assignedDataModal) {
    assignedDataModal.addEventListener("click", (e) => {
      if (e.target === assignedDataModal) assignedDataModal.classList.remove("active");
    });
  }

  // Import/Export click events
  if (adminExportBtn) adminExportBtn.addEventListener("click", exportToExcel);
  if (adminImportBtn) adminImportBtn.addEventListener("click", () => {
    importFileInput.value = "";
    importFileInput.click();
  });
  if (importFileInput) importFileInput.addEventListener("change", handleFileImport);
  if (importMappingClose) {
    importMappingClose.addEventListener("click", () => importMappingModal.classList.remove("active"));
  }
  if (importMappingModal) {
    importMappingModal.addEventListener("click", (e) => {
      if (e.target === importMappingModal) importMappingModal.classList.remove("active");
    });
  }
  if (importMappingSubmit) importMappingSubmit.addEventListener("click", submitImportData);

  historyCloseBtn.addEventListener("click", () => {
    historyModal.classList.remove("active");
  });

  // Click outside to close history modal
  historyModal.addEventListener("click", (e) => {
    if (e.target === historyModal) historyModal.classList.remove("active");
  });

  assignBtn.addEventListener("click", autoAssignContacts);
  saveAllBtn.addEventListener("click", submitBulkChanges);
  backToDashboardBtn.addEventListener("click", returnToDashboard);

  let currentMessageKey = "";
  let pendingPosterFile = null;

  function openMessageModal(type) {
    currentMessageKey = type;
    pendingPosterFile = null;
    if (messagePosterUpload) messagePosterUpload.value = "";

    if (messageModalTitle) {
      messageModalTitle.textContent = type === "festival_message" ? "Festival Promotions Message" : "Thursday Calling Message";
    }
    if (messageTemplateInput) {
      messageTemplateInput.value = settings[type] || "";
    }
    
    // Preview existing poster if configured
    const posterKey = type === "festival_message" ? "festival_poster" : "calling_poster";
    const currentPosterUrl = (settings && settings[posterKey]) || "";
    if (currentPosterUrl && messagePosterPreview && messagePosterPreviewContainer) {
      messagePosterPreview.src = currentPosterUrl;
      messagePosterPreviewContainer.classList.remove("hidden");
    } else if (messagePosterPreviewContainer) {
      messagePosterPreviewContainer.classList.add("hidden");
    }

    if (messageModalError) messageModalError.textContent = "";
    if (messageModal) messageModal.classList.add("active");
  }

  // Handle local poster file selection & preview
  if (messagePosterUpload) {
    messagePosterUpload.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = () => {
        pendingPosterFile = {
          data: reader.result,
          mimeType: file.type,
          name: file.name
        };
        if (messagePosterPreview && messagePosterPreviewContainer) {
          messagePosterPreview.src = reader.result;
          messagePosterPreviewContainer.classList.remove("hidden");
        }
      };
      reader.readAsDataURL(file);
    });
  }

  // Handle removing poster image
  if (messagePosterRemoveBtn) {
    messagePosterRemoveBtn.addEventListener("click", () => {
      if (messagePosterUpload) messagePosterUpload.value = "";
      pendingPosterFile = { remove: true };
      if (messagePosterPreviewContainer) messagePosterPreviewContainer.classList.add("hidden");
      if (messagePosterPreview) messagePosterPreview.src = "";
    });
  }

  async function saveMessageTemplate() {
    if (!currentMessageKey) return;
    const inputVal = messageTemplateInput.value.trim();
    
    saveMessageBtn.classList.add("busy");
    saveMessageBtn.disabled = true;
    if (messageModalError) messageModalError.textContent = "";

    // 1. Save template message text
    const response = await api(null, {
      action: "saveSetting",
      requesterPhone: currentUser.phone,
      key: currentMessageKey,
      value: encodeURIComponent(inputVal)
    });

    if (response.status !== "success") {
      saveMessageBtn.classList.remove("busy");
      saveMessageBtn.disabled = false;
      if (messageModalError) {
        messageModalError.textContent = response.message || "Failed to save message template.";
      }
      return;
    }

    // 2. Save poster file if changed
    const posterKey = currentMessageKey === "festival_message" ? "festival_poster" : "calling_poster";
    let posterResponse = { status: "success" };

    if (pendingPosterFile) {
      if (pendingPosterFile.remove) {
        // Clear poster setting
        posterResponse = await api(null, {
          action: "saveSetting",
          requesterPhone: currentUser.phone,
          key: posterKey,
          value: ""
        });
      } else {
        // Upload new poster file to Google Drive via backend
        posterResponse = await api(null, {
          action: "saveSettingImage",
          requesterPhone: currentUser.phone,
          key: posterKey,
          base64Data: pendingPosterFile.data,
          mimeType: pendingPosterFile.mimeType,
          fileName: pendingPosterFile.name
        });
      }
    }

    saveMessageBtn.classList.remove("busy");
    saveMessageBtn.disabled = false;

    if (posterResponse.status === "success") {
      if (posterResponse.settings) settings = posterResponse.settings;
      if (messageModal) messageModal.classList.remove("active");
      showToast("Template and poster saved successfully!", "success");
      renderPosterBanner();
    } else {
      if (messageModalError) {
        messageModalError.textContent = posterResponse.message || "Failed to upload poster.";
      } else {
        showToast(posterResponse.message || "Failed to upload poster.", "error");
      }
    }
  }


  if (adminMessageBtnCalling) {
    adminMessageBtnCalling.addEventListener("click", () => openMessageModal("calling_message"));
  }
  if (adminMessageBtnFestival) {
    adminMessageBtnFestival.addEventListener("click", () => openMessageModal("festival_message"));
  }
  if (messageModalClose) {
    messageModalClose.addEventListener("click", () => messageModal.classList.remove("active"));
  }
  if (messageModal) {
    messageModal.addEventListener("click", (e) => {
      if (e.target === messageModal) messageModal.classList.remove("active");
    });
  }
  if (saveMessageBtn) {
    saveMessageBtn.addEventListener("click", saveMessageTemplate);
  }

  document.querySelectorAll(".dashboard-card").forEach(card => {
    card.addEventListener("click", () => {
      if (card.classList.contains("disabled")) return;
      const campaign = card.dataset.campaign;
      selectCampaign(campaign);
    });
  });

  document.querySelectorAll(".admin-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".admin-tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");

      const target = tab.dataset.target;
      if (target === "master-data-section") {
        activeCampaign = "Thursday Calling";
        masterSection.classList.remove("hidden");
        document.getElementById("calling-section").style.display = "none";
        if (adminFestivalSection) adminFestivalSection.classList.add("hidden");
        if (assignedDataBtn) assignedDataBtn.classList.add("hidden");
        refreshContacts(); // Load Thursday Calling master data
      } else if (target === "calling-section") {
        activeCampaign = "Thursday Calling";
        masterSection.classList.add("hidden");
        document.getElementById("calling-section").style.display = "";
        if (adminFestivalSection) adminFestivalSection.classList.add("hidden");
        if (assignedDataBtn) assignedDataBtn.classList.remove("hidden");
        refreshContacts(); // Load Thursday Calling round data
      } else {
        activeCampaign = "Festival Promotions";
        masterSection.classList.add("hidden");
        document.getElementById("calling-section").style.display = "none";
        if (adminFestivalSection) {
          adminFestivalSection.classList.remove("hidden");
          loadAdminFestivalData(); // Load Festival Promotions data
        }
        if (assignedDataBtn) assignedDataBtn.classList.add("hidden");
      }
    });
  });

  phoneInput.addEventListener("input", () => {
    phoneInput.value = phoneInput.value.replace(/\D/g, "").slice(0, 10);
    loginError.textContent = "";
  });

  // Auto-login if a user is remembered on this device
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      const user = JSON.parse(saved);
      if (user && user.phone) {
        phoneInput.value = user.phone;
        attemptLogin(user.phone, true);
      }
    } catch (ignored) {
      localStorage.removeItem(STORAGE_KEY);
    }
  }

  startBackgroundPolling();
  registerServiceWorker();
});

// Poster banner rendering — must be at module scope so renderContacts() can call it
function renderPosterBanner() {
  if (!campaignPosterBanner || !campaignPosterImg) return;
  const posterKey = activeCampaign === "Festival Promotions" ? "festival_poster" : "calling_poster";
  const posterUrl = (settings && settings[posterKey]) || "";
  const showingCallingSection = document.getElementById("calling-section").style.display !== "none" && !mainContent.classList.contains("hidden");

  if (showingCallingSection && posterUrl) {
    campaignPosterImg.src = posterUrl;
    campaignPosterBanner.classList.remove("hidden");
  } else {
    campaignPosterBanner.classList.add("hidden");
  }
}

// ============================================================
// API helper — always resolves to a {status, ...} object
// ============================================================

async function api(params, postBody) {
  if (params && (params.action === "data" || params.action === "login")) {
    if (!params.sheet) params.sheet = activeCampaign;
  }
  if (postBody) {
    if (!postBody.sheet) postBody.sheet = activeCampaign;
  }

  const queryParams = params ? { ...params, _t: Date.now() } : { _t: Date.now() };
  const url = API_URL + "?" + new URLSearchParams(queryParams).toString();
  const options = postBody
    ? {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(postBody)
      }
    : { method: "GET" };

  try {
    const response = await fetch(url, options);
    if (!response.ok) {
      return { status: "error", code: "HTTP_" + response.status, message: "Server error (" + response.status + ")" };
    }
    const data = await response.json();
    if (!data || typeof data.status === "undefined") {
      return { status: "error", code: "BAD_RESPONSE", message: "Unexpected response from server" };
    }
    return data;
  } catch (err) {
    return {
      status: "error",
      code: "NETWORK",
      message: navigator.onLine
        ? "Could not reach the server. Please try again."
        : "You are offline. Please check your connection."
    };
  }
}

// Applies a successful login/data/save payload to local state
function applyPayload(data) {
  if (data.user) currentUser = data.user;
  contacts = data.contacts || [];
  if (Array.isArray(data.wsOptions) && data.wsOptions.length) wsOptions = data.wsOptions;
  if (Array.isArray(data.statusOptions) && data.statusOptions.length) statusOptions = data.statusOptions;
  if (typeof data.activeDate === "string") activeDate = data.activeDate;
  if (Array.isArray(data.userNames)) userNames = data.userNames;
  if (data.userLimits) userLimits = data.userLimits;
  if (Array.isArray(data.userPhones)) registeredUserPhones = data.userPhones;
  if (data.settings) settings = data.settings;
}

// ============================================================
// Login / Logout
// ============================================================

function onLoginSubmit(event) {
  event.preventDefault();
  const phone = phoneInput.value.trim();
  if (phone.length !== 10) {
    loginError.textContent = "Please enter a valid 10-digit phone number.";
    return;
  }
  attemptLogin(phone, false);
}

async function attemptLogin(phone, silent) {
  setLoginBusy(true);
  loginError.textContent = "";

  const data = await api({ action: "login", phone: phone });
  setLoginBusy(false);

  if (data.status === "success" && data.user) {
    applyPayload(data);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(currentUser));
    showApp();
  } else {
    // Auto-login failure falls back to the login screen quietly
    if (data.code === "USER_NOT_FOUND") {
      logout();
    }
    if (!silent || data.code === "USER_NOT_FOUND") {
      loginError.textContent = data.message || "Login failed. Please try again.";
    } else {
      showToast(data.message || "Could not connect. Pull to retry.", "error");
    }
  }
}

function setLoginBusy(busy) {
  loginBtn.disabled = busy;
  loginBtn.classList.toggle("busy", busy);
  loginBtn.querySelector(".btn-label").textContent = busy ? "Checking…" : "Enter";
}

function logout() {
  localStorage.removeItem(STORAGE_KEY);
  currentUser = null;
  contacts = [];
  userNames = [];
  userLimits = {};
  dirtyContacts.clear();
  phoneInput.value = "";
  loginError.textContent = "";
  masterContactsBody.innerHTML = "";
  contactsBody.innerHTML = "";

  dashboardView.classList.add("hidden");
  mainContent.classList.add("hidden");
  adminTabs.classList.add("hidden");
  backToDashboardBtn.classList.add("hidden");
  statsSummaryBar.classList.add("hidden");
  saveAllBtn.classList.add("hidden");
  saveAllBtn.disabled = true;

  appView.classList.add("hidden");
  loginView.classList.remove("hidden");
}

function showApp() {
  userNameEl.textContent = currentUser.name;
  roleBadge.classList.toggle("hidden", !isAdmin());

  loginView.classList.add("hidden");
  appView.classList.remove("hidden");

  if (isAdmin()) {
    // Admin skips dashboard — goes straight to Festival Promotions
    dashboardView.classList.add("hidden");
    activeCampaign = "Festival Promotions";
    selectCampaign("Festival Promotions", true);
  } else {
    // Regular users see the campaign dashboard
    dashboardView.classList.remove("hidden");
    mainContent.classList.add("hidden");
    adminTabs.classList.add("hidden");
    backToDashboardBtn.classList.add("hidden");
    statsSummaryBar.classList.add("hidden");
    
    // Clear hash and push dashboard state
    window.history.replaceState({ page: "dashboard" }, "", "./");
  }
}

function selectCampaign(name, skipPushHistory) {
  activeCampaign = name;
  
  // Update campaign title head immediately
  const titleEl = document.getElementById("calling-page-title");
  if (titleEl) {
    let displayName = name;
    if (name === "Festival Promotions" && !isAdmin() && currentUser && currentUser.festival) {
      displayName = currentUser.festival;
    }
    titleEl.innerHTML = displayName + ' <span id="active-date" class="active-date"></span>';
  }

  // Clear campaign contacts and state immediately to prevent visual bleed during loads
  contacts = [];
  activeDate = "";
  contactsBody.innerHTML = "";
  masterContactsBody.innerHTML = "";
  if (typeof adminFestivalContactsBody !== "undefined" && adminFestivalContactsBody) {
    adminFestivalContactsBody.innerHTML = "";
  }

  dashboardView.classList.add("hidden");
  mainContent.classList.remove("hidden");
  backToDashboardBtn.classList.toggle("hidden", isAdmin());

  // Hide all campaign-specific sections
  masterSection.classList.add("hidden");
  document.getElementById("calling-section").style.display = "none";
  adminTabs.classList.add("hidden");
  statsSummaryBar.classList.add("hidden");
  saveAllBtn.classList.add("hidden");
  assignBtn.classList.add("hidden");
  if (receptionSection) receptionSection.classList.add("hidden");
  if (cultivationSection) cultivationSection.classList.add("hidden");

  // Push history state unless skipped
  if (!skipPushHistory && !isAdmin()) {
    const slug = name.replace(/\s+/g, "-").toLowerCase();
    window.history.pushState({ page: name }, "", "./#" + slug);
  }

  if (name === "Reception") {
    if (receptionSection) {
      receptionSection.classList.remove("hidden");
      receptionSearch.value = "";
      receptionResult.classList.add("hidden");
      receptionNotFound.classList.add("hidden");
      renderReceptionAttendanceList(); // Load today's marked list
    }
    return;
  }

  if (name === "Cultivation") {
    if (cultivationSection) {
      cultivationSection.classList.remove("hidden");
      loadCultivationContacts();
    }
    return;
  }

  // Standard calling campaigns (Thursday Calling / Festival Promotions)
  const isAdm = isAdmin();
  adminTabs.classList.toggle("hidden", !isAdm);

  if (isAdm) {
    // Hide Thursday Calling and Calling Round tabs, activate Festival tab
    document.querySelectorAll(".admin-tab").forEach(t => {
      if (t.dataset.target === "admin-festival-section") {
        t.style.display = "";
        t.classList.remove("hidden");
        t.classList.add("active");
      } else {
        t.style.display = "none";
        t.classList.add("hidden");
        t.classList.remove("active");
      }
    });

    masterSection.classList.add("hidden");
    document.getElementById("calling-section").style.display = "none";
    if (adminFestivalSection) adminFestivalSection.classList.remove("hidden");
    if (assignedDataBtn) assignedDataBtn.classList.add("hidden");

    statsSummaryBar.classList.add("hidden");
    saveAllBtn.classList.add("hidden");
    assignBtn.classList.add("hidden");
    if (adminMessageBtnCalling) adminMessageBtnCalling.classList.add("hidden");
  } else {
    masterSection.classList.add("hidden");
    document.getElementById("calling-section").style.display = "";
    statsSummaryBar.classList.remove("hidden");
    
    saveAllBtn.classList.add("hidden");
    assignBtn.classList.add("hidden");
    if (adminMessageBtnCalling) adminMessageBtnCalling.classList.add("hidden");
  }

  // Render cached contacts immediately before network load
  loadCachedContacts();
  refreshContacts();
}

function returnToDashboard(skipPushHistory) {
  contacts = [];
  dirtyContacts.clear();
  saveAllBtn.disabled = true;
  saveAllBtn.classList.add("hidden");

  contactsBody.innerHTML = "";
  masterContactsBody.innerHTML = "";

  mainContent.classList.add("hidden");
  adminTabs.classList.add("hidden");
  backToDashboardBtn.classList.add("hidden");
  dashboardView.classList.remove("hidden");
  statsSummaryBar.classList.add("hidden");
  if (campaignPosterBanner) campaignPosterBanner.classList.add("hidden");

  if (receptionSection) receptionSection.classList.add("hidden");
  if (cultivationSection) cultivationSection.classList.add("hidden");

  if (!skipPushHistory && !isAdmin()) {
    window.history.pushState({ page: "dashboard" }, "", "./");
  }
}

// ============================================================
// "Others" status popup
// ============================================================

let othersHandlers = null; // { onConfirm(text), onCancel() }

function setupOthersModal() {
  othersConfirm.addEventListener("click", () => {
    const text = othersInput.value.trim();
    if (!text) {
      othersError.textContent = "Details are mandatory — please type something.";
      othersInput.focus();
      return;
    }
    closeOthersModal();
    if (othersHandlers) othersHandlers.onConfirm(text);
  });

  othersCancel.addEventListener("click", () => {
    closeOthersModal();
    if (othersHandlers) othersHandlers.onCancel();
  });

  othersInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") othersConfirm.click();
    if (event.key === "Escape") othersCancel.click();
  });
}

function openOthersModal(handlers) {
  othersHandlers = handlers;
  othersInput.value = "";
  othersError.textContent = "";
  othersModal.classList.add("active");
  setTimeout(() => othersInput.focus(), 60);
}

function closeOthersModal() {
  othersModal.classList.remove("active");
}

// ============================================================
// Thursday Calling — render
// ====================================================// ============ Dynamic Row Render Helper ============

function createContactRow(contact, isMasterTable, isCultivation) {
  const tr = document.createElement("tr");
  tr.dataset.phone = contact.phone; // save key for lookup
  const controls = {};

  const submitBtn = document.createElement("button");
  submitBtn.type = "button";
  submitBtn.className = "row-submit-btn";
  submitBtn.textContent = "Submit";
  submitBtn.setAttribute("aria-label", "Submit " + contact.name);

  // Master row doesn't have status, so it's always "submitted" unless edited
  let submitted = isMasterTable ? true : (contact.status !== STATUS_DEFAULT);
  let dirty = false;

  const updateSubmitState = () => {
    if (isMasterTable) {
      submitBtn.disabled = !dirty;
      submitBtn.textContent = dirty ? "Submit" : "Submitted";
    } else {
      const statusValue = controls.status ? controls.status.value : contact.status;
      const noStatus = statusValue === STATUS_DEFAULT || statusValue === STATUS_OTHERS;
      if (!isAdmin() && noStatus) {
        submitBtn.disabled = true;
        submitBtn.textContent = "Submit";
      } else if (submitted && !dirty) {
        submitBtn.disabled = true;
        submitBtn.textContent = "Submitted";
      } else {
        submitBtn.disabled = false;
        submitBtn.textContent = "Submit";
      }
    }
  };

  const markDirty = () => {
    dirty = true;
    tr.classList.add("row-dirty");
    controls.updateSubmitState();
    if (isAdmin() && !isMasterTable) {
      dirtyContacts.add(contact.phone);
      saveAllBtn.disabled = false;
    }
  };
  controls.setSubmitted = () => { submitted = true; dirty = false; tr.classList.remove("row-dirty"); controls.updateSubmitState(); };
  controls.updateSubmitState = updateSubmitState;

  // 1. Name (with edit on mobile fix)
  const nameTd = document.createElement("td");
  nameTd.className = "cell-name";
  nameTd.dataset.label = "Name";

  const nameWrap = document.createElement("div");
  nameWrap.className = "name-cell-wrap";

  const nameDisplay = document.createElement("span");
  nameDisplay.className = "name-display";
  nameDisplay.textContent = contact.name;
  const nameHint = document.createElement("span");
  nameHint.className = "name-edit-hint";
  nameHint.textContent = "✎";
  nameDisplay.appendChild(nameHint);

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.className = "name-edit-input";
  nameInput.maxLength = 100;
  nameInput.setAttribute("aria-label", "Edit contact name");
  nameInput.style.display = "none";

  const nameActions = document.createElement("span");
  nameActions.className = "name-edit-actions";
  nameActions.style.display = "none";
  const nameSaveBtn = document.createElement("button");
  nameSaveBtn.type = "button";
  nameSaveBtn.className = "name-action-btn save";
  nameSaveBtn.textContent = "Save";
  const nameCancelBtn = document.createElement("button");
  nameCancelBtn.type = "button";
  nameCancelBtn.className = "name-action-btn cancel";
  nameCancelBtn.textContent = "Cancel";
  nameActions.appendChild(nameSaveBtn);
  nameActions.appendChild(nameCancelBtn);

  function startNameEdit() {
    nameDisplay.style.display = "none";
    nameInput.value = nameDisplay.firstChild.textContent || nameDisplay.textContent.replace("✎", "").trim();
    nameInput.style.display = "";
    nameActions.style.display = "";
    setTimeout(() => { nameInput.focus(); nameInput.select(); }, 30);
  }

  function cancelNameEdit() {
    nameInput.style.display = "none";
    nameActions.style.display = "none";
    nameDisplay.style.display = "";
  }

  function saveNameEdit() {
    const newName = nameInput.value.trim();
    if (!newName) { nameInput.focus(); return; }
    nameDisplay.firstChild.textContent = newName;
    cancelNameEdit();
    markDirty();
  }

  nameDisplay.addEventListener("dblclick", startNameEdit);
  nameDisplay.addEventListener("click", (e) => {
    if (e.target.classList.contains("name-edit-hint")) {
      startNameEdit();
      e.stopPropagation();
    }
  });
  nameSaveBtn.addEventListener("click", saveNameEdit);
  nameCancelBtn.addEventListener("click", cancelNameEdit);
  nameInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") saveNameEdit();
    if (e.key === "Escape") cancelNameEdit();
  });

  controls.name = {
    get value() {
      const text = nameDisplay.firstChild ? nameDisplay.firstChild.textContent : nameDisplay.textContent;
      return text.replace("✎", "").trim();
    }
  };

  nameWrap.appendChild(nameDisplay);
  if (contact.campaign) {
    const badge = document.createElement("span");
    badge.className = "cultivation-campaign-badge";
    badge.textContent = contact.campaign;
    nameWrap.appendChild(badge);
  }
  nameWrap.appendChild(nameInput);
  nameWrap.appendChild(nameActions);
  nameTd.appendChild(nameWrap);
  tr.appendChild(nameTd);

  // 2. Phone
  const phoneTd = document.createElement("td");
  phoneTd.className = "cell-phone";
  phoneTd.dataset.label = "Phone";
  const tel = document.createElement("a");
  tel.href = "tel:+91" + contact.phone;
  tel.textContent = formatPhone(contact.phone);
  phoneTd.appendChild(tel);
  tr.appendChild(phoneTd);

  // 3. W/S
  const wsTd = document.createElement("td");
  wsTd.className = "cell-ws";
  wsTd.dataset.label = "W/S";
  const wsSelect = document.createElement("select");
  wsSelect.className = "ws-select";
  wsSelect.setAttribute("aria-label", contact.name + " W/S");
  wsOptions.forEach((opt) => {
    const option = document.createElement("option");
    option.value = opt;
    option.textContent = opt;
    if (opt === contact.ws) option.selected = true;
    wsSelect.appendChild(option);
  });
  wsSelect.dataset.value = wsSelect.value;
  wsSelect.addEventListener("change", () => {
    wsSelect.dataset.value = wsSelect.value;
    markDirty();
  });
  controls.ws = wsSelect;
  wsTd.appendChild(wsSelect);
  tr.appendChild(wsTd);

  // 4. Sessions
  const sessionsTd = document.createElement("td");
  sessionsTd.className = "cell-readonly-num cell-sessions";
  sessionsTd.dataset.label = "Sessions";
  sessionsTd.textContent = contact.sessions;
  sessionsTd.addEventListener("click", () => showSessionHistoryModal(contact));
  tr.appendChild(sessionsTd);

  // 5. No. of Calls (Clickable to open modal popup)
  const callsTd = document.createElement("td");
  callsTd.className = "cell-readonly-num cell-calls";
  callsTd.dataset.label = "Calls";
  callsTd.textContent = contact.calls;
  callsTd.addEventListener("click", () => showCallHistoryModal(contact));
  controls.callsTd = callsTd;
  tr.appendChild(callsTd);

  // 5.5 Event Column (Festival Promotions only, for calling round view)
  if (activeCampaign === "Festival Promotions" && !isMasterTable && isAdmin()) {
    const eventTd = document.createElement("td");
    eventTd.className = "cell-event";
    eventTd.dataset.label = "Event";
    const eventSelect = document.createElement("select");
    eventSelect.className = "status-select";
    eventSelect.style.minWidth = "120px";
    eventSelect.setAttribute("aria-label", contact.name + " event");
    const emptyOpt = document.createElement("option");
    emptyOpt.value = "";
    emptyOpt.textContent = "— Select Event —";
    eventSelect.appendChild(emptyOpt);
    
    (loadedFestivals || []).forEach(f => {
      const opt = document.createElement("option");
      opt.value = f;
      opt.textContent = f;
      if (f === contact.event) opt.selected = true;
      eventSelect.appendChild(opt);
    });
    if (contact.event && !(loadedFestivals || []).includes(contact.event)) {
      const customOpt = document.createElement("option");
      customOpt.value = contact.event;
      customOpt.textContent = contact.event;
      customOpt.selected = true;
      eventSelect.appendChild(customOpt);
    }
    
    eventSelect.addEventListener("change", () => {
      contact.event = eventSelect.value;
      markDirty();
    });
    controls.eventSelect = eventSelect;
    eventTd.appendChild(eventSelect);
    tr.appendChild(eventTd);
  }

  if (isMasterTable) {
    // 6. Cultivated By
    const cultivatedTd = document.createElement("td");
    cultivatedTd.className = "cell-cultivated";
    cultivatedTd.dataset.label = "Cultivated By";
    const cultivatedSelect = document.createElement("select");
    cultivatedSelect.className = "assign-select";
    cultivatedSelect.setAttribute("aria-label", contact.name + " cultivated by");

    const uncultivated = document.createElement("option");
    uncultivated.value = "";
    uncultivated.textContent = "— Unassigned —";
    cultivatedSelect.appendChild(uncultivated);

    userNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      if (name === contact.cultivatedBy) option.selected = true;
      cultivatedSelect.appendChild(option);
    });
    // Orphan handling
    if (contact.cultivatedBy && !userNames.includes(contact.cultivatedBy)) {
      const orphan = document.createElement("option");
      orphan.value = contact.cultivatedBy;
      orphan.textContent = contact.cultivatedBy;
      orphan.selected = true;
      cultivatedSelect.appendChild(orphan);
    }
    cultivatedSelect.addEventListener("change", markDirty);
    controls.cultivatedBy = cultivatedSelect;
    cultivatedTd.appendChild(cultivatedSelect);
    tr.appendChild(cultivatedTd);

    // 7. Actions (Submit only)
    const actionsTd = document.createElement("td");
    actionsTd.className = "cell-actions";
    submitBtn.addEventListener("click", () => submitMasterRow(contact, tr, controls, submitBtn));
    actionsTd.appendChild(submitBtn);
    tr.appendChild(actionsTd);

  } else {

    // 6. Calling Status
    const statusTd = document.createElement("td");
    statusTd.className = "cell-status";
    statusTd.dataset.label = "Calling Status";
    const statusSelect = document.createElement("select");
    statusSelect.className = "status-select";
    statusSelect.setAttribute("aria-label", contact.name + " calling status");
    statusOptions.forEach((opt) => {
      const option = document.createElement("option");
      option.value = opt;
      option.textContent = opt;
      statusSelect.appendChild(option);
    });
    if (contact.status && !statusOptions.includes(contact.status)) {
      const custom = document.createElement("option");
      custom.value = contact.status;
      custom.textContent = contact.status;
      statusSelect.appendChild(custom);
    }
    statusSelect.value = statusOptions.includes(contact.status) || contact.status ? contact.status : STATUS_DEFAULT;
    if (!statusSelect.value) statusSelect.value = STATUS_DEFAULT;
    statusSelect.dataset.value = statusSelect.value;
    statusSelect.dataset.prev = statusSelect.value;

    statusSelect.addEventListener("change", () => {
      if (statusSelect.value === STATUS_OTHERS) {
        openOthersModal({
          onConfirm: (text) => {
            let custom = statusSelect.querySelector("option[data-custom]");
            if (!custom) {
              custom = document.createElement("option");
              custom.setAttribute("data-custom", "1");
              statusSelect.appendChild(custom);
            }
            custom.value = text;
            custom.textContent = text;
            statusSelect.value = text;
            statusSelect.dataset.value = text;
            statusSelect.dataset.prev = text;
            contact.status = text; // temporarily update state
            updateStatsBar();
            markDirty();
          },
          onCancel: () => {
            statusSelect.value = statusSelect.dataset.prev;
            statusSelect.dataset.value = statusSelect.value;
          }
        });
      } else {
        statusSelect.dataset.value = statusSelect.value;
        statusSelect.dataset.prev = statusSelect.value;
        contact.status = statusSelect.value; // temporarily update state
        updateStatsBar();
        markDirty();
      }
    });
    controls.status = statusSelect;
    statusTd.appendChild(statusSelect);
    tr.appendChild(statusTd);

    // 7. Assigned To (Admin only)
    if (isAdmin()) {
      tr.classList.add("with-assigned");
      const assignTd = document.createElement("td");
      assignTd.className = "cell-assigned";
      assignTd.dataset.label = "Assigned To";
      const assignSelect = document.createElement("select");
      assignSelect.className = "assign-select";
      assignSelect.setAttribute("aria-label", contact.name + " assigned to");

      const unassigned = document.createElement("option");
      unassigned.value = "";
      unassigned.textContent = "— Unassigned —";
      assignSelect.appendChild(unassigned);

      userNames.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        if (name === contact.assignedTo) option.selected = true;
        assignSelect.appendChild(option);
      });
      if (contact.assignedTo && !userNames.includes(contact.assignedTo)) {
        const orphan = document.createElement("option");
        orphan.value = contact.assignedTo;
        orphan.textContent = contact.assignedTo;
        orphan.selected = true;
        assignSelect.appendChild(orphan);
      }
      
      const updateAssignedStyle = () => {
        if (assignSelect.value !== "") {
          tr.classList.add("row-assigned");
        } else {
          tr.classList.remove("row-assigned");
        }
      };

      assignSelect.addEventListener("change", () => {
        markDirty();
        updateAssignedStyle();
      });

      // Initialize
      updateAssignedStyle();

      controls.assignedTo = assignSelect;
      assignTd.appendChild(assignSelect);
      tr.appendChild(assignTd);
    }

    // Track whether user clicked "Send" (made the call)
    let callMade = false;

    // Override updateSubmitState for non-admin users:
    // Submit stays disabled until Send is clicked, unless status is "Wrong Number"
    const origUpdateSubmitState = updateSubmitState;
    const gatedUpdateSubmitState = () => {
      try {
        if (isMasterTable || isAdmin()) {
          origUpdateSubmitState();
          return;
        }
        const statusValue = String(controls.status ? controls.status.value : (contact.status || ""));
        const noStatus = statusValue === STATUS_DEFAULT || statusValue === STATUS_OTHERS;
        const isWrongNumber = statusValue.toLowerCase() === "wrong number";

        if (noStatus) {
          submitBtn.disabled = true;
          submitBtn.textContent = "Submit";
        } else if (submitted && !dirty) {
          submitBtn.disabled = true;
          submitBtn.textContent = "Submitted";
        } else if (!isWrongNumber && !callMade) {
          // Status is set but not "Wrong Number" and user hasn't clicked Send yet
          submitBtn.disabled = true;
          submitBtn.textContent = "Call first";
        } else {
          submitBtn.disabled = false;
          submitBtn.textContent = "Submit";
        }
      } catch (err) {
        console.error("Error in gatedUpdateSubmitState:", err);
      }
    };
    controls.updateSubmitState = gatedUpdateSubmitState;

    // 8. Actions (Send to + Submit)
    const actionsTd = document.createElement("td");
    actionsTd.className = "cell-actions";
    
    if (!isAdmin()) {
      const sendBtn = document.createElement("button");
      sendBtn.type = "button";
      sendBtn.className = "send-btn";
      sendBtn.textContent = "Send message";
      sendBtn.setAttribute("aria-label", "WhatsApp message " + contact.name);
      sendBtn.addEventListener("click", () => {
        callMade = true;
        gatedUpdateSubmitState();
        const templateKey = activeCampaign === "Festival Promotions" ? "festival_message" : "calling_message";
        let messageText = (settings && settings[templateKey]) || "";
        messageText = messageText.replace(/{name}/g, contact.name);
        const url = "https://wa.me/91" + contact.phone + (messageText ? "?text=" + encodeURIComponent(messageText) : "");
        window.open(url, "_blank");
      });
      actionsTd.appendChild(sendBtn);

      submitBtn.addEventListener("click", () => submitRow(contact, tr, controls, submitBtn));
      actionsTd.appendChild(submitBtn);
    }
    
    tr.appendChild(actionsTd);
  }

  tr.controls = controls; // Attach controls for easy bulk lookup
  if (controls.updateSubmitState) controls.updateSubmitState();
  return tr;
}

// ============================================================
// Stats Bar Calculation
// ============================================================

function updateStatsBar() {
  if (isAdmin()) {
    statsSummaryBar.classList.add("hidden");
    return;
  }
  statsSummaryBar.classList.remove("hidden");

  const visibleContacts = contacts;
  const total = visibleContacts.length;

  let positive = 0;
  let pending = 0;

  visibleContacts.forEach(contact => {
    const status = (contact.status || "").trim().toLowerCase();
    if (status === "joining the session" || status === "will try to attend") {
      positive++;
    } else if (status === "not done" || status === "yet to call" || status === "") {
      pending++;
    }
  });

  statTotal.textContent = total;
  statPositive.textContent = positive;
  statPending.textContent = pending;
}

// ============================================================
// Thursday Calling — render
// ============================================================

function renderContacts() {
  // Update stats summary bar
  updateStatsBar();

  // Render campaign poster if configured
  renderPosterBanner();

  // Dynamic campaign heading and date
  const titleEl = document.getElementById("calling-page-title");
  if (titleEl) {
    let displayName = activeCampaign;
    if (activeCampaign === "Festival Promotions" && !isAdmin() && currentUser && currentUser.festival) {
      displayName = currentUser.festival;
    }
    titleEl.innerHTML = displayName + ' <span id="active-date" class="active-date">' + (activeDate ? "for " + activeDate : "") + '</span>';
  }

  // 1. Render Section 1: All People Data (Admin Only)
  if (isAdmin()) {
    masterContactsBody.innerHTML = "";
    if (!contacts.length) {
      masterContactsBody.innerHTML = '<tr><td colspan="7" class="loading-row">No master contacts found.</td></tr>';
    } else {
      contacts.forEach(contact => {
        const row = createContactRow(contact, true);
        masterContactsBody.appendChild(row);
      });
    }
  }

  // 2. Render Section 2: Thursday Calling (All Users)
  contactsBody.innerHTML = "";

  // Header row setup (Assigned To column depends on Admin)
  contactsHead.innerHTML = "";
  const headers = ["Name", "Phone Number", "W/S", "Sessions", "No. of Calls"];
  if (activeCampaign === "Festival Promotions" && isAdmin()) {
    headers.push("Event");
  }
  if (isAdmin()) {
    headers.push("Assigned To");
  }
  headers.push("Calling Status");
  headers.push(""); // actions column
  headers.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    contactsHead.appendChild(th);
  });

  if (!contacts.length) {
    const message = isAdmin()
      ? "No contacts found in the sheet yet."
      : "No contacts are assigned to you yet. Please contact your coordinator.";
    contactsBody.innerHTML =
      '<tr class="loading-row"><td colspan="' + headers.length + '">' + message + "</td></tr>";
    return;
  }

  contacts.forEach((contact) => {
    const row = createContactRow(contact, false);
    contactsBody.appendChild(row);
  });
}

// ============================================================
// Submit one row to the sheet
// ============================================================

async function submitRow(contact, tr, controls, btn) {
  // "Others" must have its details typed before submitting
  if (controls.status.value === STATUS_OTHERS) {
    showToast("Please type the details for the Others status first.", "error");
    return;
  }

  const newName = controls.name.value.trim();
  if (!newName) {
    showToast("Name cannot be empty.", "error");
    return;
  }

  const updates = {
    name: newName,
    ws: controls.ws.value,
    status: controls.status.value || STATUS_DEFAULT
  };
  if (controls.eventSelect) {
    updates.event = controls.eventSelect.value;
  }
  if (isAdmin()) {
    updates.assignedTo = controls.assignedTo.value;
  }

  // Instant feedback: thanks popup for 1s + button lightened & locked
  const startedAt = Date.now();
  btn.disabled = true;
  showToast("Thanks for submitting 🙏", "success", 1000);
  tr.classList.remove("row-saved", "row-error");
  tr.classList.add("row-saving");
  setSyncStatus("Saving…", "saving");

  const data = await api(null, {
    action: "updateContact",
    requesterPhone: currentUser.phone,
    phone: contact.phone,
    sheet: contact.campaign || activeCampaign,
    updates: updates
  });

  // Keep the thanks popup visible for at least 1 second before the result
  const popupWait = Math.max(0, 1000 - (Date.now() - startedAt));
  setTimeout(() => {
    tr.classList.remove("row-saving");

    if (data.status === "success") {
      Object.assign(contact, updates);
      // The backend returns fresh data — update this row's call count live
      const fresh = (data.contacts || []).find((c) => c.phone === contact.phone);
      if (fresh) {
        contact.calls = fresh.calls;
        controls.callsTd.textContent = fresh.calls;
      }
      
      // Sync master table row if visible
      const masterRow = document.querySelector(`#master-contacts-body tr[data-phone="${contact.phone}"]`);
      if (masterRow) {
        const callsTd = masterRow.querySelector(".cell-calls");
        if (callsTd) callsTd.textContent = contact.calls;
        const nameDisp = masterRow.querySelector(".name-display");
        if (nameDisp) nameDisp.firstChild.textContent = updates.name;
        const wsSelect = masterRow.querySelector(".ws-select");
        if (wsSelect) {
          wsSelect.value = updates.ws;
          wsSelect.dataset.value = updates.ws;
        }
      }

      updateStatsBar(); // Update top stats summary numbers
      tr.classList.add("row-saved");
      setSyncStatus("All changes saved", "saved");
      setTimeout(() => tr.classList.remove("row-saved"), 1200);
      // Lock as "Submitted" until some field in this row changes again
      controls.setSubmitted();
    } else {
      tr.classList.add("row-error");
      setSyncStatus("Save failed", "error");
      showToast(data.message || "Could not save. Please try again.", "error");
      setTimeout(() => tr.classList.remove("row-error"), 1600);
      if (data.code === "AUTH") logout();
      // Errors re-enable the button (no earlier than 3s after the click)
      const enableWait = Math.max(0, 3000 - (Date.now() - startedAt));
      setTimeout(() => controls.updateSubmitState(), enableWait);
    }
  }, popupWait);
}

async function refreshContacts() {
  refreshBtn.classList.add("spinning");
  const data = await api({
    action: "data",
    phone: currentUser.phone,
    sheet: activeCampaign === "Reception" || activeCampaign === "Cultivation" ? "Thursday Calling" : activeCampaign,
    campaignType: activeCampaign === "Cultivation" ? "cultivation" : (activeCampaign === "Festival Promotions" ? "festival" : "calling"),
    skipCache: "true"
  });
  refreshBtn.classList.remove("spinning");

  if (data.status === "success") {
    applyPayload(data);
    try {
      localStorage.setItem(getCacheKey(activeCampaign), JSON.stringify(data));
    } catch (e) {
      console.warn("Storage quota exceeded or error caching:", e);
    }
    renderContacts();
    setSyncStatus("Up to date", "saved");
  } else {
    showToast(data.message || "Refresh failed.", "error");
  }
}

// ============================================================
// Small utilities
// ============================================================

function formatPhone(phone) {
  return phone.length === 10 ? phone.slice(0, 5) + " " + phone.slice(5) : phone;
}

function formatTime12h(ts) {
  if (!ts) return "—";
  const date = new Date(ts);
  if (isNaN(date.getTime())) return ts;
  
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "pm" : "am";
  hours = hours % 12;
  hours = hours ? hours : 12; // 0 should be 12
  return String(hours).padStart(2, "0") + ":" + minutes + " " + ampm;
}

function formatDateAndTime12h(ts) {
  if (!ts) return "—";
  const date = new Date(ts);
  if (isNaN(date.getTime())) return ts;

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const day = date.getDate();
  const month = months[date.getMonth()];
  
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "pm" : "am";
  hours = hours % 12;
  hours = hours ? hours : 12;
  const timeStr = String(hours).padStart(2, "0") + ":" + minutes + " " + ampm;
  
  return day + " " + month + ", " + timeStr;
}

let syncStatusTimer = null;
function setSyncStatus(text, kind) {
  syncStatus.textContent = text;
  syncStatus.className = "sync-status " + (kind || "");
  clearTimeout(syncStatusTimer);
  if (kind === "saved") {
    syncStatusTimer = setTimeout(() => {
      syncStatus.textContent = "";
      syncStatus.className = "sync-status";
    }, 2500);
  }
}

let toastTimer = null;
function showToast(message, kind, duration) {
  toastEl.textContent = message;
  toastEl.className = "toast show " + (kind || "");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toastEl.classList.remove("show"), duration || 3200);
}

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {
      /* PWA caching unavailable (e.g. opened via file://) — app still works */
    });
  }
}

// ============================================================
// Submit master row info (All People Data)
// ============================================================

async function submitMasterRow(contact, tr, controls, btn) {
  const newName = controls.name.value.trim();
  if (!newName) {
    showToast("Name cannot be empty.", "error");
    return;
  }

  const updates = {
    name: newName,
    ws: controls.ws.value,
    cultivatedBy: controls.cultivatedBy.value
  };

  const startedAt = Date.now();
  btn.disabled = true;
  showToast("Saving master info...", "success", 800);
  tr.classList.remove("row-saved", "row-error");
  tr.classList.add("row-saving");
  setSyncStatus("Saving…", "saving");

  const data = await api(null, {
    action: "updateContact",
    requesterPhone: currentUser.phone,
    phone: contact.phone,
    updates: updates
  });

  const popupWait = Math.max(0, 1000 - (Date.now() - startedAt));
  setTimeout(() => {
    tr.classList.remove("row-saving");

    if (data.status === "success") {
      Object.assign(contact, updates);

      // Sync with Section 2 (Thursday Calling) row if it exists
      const callingRow = document.querySelector(`#calling-section tr[data-phone="${contact.phone}"]`);
      if (callingRow) {
        const wsSelect = callingRow.querySelector(".ws-select");
        if (wsSelect) {
          wsSelect.value = updates.ws;
          wsSelect.dataset.value = updates.ws;
        }
        const nameDisp = callingRow.querySelector(".name-display");
        if (nameDisp) nameDisp.firstChild.textContent = updates.name;
      }

      tr.classList.add("row-saved");
      setSyncStatus("Master info saved", "saved");
      setTimeout(() => tr.classList.remove("row-saved"), 1200);
      controls.setSubmitted();
    } else {
      tr.classList.add("row-error");
      setSyncStatus("Save failed", "error");
      showToast(data.message || "Could not save. Please try again.", "error");
      setTimeout(() => tr.classList.remove("row-error"), 1600);
      if (data.code === "AUTH") logout();
      controls.updateSubmitState();
    }
  }, popupWait);
}

// ============================================================
// Auto-Assignment load balancing
// ============================================================

function autoAssignContacts() {
  if (!isAdmin()) return;
  if (!userNames || !userNames.length) {
    showToast("No registered users found to assign calls.", "error");
    return;
  }

  // Find all rows in Section 2 (calling-section)
  const rows = document.querySelectorAll("#contacts-body tr:not(.loading-row)");
  if (!rows.length) {
    showToast("No contacts to assign.", "error");
    return;
  }

  // Step 1: Count current assignments that are pinned via 'Cultivated By' or pre-selected
  const assignmentCounts = {};
  userNames.forEach(name => {
    assignmentCounts[name] = 0;
  });

  const cleanPhone = (num) => String(num || "").replace(/\D/g, "").slice(-10);
  const userPhonesSet = new Set((registeredUserPhones || []).map(p => cleanPhone(p)));
  const rowDataList = [];

  rows.forEach(row => {
    const phone = row.dataset.phone;
    const contact = contacts.find(c => c.phone === phone);
    if (!contact) return;

    // Skip if the contact is one of the callers themselves
    const cleanedPhone = cleanPhone(contact.phone);
    if (userPhonesSet.has(cleanedPhone)) {
      const assignSelect = row.querySelector(".assign-select");
      if (assignSelect) {
        assignSelect.value = "";
        assignSelect.dispatchEvent(new Event("change"));
      }
      return;
    }

    const assignSelect = row.querySelector(".assign-select");
    const cultivatedBy = contact.cultivatedBy || "";

    rowDataList.push({
      row: row,
      contact: contact,
      assignSelect: assignSelect,
      cultivatedBy: cultivatedBy
    });
  });

  // Step 2: Pin the contacts with valid cultivatedBy field
  const unassignedRows = [];
  rowDataList.forEach(item => {
    if (item.cultivatedBy && userNames.includes(item.cultivatedBy)) {
      item.assignSelect.value = item.cultivatedBy;
      assignmentCounts[item.cultivatedBy]++;
      // Trigger change event to mark as dirty and enable Submit
      item.assignSelect.dispatchEvent(new Event("change"));
    } else {
      unassignedRows.push(item);
    }
  });

  // Step 3: Shuffle the remaining unassigned rows randomly
  const shuffledRows = unassignedRows.sort(() => Math.random() - 0.5);

  // Step 4: Greedy load balance remainder, respecting limits
  let unassignedCount = 0;
  shuffledRows.forEach(item => {
    const availableUsers = userNames.filter(name => {
      const limit = userLimits[name] || 999;
      return assignmentCounts[name] < limit;
    });

    if (availableUsers.length === 0) {
      unassignedCount++;
      return; // Leave unassigned
    }

    let minUser = availableUsers[0];
    let minCount = assignmentCounts[minUser];

    availableUsers.forEach(name => {
      if (assignmentCounts[name] < minCount) {
        minUser = name;
        minCount = assignmentCounts[name];
      }
    });

    item.assignSelect.value = minUser;
    assignmentCounts[minUser]++;
    item.assignSelect.dispatchEvent(new Event("change"));
  });

  if (unassignedCount > 0) {
    showToast(`Auto-assigned! ${unassignedCount} calls left unassigned due to user limits.`, "warning");
  } else {
    showToast("Auto-assigned successfully! Click Save Assignments to apply.", "success");
  }
}

// ============================================================
// Show Call History Modal
// ============================================================

async function showCallHistoryModal(contact) {
  historyModalTitle.textContent = "Call History";
  historyTh2.textContent = "Date";
  historyTh3.textContent = "Status";
  historyTh4.style.display = "";

  historyContactInfo.textContent = `Contact: ${contact.name} (${formatPhone(contact.phone)})`;
  historyBody.innerHTML = `
    <tr>
      <td colspan="4" class="no-history">
        <span class="btn-spinner dark"></span> Loading history logs...
      </td>
    </tr>
  `;
  historyModal.classList.add("active");

  const data = await api({ action: "history", phone: contact.phone });
  if (data.status === "success" && Array.isArray(data.history)) {
    if (data.history.length === 0) {
      historyBody.innerHTML = `<tr><td colspan="4" class="no-history">No calls logged yet.</td></tr>`;
    } else {
      historyBody.innerHTML = "";
      data.history.forEach(log => {
        const tr = document.createElement("tr");

        const timeTd = document.createElement("td");
        timeTd.textContent = formatTime12h(log.time);

        const roundTd = document.createElement("td");
        roundTd.textContent = log.round || "—";

        const statusTd = document.createElement("td");
        statusTd.textContent = log.status || "—";

        const byTd = document.createElement("td");
        byTd.textContent = log.by || "—";

        tr.appendChild(timeTd);
        tr.appendChild(roundTd);
        tr.appendChild(statusTd);
        tr.appendChild(byTd);
        historyBody.appendChild(tr);
      });
    }
  } else {
    historyBody.innerHTML = `<tr><td colspan="4" class="no-history error-text">Failed to load history.</td></tr>`;
  }
}

// ============================================================
// Show Session History Modal
// ============================================================

async function showSessionHistoryModal(contact) {
  historyModalTitle.textContent = "Session History";
  historyTh2.textContent = "Session Name";
  historyTh3.textContent = "Attended";
  historyTh4.style.display = "none";

  historyContactInfo.textContent = `Contact: ${contact.name} (${formatPhone(contact.phone)})`;
  historyBody.innerHTML = `
    <tr>
      <td colspan="3" class="no-history">
        <span class="btn-spinner dark"></span> Loading session history...
      </td>
    </tr>
  `;
  historyModal.classList.add("active");

  const data = await api({ action: "sessions", phone: contact.phone });
  if (data.status === "success" && Array.isArray(data.sessions)) {
    if (data.sessions.length === 0) {
      historyBody.innerHTML = `<tr><td colspan="3" class="no-history">No session attendance logged yet.</td></tr>`;
    } else {
      historyBody.innerHTML = "";
      data.sessions.forEach(log => {
        const tr = document.createElement("tr");

        const timeTd = document.createElement("td");
        timeTd.textContent = formatDateAndTime12h(log.time);

        const sessionTd = document.createElement("td");
        sessionTd.textContent = log.session || "—";

        const attendedTd = document.createElement("td");
        attendedTd.textContent = log.attended || "—";

        tr.appendChild(timeTd);
        tr.appendChild(sessionTd);
        tr.appendChild(attendedTd);
        historyBody.appendChild(tr);
      });
    }
  } else {
    historyBody.innerHTML = `<tr><td colspan="3" class="no-history error-text">Failed to load session history.</td></tr>`;
  }
}

// ============================================================
// Bulk Save Changes for Admin
// ============================================================

async function submitBulkChanges() {
  if (!isAdmin()) return;
  if (dirtyContacts.size === 0) return;

  const updatesArray = [];
  dirtyContacts.forEach(phone => {
    const row = document.querySelector(`#contacts-body tr[data-phone="${phone}"]`);
    if (row && row.controls) {
      const controls = row.controls;
      const contact = contacts.find(c => c.phone === phone);
      if (contact) {
        const updatesObj = {
          name: controls.name.value,
          ws: controls.ws.value,
          status: controls.status ? controls.status.value : contact.status,
          assignedTo: controls.assignedTo ? controls.assignedTo.value : contact.assignedTo
        };
        if (controls.eventSelect) {
          updatesObj.event = controls.eventSelect.value;
        }
        updatesArray.push({
          phone: phone,
          updates: updatesObj
        });
      }
    }
  });

  if (updatesArray.length === 0) return;

  saveAllBtn.disabled = true;
  saveAllBtn.textContent = "Saving…";
  showToast("Saving all assignments...", "success", 1000);
  setSyncStatus("Saving assignments…", "saving");

  const data = await api(null, {
    action: "updateContactsBulk",
    requesterPhone: currentUser.phone,
    updates: updatesArray
  });

  saveAllBtn.textContent = "Save Assignments";

  if (data.status === "success") {
    // Apply updates locally and animate rows
    updatesArray.forEach(item => {
      const contact = contacts.find(c => c.phone === item.phone);
      if (contact) {
        Object.assign(contact, item.updates);

        // Fetch fresh calls count if returned
        const fresh = (data.contacts || []).find(c => c.phone === item.phone);
        if (fresh) {
          contact.calls = fresh.calls;
          contact.sessions = fresh.sessions;
        }

        // Animate row in calling section
        const callingRow = document.querySelector(`#contacts-body tr[data-phone="${item.phone}"]`);
        if (callingRow && callingRow.controls) {
          callingRow.controls.callsTd.textContent = contact.calls;
          // Sessions formula was recalculated
          const sessionsCell = callingRow.querySelector(".cell-sessions");
          if (sessionsCell) sessionsCell.textContent = contact.sessions;

          callingRow.classList.add("row-saved");
          setTimeout(() => callingRow.classList.remove("row-saved"), 1200);
          callingRow.controls.setSubmitted();
        }

        // Animate master row if visible
        const masterRow = document.querySelector(`#master-contacts-body tr[data-phone="${item.phone}"]`);
        if (masterRow && masterRow.controls) {
          masterRow.controls.callsTd.textContent = contact.calls;
          const sessionsCell = masterRow.querySelector(".cell-sessions");
          if (sessionsCell) sessionsCell.textContent = contact.sessions;

          const nameDisp = masterRow.querySelector(".name-display");
          if (nameDisp) nameDisp.firstChild.textContent = contact.name;

          const wsSelect = masterRow.querySelector(".ws-select");
          if (wsSelect) {
            wsSelect.value = contact.ws;
            wsSelect.dataset.value = contact.ws;
          }

          const cultivatedSelect = masterRow.querySelector(".assign-select");
          if (cultivatedSelect) {
            cultivatedSelect.value = contact.cultivatedBy || "";
          }

          masterRow.classList.add("row-saved");
          setTimeout(() => masterRow.classList.remove("row-saved"), 1200);
          masterRow.controls.setSubmitted();
        }
      }
    });

    dirtyContacts.clear();
    saveAllBtn.disabled = true;
    updateStatsBar();
    setSyncStatus("All changes saved", "saved");
    showToast("All assignments saved successfully! 🎉", "success");
  } else {
    saveAllBtn.disabled = false;
    setSyncStatus("Save failed", "error");
    showToast(data.message || "Failed to save assignments. Please try again.", "error");
    if (data.code === "AUTH") logout();
  }
}

// ============================================================
// Pull to Refresh Implementation
// ============================================================

let startY = 0;
let isPulling = false;
let pullIndicator = null;

document.addEventListener("DOMContentLoaded", () => {
  pullIndicator = document.createElement("div");
  pullIndicator.id = "pull-to-refresh";
  pullIndicator.className = "pull-to-refresh";
  pullIndicator.innerHTML = '<span class="btn-spinner dark"></span> Release to refresh…';
  document.body.prepend(pullIndicator);

  window.addEventListener("touchstart", (e) => {
    if (window.scrollY === 0 && e.touches.length === 1 && currentUser) {
      startY = e.touches[0].pageY;
      isPulling = true;
    }
  }, { passive: true });

  window.addEventListener("touchmove", (e) => {
    if (!isPulling) return;
    const currentY = e.touches[0].pageY;
    const diffY = currentY - startY;

    if (diffY > 0) {
      document.body.classList.add("pulling");
      const appViewY = Math.min(diffY * 0.4, 50);
      const indicatorY = appViewY - 50;

      pullIndicator.style.transform = `translateY(${indicatorY}px)`;
      pullIndicator.style.opacity = String(appViewY / 50);
      
      const appViewEl = document.getElementById("app-view");
      if (appViewEl) {
        appViewEl.style.transform = `translateY(${appViewY}px)`;
      }
    } else {
      isPulling = false;
      resetPull();
    }
  }, { passive: true });

  window.addEventListener("touchend", async (e) => {
    if (!isPulling) return;
    const currentY = e.changedTouches[0].pageY;
    const diffY = currentY - startY;
    isPulling = false;
    document.body.classList.remove("pulling");

    if (diffY > 120) {
      pullIndicator.innerHTML = '<span class="btn-spinner dark"></span> Refreshing data…';
      pullIndicator.style.transform = "translateY(0px)";
      pullIndicator.style.opacity = "1";
      const appViewEl = document.getElementById("app-view");
      if (appViewEl) appViewEl.style.transform = "translateY(50px)";
      
      await refreshContacts();
      
      pullIndicator.innerHTML = '<span class="btn-spinner dark"></span> Release to refresh…';
      resetPull();
    } else {
      resetPull();
    }
  });

  function resetPull() {
    pullIndicator.style.transform = "";
    pullIndicator.style.opacity = "";
    const appViewEl = document.getElementById("app-view");
    if (appViewEl) appViewEl.style.transform = "";
  }
});

// ============================================================
// Reception — Debounced phone search + Mark Attendance
// ============================================================

let receptionDebounce = null;
let lastSearchedContact = null;

function setupReception() {
  receptionSearch.addEventListener("input", () => {
    receptionSearch.value = receptionSearch.value.replace(/\D/g, "").slice(0, 10);
    
    receptionResult.classList.add("hidden");
    receptionNotFound.classList.add("hidden");
    lastSearchedContact = null;

    clearTimeout(receptionDebounce);
    const phone = receptionSearch.value.trim();
    if (phone.length < 10) return;

    receptionDebounce = setTimeout(() => searchReceptionContact(phone), 1500);
  });

  markAttendanceBtn.addEventListener("click", markAttendance);
}

async function searchReceptionContact(phone) {
  receptionLoader.classList.remove("hidden");
  receptionResult.classList.add("hidden");
  receptionNotFound.classList.add("hidden");

  const data = await api({ action: "searchContact", phone: phone });
  receptionLoader.classList.add("hidden");

  if (data.status === "success" && data.found) {
    lastSearchedContact = data.contact;
    receptionNameEl.textContent = data.contact.name;
    receptionPhoneEl.textContent = data.contact.phone;
    receptionSessionsEl.textContent = "Sessions attended: " + (data.contact.sessions || 0);
    receptionResult.classList.remove("hidden");
  } else {
    receptionNotFound.classList.remove("hidden");
    lastSearchedContact = null;
  }
}

async function markAttendance() {
  if (!lastSearchedContact) {
    showToast("No contact found. Search a number first.", "error");
    return;
  }

  markAttendanceBtn.disabled = true;
  markAttendanceBtn.textContent = "Marking…";

  const sessionName = receptionSessionName.value.trim() || "General Session";

  const data = await api(null, {
    action: "markAttendance",
    requesterPhone: currentUser.phone,
    phone: lastSearchedContact.phone,
    name: lastSearchedContact.name,
    sessionName: sessionName
  });

  markAttendanceBtn.disabled = false;
  markAttendanceBtn.textContent = "🙏 Mark Attendance";

  if (data.status === "success") {
    showToast("Attendance marked for " + lastSearchedContact.name + " 🙏", "success");
    saveReceptionAttendanceLocally(lastSearchedContact.name, lastSearchedContact.phone, sessionName);
    // Re-fetch to update session count
    await searchReceptionContact(lastSearchedContact.phone);
    receptionSessionName.value = "";
  } else {
    showToast(data.message || "Failed to mark attendance.", "error");
  }
}

// ============================================================
// Cultivation — Load contacts where user is the cultivator
// ============================================================

async function loadCultivationContacts() {
  cultivationBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:24px;"><span class="btn-spinner dark"></span> Loading…</td></tr>';
  cultivationTable.classList.remove("hidden");
  cultivationEmpty.classList.add("hidden");

  // Load cultivation contacts from both sheets
  const allContacts = [];

  const data1 = await api({ action: "data", phone: currentUser.phone, sheet: "Thursday Calling", campaignType: "cultivation", skipCache: "true" });
  if (data1.status === "success" && data1.contacts) {
    data1.contacts.forEach(c => {
      allContacts.push({ ...c, campaign: "Thursday Calling" });
    });
  }

  const data2 = await api({ action: "data", phone: currentUser.phone, sheet: "Festival Promotions", campaignType: "cultivation", skipCache: "true" });
  if (data2.status === "success" && data2.contacts) {
    data2.contacts.forEach(c => {
      allContacts.push({ ...c, campaign: "Festival Promotions" });
    });
  }

  // Setup table header dynamically to match Thursday Calling
  const cultivationHead = cultivationTable.querySelector("thead tr");
  if (cultivationHead) {
    cultivationHead.innerHTML = "";
    const headers = ["Name", "Phone Number", "W/S", "Sessions", "No. of Calls", "Calling Status", ""];
    headers.forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      cultivationHead.appendChild(th);
    });
  }

  cultivationBody.innerHTML = "";

  if (allContacts.length === 0) {
    cultivationTable.classList.add("hidden");
    cultivationEmpty.classList.remove("hidden");
    return;
  }

  cultivationEmpty.classList.add("hidden");
  cultivationTable.classList.remove("hidden");

  allContacts.forEach(c => {
    const row = createContactRow(c, false, true);
    cultivationBody.appendChild(row);
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str || "";
  return div.innerHTML;
}

// ============================================================
// Add Person Modal Logic (Admin & Reception)
// ============================================================

let addPersonContext = "admin"; // "admin" or "reception"

function openAddPersonModal(context) {
  addPersonContext = context;
  addPersonNameInput.value = "";
  addPersonPhoneInput.value = "";
  addPersonError.textContent = "";

  if (context === "admin") {
    addPersonTitle.textContent = "Add Person to Thursday Calling";
    addPersonSubmitBtn.textContent = "Add Person";
  } else {
    addPersonTitle.textContent = "New Registration & Attendance";
    addPersonSubmitBtn.textContent = "Register & Mark Attendance";
  }

  addPersonModal.classList.add("active");
  setTimeout(() => addPersonNameInput.focus(), 60);
}

function closeAddPersonModal() {
  addPersonModal.classList.remove("active");
}

async function submitAddPersonForm() {
  const name = addPersonNameInput.value.trim();
  const phone = addPersonPhoneInput.value.trim().replace(/\D/g, "");

  if (!name || phone.length !== 10) {
    addPersonError.textContent = "Please enter a valid Name and 10-digit Phone Number.";
    return;
  }

  addPersonSubmitBtn.disabled = true;
  addPersonSubmitBtn.textContent = "Saving…";
  addPersonError.textContent = "";

  if (addPersonContext === "admin") {
    // Admin directly adds contact to sheet
    const data = await api(null, {
      action: "addContact",
      requesterPhone: currentUser.phone,
      phone: phone,
      name: name,
      sheet: activeCampaign || "Thursday Calling",
      event: (activeCampaign === "Festival Promotions" && typeof adminFestivalFilter !== "undefined" && adminFestivalFilter && adminFestivalFilter.value !== "ALL") ? adminFestivalFilter.value : ""
    });

    addPersonSubmitBtn.disabled = false;
    addPersonSubmitBtn.textContent = "Add Person";

    if (data.status === "success") {
      showToast("Added " + name + " successfully!", "success");
      closeAddPersonModal();
      applyPayload(data);
      renderContacts();
      // Update cache
      try {
        localStorage.setItem(getCacheKey(activeCampaign), JSON.stringify(data));
      } catch (e) {
        console.warn("Storage quota exceeded or error caching:", e);
      }
    } else {
      addPersonError.textContent = data.message || "Failed to add person.";
    }
  } else {
    // Reception adds contact AND marks attendance automatically
    const sessionName = receptionSessionName.value.trim() || "General Session";
    const data = await api(null, {
      action: "markAttendance",
      requesterPhone: currentUser.phone,
      phone: phone,
      name: name,
      sessionName: sessionName
    });

    addPersonSubmitBtn.disabled = false;
    addPersonSubmitBtn.textContent = "Register & Mark Attendance";

    if (data.status === "success") {
      showToast("Registered & marked attendance for " + name + "!", "success");
      closeAddPersonModal();
      
      // Save to Reception local storage marked attendance list
      saveReceptionAttendanceLocally(name, phone, sessionName);
      
      // Trigger search automatically to show registered details
      receptionSearch.value = phone;
      searchReceptionContact(phone);
    } else {
      addPersonError.textContent = data.message || "Failed to register person.";
    }
  }
}

// ============================================================
// Reception Local Storage Attendance List
// ============================================================

const RECEPTION_MARKED_KEY = "reception_marked_attendance_today";

function saveReceptionAttendanceLocally(name, phone, sessionName) {
  try {
    const todayStr = new Date().toDateString();
    const current = localStorage.getItem(RECEPTION_MARKED_KEY);
    let list = [];
    if (current) {
      const parsed = JSON.parse(current);
      if (parsed && parsed.date === todayStr && Array.isArray(parsed.data)) {
        list = parsed.data;
      }
    }

    const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    list.unshift({ name, phone, time: timeStr, session: sessionName });

    localStorage.setItem(RECEPTION_MARKED_KEY, JSON.stringify({
      date: todayStr,
      data: list
    }));

    renderReceptionAttendanceList();
  } catch (e) {
    console.error("Error saving local attendance:", e);
  }
}

function renderReceptionAttendanceList() {
  try {
    const todayStr = new Date().toDateString();
    const current = localStorage.getItem(RECEPTION_MARKED_KEY);
    let list = [];
    if (current) {
      const parsed = JSON.parse(current);
      if (parsed && parsed.date === todayStr && Array.isArray(parsed.data)) {
        list = parsed.data;
      }
    }

    if (!receptionAttendanceBody) return;
    receptionAttendanceBody.innerHTML = "";

    if (list.length === 0) {
      receptionAttendanceBody.innerHTML = `
        <tr>
          <td colspan="4" style="text-align:center; color:var(--text-muted); padding:16px;">No attendance marked today.</td>
        </tr>
      `;
      return;
    }

    list.forEach(item => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(item.name)}</td>
        <td class="phone-cell">${escapeHtml(item.phone)}</td>
        <td>${escapeHtml(item.time)}</td>
        <td>${escapeHtml(item.session)}</td>
      `;
      receptionAttendanceBody.appendChild(tr);
    });
  } catch (e) {
    console.error("Error rendering local attendance:", e);
  }
}

// ============================================================
// Stale-While-Revalidate Local Cache Loader
// ============================================================

const CACHE_CONTACTS_KEY = "fnrg_cached_contacts";

function getCacheKey(campaign) {
  const phone = currentUser ? currentUser.phone : "";
  return CACHE_CONTACTS_KEY + "_" + phone + "_" + campaign;
}

function loadCachedContacts() {
  const cached = localStorage.getItem(getCacheKey(activeCampaign));
  if (cached) {
    try {
      const payload = JSON.parse(cached);
      if (payload && payload.contacts) {
        applyPayload(payload);
        renderContacts();
      }
    } catch (e) {
      console.warn("Error parsing cached contacts:", e);
    }
  }
}

// ============================================================
// Assigned Calls Summary Modal (Admin Only)
// ============================================================

function openAssignedDataModal() {
  if (!assignedDataBody) return;
  assignedDataBody.innerHTML = "";

  // Calculate counts
  const counts = {};
  userNames.forEach(name => {
    counts[name] = 0;
  });

  contacts.forEach(contact => {
    const assigned = contact.assignedTo || "";
    if (assigned && userNames.includes(assigned)) {
      counts[assigned] = (counts[assigned] || 0) + 1;
    }
  });

  userNames.forEach(name => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(name)}</td>
      <td style="font-weight:600; text-align:center;">${counts[name]}</td>
    `;
    assignedDataBody.appendChild(tr);
  });

  assignedDataModal.classList.add("active");
}

// ============================================================
// Excel Import/Export (Admin Only)
// ============================================================

let importedRows = []; // temporary parsed rows
let excelHeaders = []; // headers of the uploaded file

function exportToExcel() {
  if (contacts.length === 0) {
    showToast("No data available to export.", "error");
    return;
  }

  // Map contacts to clear row structure
  const data = contacts.map(c => ({
    "Name": c.name,
    "Phone Number": c.phone,
    "W/S": c.ws,
    "Sessions": c.sessions,
    "No. of Calls": c.calls,
    "Cultivated By": c.cultivatedBy || "",
    "Assigned To": c.assignedTo || ""
  }));

  try {
    // Create worksheet
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeCampaign || "Contacts");

    // Save XLSX file
    XLSX.writeFile(wb, (activeCampaign || "contacts").replace(/\s+/g, "_") + "_export.xlsx");
    showToast("Exported successfully! Check downloads.", "success");
  } catch (err) {
    showToast("Failed to export Excel file: " + err.message, "error");
  }
}

function handleFileImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to JSON array of arrays (including headers)
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (rows.length === 0) {
        showToast("The uploaded file is empty.", "error");
        return;
      }

      excelHeaders = rows[0]; // first row is headers
      importedRows = XLSX.utils.sheet_to_json(worksheet); // full objects

      // Populate column mapping selectors
      populateMappingSelectors(excelHeaders);
      
      // Show mapping modal
      importMappingModal.classList.add("active");
    } catch(err) {
      showToast("Failed to read Excel file: " + err.message, "error");
    }
  };
  reader.readAsBinaryString(file);
}

function populateMappingSelectors(headers) {
  mappingNameSelect.innerHTML = "";
  mappingPhoneSelect.innerHTML = "";

  headers.forEach(h => {
    const optName = document.createElement("option");
    optName.value = h;
    optName.textContent = h;
    // Match "name"
    if (h.toLowerCase().includes("name") || h.toLowerCase() === "name") {
      optName.selected = true;
    }
    mappingNameSelect.appendChild(optName);

    const optPhone = document.createElement("option");
    optPhone.value = h;
    optPhone.textContent = h;
    // Match "phone" or "number" or "mobile"
    if (h.toLowerCase().includes("phone") || h.toLowerCase().includes("number") || h.toLowerCase().includes("mobile")) {
      optPhone.selected = true;
    }
    mappingPhoneSelect.appendChild(optPhone);
  });
}

async function submitImportData() {
  const nameCol = mappingNameSelect.value;
  const phoneCol = mappingPhoneSelect.value;

  if (!nameCol || !phoneCol) {
    importMappingError.textContent = "Please map both Name and Phone Number columns.";
    return;
  }

  // Parse contacts based on mappings
  const parsedContacts = [];
  importedRows.forEach(row => {
    const nameVal = String(row[nameCol] || "").trim();
    const phoneVal = String(row[phoneCol] || "").trim().replace(/\D/g, "");
    if (nameVal && phoneVal.length === 10) {
      parsedContacts.push({ name: nameVal, phone: phoneVal });
    }
  });

  if (parsedContacts.length === 0) {
    importMappingError.textContent = "No valid contacts found. Phone numbers must be 10 digits.";
    return;
  }

  importMappingSubmit.disabled = true;
  importMappingSubmit.textContent = "Importing…";
  importMappingError.textContent = "";

  const response = await api(null, {
    action: "importContacts",
    requesterPhone: currentUser.phone,
    sheet: activeCampaign || "Thursday Calling",
    contacts: parsedContacts,
    event: (activeCampaign === "Festival Promotions" && typeof adminFestivalFilter !== "undefined" && adminFestivalFilter && adminFestivalFilter.value !== "ALL") ? adminFestivalFilter.value : ""
  });

  importMappingSubmit.disabled = false;
  importMappingSubmit.textContent = "Import Data";

  if (response.status === "success") {
    showToast(`Successfully imported ${response.importedCount} contacts (${response.skippedCount} skipped/duplicates).`, "success", 4000);
    importMappingModal.classList.remove("active");
    applyPayload(response);
    renderContacts();
    // Update local cache
    try {
      localStorage.setItem(getCacheKey(activeCampaign), JSON.stringify(response));
    } catch (e) {
      console.warn("Storage quota exceeded or error caching:", e);
    }
  } else {
    importMappingError.textContent = response.message || "Failed to import contacts.";
  }
}

// ============================================================
// Festival Promotions (Admin Only)
// ============================================================

let loadedFestivals = [];

async function loadAdminFestivalData() {
  const tableBody = adminFestivalContactsBody;
  if (!tableBody) return;

  tableBody.innerHTML = `
    <tr class="loading-row">
      <td colspan="9">
        <span class="btn-spinner dark"></span> Loading festival list…
      </td>
    </tr>
  `;

  const response = await api({
    action: "data",
    phone: currentUser.phone,
    sheet: "Festival Promotions",
    campaignType: "festival"
  });

  if (response.status !== "success") {
    tableBody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align:center; color:#ef4444; padding:20px;">
          Failed to load festival promotions data: ${escapeHtml(response.message)}
        </td>
      </tr>
    `;
    return;
  }

  loadedFestivals = response.festivals || [];
  populateFestivalFilter();
  populateAutoAssignUsers(response.userNames || [], response.autoAssignUsers || [], response.autoAssignUserFestivals || {});

  contacts = response.contacts || [];
  userNames = response.userNames || [];
  renderAdminFestivalContacts();
}

function populateFestivalFilter() {
  if (!adminFestivalFilter) return;
  adminFestivalFilter.innerHTML = `<option value="ALL">All Festivals</option>`;
  loadedFestivals.forEach(f => {
    const o = document.createElement("option");
    o.value = f;
    o.textContent = f;
    adminFestivalFilter.appendChild(o);
  });
  adminFestivalFilter.removeEventListener("change", renderAdminFestivalContacts);
  adminFestivalFilter.addEventListener("change", renderAdminFestivalContacts);
}

function populateAutoAssignUsers(allUsers, selectedUsers, selectedUserFestivals) {
  if (!autoAssignUsersList) return;
  autoAssignUsersList.innerHTML = "";
  selectedUserFestivals = selectedUserFestivals || {};

  allUsers.forEach(u => {
    // Row wrapper container
    const rowDiv = document.createElement("div");
    rowDiv.style.display = "flex";
    rowDiv.style.alignItems = "center";
    rowDiv.style.justifyContent = "space-between";
    rowDiv.style.gap = "15px";
    rowDiv.style.padding = "6px 0";
    rowDiv.style.borderBottom = "1px dashed #f1ebd9";

    // Left checkbox and name label
    const label = document.createElement("label");
    label.style.display = "flex";
    label.style.alignItems = "center";
    label.style.gap = "8px";
    label.style.fontSize = "13.5px";
    label.style.fontWeight = "600";
    label.style.cursor = "pointer";
    label.style.margin = "0";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = u;
    if (selectedUsers.includes(u)) {
      cb.checked = true;
    }
    cb.addEventListener("change", async () => {
      cb.disabled = true;
      const response = await api(null, {
        action: "updateAutoAssignUser",
        requesterPhone: currentUser.phone,
        username: u,
        enabled: cb.checked
      });
      cb.disabled = false;
      if (response.status === "success") {
        showToast(cb.checked ? `${u} enabled for auto assignment.` : `${u} disabled for auto assignment.`, "success");
        if (response.contacts) contacts = response.contacts;
        if (response.userNames) userNames = response.userNames;
        const freshSelected = response.autoAssignUsers || [];
        populateAutoAssignUsers(userNames, freshSelected, response.autoAssignUserFestivals || {});
        renderAdminFestivalContacts();
      } else {
        showToast("Error updating assignment settings: " + response.message, "error");
        cb.checked = !cb.checked; // revert
      }
    });

    label.appendChild(cb);
    label.appendChild(document.createTextNode(u));
    rowDiv.appendChild(label);

    // Right select dropdown for festival assignment
    const select = document.createElement("select");
    select.className = "status-select";
    select.style.padding = "4px 8px";
    select.style.fontSize = "12px";
    select.style.minWidth = "130px";
    select.style.margin = "0";

    const defaultOpt = document.createElement("option");
    defaultOpt.value = "";
    defaultOpt.textContent = "All Festivals";
    select.appendChild(defaultOpt);

    loadedFestivals.forEach(fest => {
      const opt = document.createElement("option");
      opt.value = fest;
      opt.textContent = fest;
      if (selectedUserFestivals[u] === fest) {
        opt.selected = true;
      }
      select.appendChild(opt);
    });

    select.addEventListener("change", async () => {
      select.disabled = true;
      const response = await api(null, {
        action: "updateUserFestival",
        requesterPhone: currentUser.phone,
        username: u,
        festival: select.value
      });
      select.disabled = false;
      if (response.status === "success") {
        showToast(`Festival for ${u} updated to: ${select.value || "All Festivals"}`, "success");
        if (response.contacts) contacts = response.contacts;
        if (response.userNames) userNames = response.userNames;
        const freshSelected = response.autoAssignUsers || [];
        populateAutoAssignUsers(userNames, freshSelected, response.autoAssignUserFestivals || {});
        renderAdminFestivalContacts();
      } else {
        showToast("Error updating user festival assignment: " + response.message, "error");
      }
    });

    rowDiv.appendChild(select);
    autoAssignUsersList.appendChild(rowDiv);
  });
}

function renderAdminFestivalContacts() {
  const body = adminFestivalContactsBody;
  if (!body) return;
  body.innerHTML = "";

  const filterVal = adminFestivalFilter.value;
  const filtered = contacts.filter(c => {
    if (filterVal === "ALL") return true;
    return c.event === filterVal;
  });

  if (filtered.length === 0) {
    body.innerHTML = `
      <tr>
        <td colspan="9" style="text-align:center; color:var(--text-muted); padding:20px;">
          No contacts match this festival.
        </td>
      </tr>
    `;
    return;
  }

  filtered.forEach(contact => {
    const tr = document.createElement("tr");
    tr.dataset.phone = contact.phone;
    if (contact.assignedTo) {
      tr.classList.add("row-assigned");
    }

    // 1. Name
    const nameTd = document.createElement("td");
    nameTd.className = "cell-name";
    nameTd.textContent = contact.name;
    tr.appendChild(nameTd);

    // 2. Phone
    const phoneTd = document.createElement("td");
    phoneTd.textContent = contact.phone;
    tr.appendChild(phoneTd);

    // 3. WS
    const wsTd = document.createElement("td");
    const wsSelect = document.createElement("select");
    wsSelect.className = "status-select";
    wsSelect.style.width = "75px";
    wsSelect.style.padding = "4px 8px";
    ["NA", "W", "S"].forEach(opt => {
      const o = document.createElement("option");
      o.value = opt;
      o.textContent = opt;
      if (opt === contact.ws) o.selected = true;
      wsSelect.appendChild(o);
    });
    wsSelect.addEventListener("change", () => {
      contact.ws = wsSelect.value;
      tr.classList.add("row-dirty");
    });
    wsTd.appendChild(wsSelect);
    tr.appendChild(wsTd);

    // 4. Sessions
    const sessionsTd = document.createElement("td");
    sessionsTd.style.textAlign = "center";
    sessionsTd.textContent = contact.sessions;
    tr.appendChild(sessionsTd);

    // 5. Calls
    const callsTd = document.createElement("td");
    callsTd.style.textAlign = "center";
    callsTd.textContent = contact.calls;
    tr.appendChild(callsTd);

    // 6. Assigned To
    const assignTd = document.createElement("td");
    const assignSelect = document.createElement("select");
    assignSelect.className = "assign-select";
    assignSelect.style.width = "130px";
    const unassigned = document.createElement("option");
    unassigned.value = "";
    unassigned.textContent = "— Unassigned —";
    assignSelect.appendChild(unassigned);
    userNames.forEach(u => {
      const o = document.createElement("option");
      o.value = u;
      o.textContent = u;
      if (u === contact.assignedTo) o.selected = true;
      assignSelect.appendChild(o);
    });
    assignSelect.addEventListener("change", () => {
      contact.assignedTo = assignSelect.value;
      if (assignSelect.value !== "") {
        tr.classList.add("row-assigned");
      } else {
        tr.classList.remove("row-assigned");
      }
      tr.classList.add("row-dirty");
    });
    assignTd.appendChild(assignSelect);
    tr.appendChild(assignTd);

    // 7. Event
    const eventTd = document.createElement("td");
    const eventSelect = document.createElement("select");
    eventSelect.className = "status-select";
    eventSelect.style.minWidth = "120px";
    const emptyOpt = document.createElement("option");
    emptyOpt.value = "";
    emptyOpt.textContent = "— Select Event —";
    eventSelect.appendChild(emptyOpt);
    loadedFestivals.forEach(f => {
      const o = document.createElement("option");
      o.value = f;
      o.textContent = f;
      if (f === contact.event) o.selected = true;
      eventSelect.appendChild(o);
    });
    if (contact.event && !loadedFestivals.includes(contact.event)) {
      const cust = document.createElement("option");
      cust.value = contact.event;
      cust.textContent = contact.event;
      cust.selected = true;
      eventSelect.appendChild(cust);
    }
    eventSelect.addEventListener("change", () => {
      contact.event = eventSelect.value;
      tr.classList.add("row-dirty");
    });
    eventTd.appendChild(eventSelect);
    tr.appendChild(eventTd);

    // 8. Calling Status
    const statusTd = document.createElement("td");
    statusTd.className = "cell-status";
    const statusSelect = document.createElement("select");
    statusSelect.className = "status-select";
    statusSelect.setAttribute("aria-label", contact.name + " calling status");
    statusOptions.forEach((opt) => {
      const option = document.createElement("option");
      option.value = opt;
      option.textContent = opt;
      statusSelect.appendChild(option);
    });
    if (contact.status && !statusOptions.includes(contact.status)) {
      const custom = document.createElement("option");
      custom.value = contact.status;
      custom.textContent = contact.status;
      statusSelect.appendChild(custom);
    }
    statusSelect.value = statusOptions.includes(contact.status) || contact.status ? contact.status : STATUS_DEFAULT;
    if (!statusSelect.value) statusSelect.value = STATUS_DEFAULT;
    statusSelect.dataset.value = statusSelect.value;
    statusSelect.dataset.prev = statusSelect.value;

    statusSelect.addEventListener("change", () => {
      if (statusSelect.value === STATUS_OTHERS) {
        openOthersModal({
          onConfirm: (text) => {
            let custom = statusSelect.querySelector("option[data-custom]");
            if (!custom) {
              custom = document.createElement("option");
              custom.setAttribute("data-custom", "1");
              statusSelect.appendChild(custom);
            }
            custom.value = text;
            custom.textContent = text;
            statusSelect.value = text;
            statusSelect.dataset.value = text;
            statusSelect.dataset.prev = text;
            contact.status = text;
            tr.classList.add("row-dirty");
          },
          onCancel: () => {
            statusSelect.value = statusSelect.dataset.prev;
            statusSelect.dataset.value = statusSelect.value;
          }
        });
      } else {
        statusSelect.dataset.value = statusSelect.value;
        statusSelect.dataset.prev = statusSelect.value;
        contact.status = statusSelect.value;
        tr.classList.add("row-dirty");
      }
    });
    statusTd.appendChild(statusSelect);
    tr.appendChild(statusTd);

    // 9. Actions
    const actionTd = document.createElement("td");
    const saveRowBtn = document.createElement("button");
    saveRowBtn.className = "btn-secondary";
    saveRowBtn.style.padding = "4px 10px";
    saveRowBtn.style.borderRadius = "6px";
    saveRowBtn.style.fontSize = "12px";
    saveRowBtn.textContent = "Submit";
    saveRowBtn.addEventListener("click", async () => {
      saveRowBtn.disabled = true;
      saveRowBtn.textContent = "Saving…";
      const res = await api(null, {
        action: "updateContact",
        requesterPhone: currentUser.phone,
        sheet: "Festival Promotions",
        phone: contact.phone,
        updates: {
          ws: wsSelect.value,
          assignedTo: assignSelect.value,
          event: eventSelect.value,
          status: statusSelect.value
        }
      });
      saveRowBtn.disabled = false;
      saveRowBtn.textContent = "Submit";
      if (res.status === "success") {
        showToast("Contact updated successfully.", "success");
        tr.classList.remove("row-dirty");
        const fresh = res.contacts.find(fc => fc.phone === contact.phone);
        if (fresh) {
          contact.ws = fresh.ws;
          contact.assignedTo = fresh.assignedTo;
          contact.event = fresh.event;
          contact.sessions = fresh.sessions;
          contact.calls = fresh.calls;
          contact.status = fresh.status;
        }
        renderAdminFestivalContacts();
      } else {
        showToast("Error updating contact: " + res.message, "error");
      }
    });
    actionTd.appendChild(saveRowBtn);
    tr.appendChild(actionTd);

    body.appendChild(tr);
  });
}

let backgroundPollInterval = null;

function startBackgroundPolling() {
  if (backgroundPollInterval) clearInterval(backgroundPollInterval);
  
  backgroundPollInterval = setInterval(async () => {
    if (!currentUser || document.hidden || activeCampaign === "") return;
    
    // Silent API refresh
    const response = await api({
      action: "data",
      phone: currentUser.phone,
      sheet: activeCampaign === "Reception" || activeCampaign === "Cultivation" ? "Thursday Calling" : activeCampaign,
      campaignType: activeCampaign === "Cultivation" ? "cultivation" : (activeCampaign === "Festival Promotions" ? "festival" : "calling"),
      skipCache: true
    });

    if (response && response.status === "success") {
      try {
        localStorage.setItem(getCacheKey(activeCampaign), JSON.stringify(response));
      } catch(e) {}
      
      // Update loaded data silently if there are no unsaved input changes
      if (dirtyContacts.size === 0) {
        contacts = response.contacts || [];
        userNames = response.userNames || [];
        loadedFestivals = response.festivals || [];
        
        if (activeCampaign === "Festival Promotions" && isAdmin()) {
          const currentFilter = adminFestivalFilter ? adminFestivalFilter.value : "ALL";
          populateFestivalFilter();
          if (adminFestivalFilter) adminFestivalFilter.value = currentFilter;
          renderAdminFestivalContacts();
        } else if (activeCampaign === "Cultivation") {
          loadCultivationContacts();
        } else if (activeCampaign !== "Reception") {
          renderContacts();
        }
      }
    }
  }, 15000);
}
