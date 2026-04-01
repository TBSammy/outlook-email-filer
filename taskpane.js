/* =====================================================
   Email Filer — Outlook Web Add-in
   Files the selected email as .eml to a SharePoint folder
   ===================================================== */

// === Configuration (not secrets — just public identifiers) ===
const CONFIG = {
  clientId: "0ec52073-2a04-4163-9cb1-ee33c13d8a17",
  tenantId: "9375944b-0809-47d8-9008-fc31cd8641e5",
  siteHostname: "muskratss.sharepoint.com",
  sitePath: "/sites/project-email-register",
  emailRecordsFolder: "Email Records",
  scopes: ["Files.ReadWrite", "Sites.Read.All"],
};

// === MSAL setup ===
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

let msalClient = null;
let accessToken = null;
let siteId = null;
let driveId = null;
let emailRecordsFolderId = null;

// Current navigation state
let currentFolderId = null;
let currentFolderName = null;
let breadcrumbPath = []; // [{id, name}]
let selectedFolderId = null;
let selectedFolderName = null;

// === DOM references ===
const statusText = document.getElementById("status-text");
const loginSection = document.getElementById("login-section");
const loginBtn = document.getElementById("login-btn");
const folderSection = document.getElementById("folder-section");
const breadcrumbEl = document.getElementById("breadcrumb");
const folderList = document.getElementById("folder-list");
const selectedFolderEl = document.getElementById("selected-folder");
const selectedFolderNameEl = document.getElementById("selected-folder-name");
const fileBtn = document.getElementById("file-btn");
const resultSection = document.getElementById("result-section");
const progressEl = document.getElementById("progress");
const successEl = document.getElementById("success");
const errorEl = document.getElementById("error");
const errorText = document.getElementById("error-text");
const doneBtn = document.getElementById("done-btn");
const retryBtn = document.getElementById("retry-btn");

// === Initialize ===
Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) {
    statusText.textContent = "This add-in only works in Outlook.";
    return;
  }

  msalClient = new msal.PublicClientApplication(msalConfig);

  // Check if already signed in
  const accounts = msalClient.getAllAccounts();
  if (accounts.length > 0) {
    acquireTokenSilent(accounts[0]);
  } else {
    showLogin();
  }
});

// === Auth ===
function showLogin() {
  statusText.textContent = "Please sign in to continue.";
  loginSection.classList.remove("hidden");
}

loginBtn.addEventListener("click", async function () {
  loginBtn.disabled = true;
  statusText.textContent = "Signing in...";
  try {
    const response = await msalClient.loginPopup({
      scopes: CONFIG.scopes,
    });
    await acquireTokenSilent(response.account);
  } catch (err) {
    statusText.textContent = "Sign-in failed. Try again.";
    loginBtn.disabled = false;
    console.error("Login error:", err);
  }
});

async function acquireTokenSilent(account) {
  try {
    const response = await msalClient.acquireTokenSilent({
      scopes: CONFIG.scopes,
      account: account,
    });
    accessToken = response.accessToken;
    loginSection.classList.add("hidden");
    statusText.textContent = "Signed in as " + account.username;
    await initSharePoint();
  } catch (err) {
    // Silent failed, try popup
    try {
      const response = await msalClient.acquireTokenPopup({
        scopes: CONFIG.scopes,
      });
      accessToken = response.accessToken;
      loginSection.classList.add("hidden");
      statusText.textContent = "Signed in as " + response.account.username;
      await initSharePoint();
    } catch (popupErr) {
      showLogin();
      console.error("Token error:", popupErr);
    }
  }
}

// === Graph API helper ===
async function graphGet(url) {
  const response = await fetch(url, {
    headers: { Authorization: "Bearer " + accessToken },
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error("Graph API error " + response.status + ": " + text);
  }
  return response.json();
}

async function graphUpload(url, content, contentType) {
  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: "Bearer " + accessToken,
      "Content-Type": contentType || "application/octet-stream",
    },
    body: content,
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error("Upload error " + response.status + ": " + text);
  }
  return response.json();
}

// === SharePoint init ===
async function initSharePoint() {
  statusText.textContent = "Loading SharePoint folders...";
  try {
    // Get site ID
    const site = await graphGet(
      "https://graph.microsoft.com/v1.0/sites/" +
        CONFIG.siteHostname +
        ":" +
        CONFIG.sitePath
    );
    siteId = site.id;

    // Get default drive
    const drives = await graphGet(
      "https://graph.microsoft.com/v1.0/sites/" + siteId + "/drives"
    );
    const drive = drives.value.find(function (d) {
      return d.name === "Documents";
    });
    if (!drive) throw new Error("Documents library not found");
    driveId = drive.id;

    // Find the "Email Records" folder
    const rootChildren = await graphGet(
      "https://graph.microsoft.com/v1.0/drives/" +
        driveId +
        "/root/children?$filter=name eq '" +
        CONFIG.emailRecordsFolder +
        "'"
    );
    if (rootChildren.value.length === 0) {
      throw new Error('"Email Records" folder not found in Documents library');
    }
    emailRecordsFolderId = rootChildren.value[0].id;

    // Navigate to Email Records
    breadcrumbPath = [
      { id: emailRecordsFolderId, name: CONFIG.emailRecordsFolder },
    ];
    await loadFolder(emailRecordsFolderId, CONFIG.emailRecordsFolder);

    folderSection.classList.remove("hidden");
    statusText.textContent = "Select a folder to file the email into.";
  } catch (err) {
    statusText.textContent = "Error: " + err.message;
    console.error("SharePoint init error:", err);
  }
}

// === Folder navigation ===
async function loadFolder(folderId, folderName) {
  currentFolderId = folderId;
  currentFolderName = folderName;
  selectedFolderId = null;
  selectedFolderName = null;
  selectedFolderEl.classList.add("hidden");

  folderList.innerHTML = '<div class="folder-item">Loading...</div>';

  try {
    const children = await graphGet(
      "https://graph.microsoft.com/v1.0/drives/" +
        driveId +
        "/items/" +
        folderId +
        "/children?$filter=folder ne null&$orderby=name"
    );

    renderBreadcrumb();
    renderFolders(children.value);
  } catch (err) {
    folderList.innerHTML =
      '<div class="folder-item">Error loading folders</div>';
    console.error("Load folder error:", err);
  }
}

function renderBreadcrumb() {
  breadcrumbEl.innerHTML = "";
  breadcrumbPath.forEach(function (crumb, i) {
    if (i > 0) {
      var sep = document.createElement("span");
      sep.className = "crumb-separator";
      sep.textContent = "›";
      breadcrumbEl.appendChild(sep);
    }
    var el = document.createElement("span");
    el.className = "crumb";
    el.textContent = crumb.name;
    el.dataset.id = crumb.id;
    if (i < breadcrumbPath.length - 1) {
      el.addEventListener("click", function () {
        // Navigate back to this level
        breadcrumbPath = breadcrumbPath.slice(0, i + 1);
        loadFolder(crumb.id, crumb.name);
      });
    }
    breadcrumbEl.appendChild(el);
  });
}

function renderFolders(folders) {
  folderList.innerHTML = "";

  if (folders.length === 0) {
    // No subfolders — this folder is a valid target
    folderList.innerHTML =
      '<div class="folder-item" style="color:#605e5c;font-style:italic;">No subfolders. You can file here.</div>';
    selectFolder(currentFolderId, currentFolderName);
    return;
  }

  folders.forEach(function (folder) {
    var item = document.createElement("div");
    item.className = "folder-item";
    item.innerHTML =
      '<span class="folder-icon">📁</span>' +
      '<span class="folder-name">' +
      escapeHtml(folder.name) +
      "</span>" +
      (folder.folder.childCount > 0
        ? '<span class="folder-arrow">›</span>'
        : "");

    item.addEventListener("click", function () {
      // Single click = select, double click or arrow = drill in
      if (selectedFolderId === folder.id) {
        // Already selected — drill in if it has children
        if (folder.folder.childCount > 0) {
          breadcrumbPath.push({ id: folder.id, name: folder.name });
          loadFolder(folder.id, folder.name);
        }
      } else {
        selectFolder(folder.id, folder.name);
        // Highlight
        document.querySelectorAll(".folder-item").forEach(function (el) {
          el.classList.remove("selected");
        });
        item.classList.add("selected");
      }
    });

    // Double-click to drill in
    item.addEventListener("dblclick", function () {
      if (folder.folder.childCount > 0) {
        breadcrumbPath.push({ id: folder.id, name: folder.name });
        loadFolder(folder.id, folder.name);
      }
    });

    folderList.appendChild(item);
  });
}

function selectFolder(folderId, folderName) {
  selectedFolderId = folderId;
  selectedFolderName = folderName;
  selectedFolderNameEl.textContent = folderName;
  selectedFolderEl.classList.remove("hidden");
}

// === File email ===
fileBtn.addEventListener("click", function () {
  fileEmail();
});

async function fileEmail() {
  if (!selectedFolderId) return;

  folderSection.classList.add("hidden");
  resultSection.classList.remove("hidden");
  progressEl.classList.remove("hidden");
  successEl.classList.add("hidden");
  errorEl.classList.add("hidden");

  try {
    // Get the email as EML via Office.js (Mailbox 1.14)
    var emlBase64 = await getEmailAsFile();
    var emlBytes = base64ToArrayBuffer(emlBase64);

    // Build filename from subject + date
    var subject = Office.context.mailbox.item.subject || "email";
    var dateReceived = Office.context.mailbox.item.dateTimeCreated;
    var dateStr = formatDate(dateReceived);
    var fileName = sanitizeFileName(subject) + " (" + dateStr + ").eml";

    // Upload to SharePoint
    await graphUpload(
      "https://graph.microsoft.com/v1.0/drives/" +
        driveId +
        "/items/" +
        selectedFolderId +
        ":/" +
        encodeURIComponent(fileName) +
        ":/content",
      emlBytes,
      "message/rfc822"
    );

    progressEl.classList.add("hidden");
    successEl.classList.remove("hidden");
  } catch (err) {
    progressEl.classList.add("hidden");
    errorEl.classList.remove("hidden");
    errorText.textContent = err.message;
    console.error("File email error:", err);
  }
}

function getEmailAsFile() {
  return new Promise(function (resolve, reject) {
    if (
      !Office.context.requirements.isSetSupported("Mailbox", "1.14") ||
      !Office.context.mailbox.item.getAsFileAsync
    ) {
      reject(
        new Error(
          "getAsFileAsync not supported. Requires Mailbox 1.14 or later."
        )
      );
      return;
    }
    Office.context.mailbox.item.getAsFileAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

// === Reset / retry ===
doneBtn.addEventListener("click", function () {
  resultSection.classList.add("hidden");
  folderSection.classList.remove("hidden");
});

retryBtn.addEventListener("click", function () {
  resultSection.classList.add("hidden");
  folderSection.classList.remove("hidden");
});

// === Utilities ===
function base64ToArrayBuffer(base64) {
  var binary = atob(base64);
  var len = binary.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function sanitizeFileName(name) {
  return name
    .replace(/[<>:"/\\|?*]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, 100);
}

function formatDate(date) {
  if (!date) return "no-date";
  var d = new Date(date);
  return (
    d.getFullYear() +
    "-" +
    String(d.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(d.getDate()).padStart(2, "0")
  );
}

function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
