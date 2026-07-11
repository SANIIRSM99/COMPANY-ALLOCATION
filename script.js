const demoUsers = [
   "ADMIN", "ALL", "KHALID", "ASIF", "MUZAMMIL", "HAIDER", "IMRAN", "WAQAS", "MURTAZA", "YOUSAF", "AMJID", "ALI",
    "SOHAIL", "SHOAIB", "IQBAL", "ATIF", "FAQIR", "FAROOQ", "JAVAID", "AMRAN", "BILAL", "ZEESHAN"
].map(u => ({ username: u.trim(), password: "123" }));

let unlockCode = null;
let isAppLocked = true;

function demoLogin(username, password) {
    const found = demoUsers.find(x => x.username.toUpperCase() === username.toUpperCase() && x.password === password);
    if (found) {
        localStorage.setItem("loggedUser", found.username);
        return true;
    }
    return false;
}

function getLoggedUser() {
    return localStorage.getItem("loggedUser") || null;
}

function logoutDemo() {
    localStorage.removeItem("loggedUser");
    location.reload();
}

// ---------------- Parse CSV/Excel ----------------
let allCSVData = []; // ✅ Global variable: all rows save ہوں گے

// ---------------- Parse CSV File (Excel or CSV) ----------------
function parseCSVandFilter(file, onDone) {
    const reader = new FileReader();
    const isExcel = file.name.endsWith(".xlsx") || file.name.endsWith(".xls");

    if (isExcel) {
        // ✅ Excel file reading (no limit)
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const csv = XLSX.utils.sheet_to_csv(sheet);
            processCSV(csv, onDone);
        };
        reader.readAsArrayBuffer(file);
    } else {
        // ✅ CSV (Unlimited rows)
        const CHUNK_SIZE = 1024 * 1024 * 2; // 2MB chunk
        let offset = 0;
        let csvText = "";

        const readChunk = () => {
            const slice = file.slice(offset, offset + CHUNK_SIZE);
            const chunkReader = new FileReader();

            chunkReader.onload = (e) => {
                csvText += e.target.result;
                offset += CHUNK_SIZE;

                const percent = Math.min(100, ((offset / file.size) * 100).toFixed(1));
                console.log(`⏳ Reading CSV: ${percent}%`);

                if (offset < file.size) {
                    readChunk(); // Read next chunk
                } else {
                    console.log("✅ CSV Loaded Completely. Processing...");
                    // ✅ Make sure last line is complete
                    if (!csvText.endsWith("\n")) csvText += "\n";
                    processCSV(csvText, onDone);
                }
            };

            chunkReader.onerror = (err) => {
                console.error("❌ Error reading CSV file:", err);
            };

            chunkReader.readAsText(slice, "UTF-8");
        };

        readChunk();
    }
}

// ---------------- Process CSV ----------------
function processCSV(text, onDone) {
    // ✅ Safe large-split method
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    console.log("📦 Raw Lines Found:", lines.length);

    // ✅ Map safely even for large CSVs
    const rows = lines.map((line) =>
        line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map((c) => c.replace(/^"|"$/g, ""))
    );

    const logged = getLoggedUser();
    const loggedUpper = (logged || "").toString().trim().toUpperCase();
    const activeCsvUser = getActiveDataUser();
    const filterUser = loggedUpper === "ADMIN"
        ? (activeCsvUser && activeCsvUser !== "ADMIN" && activeCsvUser !== "ALL" ? activeCsvUser : "")
        : loggedUpper;
    const filtered = filterUser
        ? rows.filter(
              (r) =>
                  (r[6] || "").toString().trim().toUpperCase() === filterUser ||
                  (r[7] || "").toString().trim().toUpperCase() === filterUser
          )
        : rows;

    const mapMainRow = (row) => ({
        City: row[0] || "",
        CustomerCode: (row[1] || "").trim().toUpperCase(),
        Customer: row[2] || "",
        Item1: (row[3] || "").trim().toUpperCase(),
        Target1: parseInt(row[4]) || 0,
        Achieve1: parseInt(row[5]) || 0,
        User1: row[6] || "",
        User2: row[7] || "",
        DealQty: parseInt(row[8]) || 0,
        DealBonus: parseInt(row[9]) || 0,
        SummaryNumber: row[10] || "",
        CompanyName: row[11] || "",
        Value: parseFloat((row[12] || "0").replace(/,/g, "")) || 0,
        Date: row[13] || "",
        ItemRate: parseFloat((row[14] || "0").replace(/,/g, "")) || 0
    });
    let mappedAllRows = rows.map(mapMainRow);
    let mapped = filtered.map(mapMainRow);
    const previousAllRows = JSON.parse(localStorage.getItem("excelDataAll") || "[]");
    const previousVisibleRows = JSON.parse(localStorage.getItem("excelData") || "[]");
    if (typeof mergeCsvRowsKeepTargets === "function") {
        mappedAllRows = mergeCsvRowsKeepTargets(previousAllRows, mappedAllRows);
        mapped = filterUser ? filterRowsForUser(mappedAllRows, filterUser) : mergeCsvRowsKeepTargets(previousVisibleRows, mapped);
    }

    console.log("✅ Total CSV Rows:", lines.length);
    console.log("✅ Filtered Rows (after user filter):", filtered.length);
    console.log("✅ Final Mapped Rows:", mapped.length);

    localStorage.setItem("excelDataAll", JSON.stringify(mappedAllRows));
    allCSVData = mapped; // Save visible rows globally

    // ✅ Invoices = Achieve > 0 rows
    invoices = mapped
        .filter((r) => r.CustomerCode && r.Item1)
        .map((r) => ({
            city: r.City,
            customerCode: r.CustomerCode,
            customer: r.Customer,
            item: r.Item1,
            target: r.Target1,
            quantity: r.Achieve1,
            rate: r.ItemRate,
            user: r.User1 || r.User2 || logged || ""
        }));

    localStorage.setItem("invoices", JSON.stringify(invoices));

    // ✅ Bonus Deals
    bonusDeals = {};
    mapped.forEach((row) => {
        const item = row.Item1;
        if (!item) return;
        if (!bonusDeals[item]) bonusDeals[item] = [];
        if (row.DealQty > 0 || row.DealBonus > 0) {
            bonusDeals[item].push({ qty: row.DealQty, bonus: row.DealBonus });
        }
    });
    localStorage.setItem("bonusDeals", JSON.stringify(bonusDeals));

    // ✅ Render updates
    if (typeof renderInvoiceTable === "function") renderInvoiceTable(mapped);
    if (typeof renderMySaleTable === "function") renderMySaleTable();

   if (onDone) onDone(mapped);

// ✅ Create a unique hash from current data to detect duplicate uploads
const uploadRows = loggedUpper === "ADMIN" ? mappedAllRows : mapped;
const currentHash = btoa(JSON.stringify(uploadRows)).slice(0, 100);
const lastMeta = JSON.parse(localStorage.getItem("lastCsvMeta") || "{}");
const loggedUser = getLoggedUser() || "UNKNOWN_USER";

// ✅ Always use same file name (if undefined)
const csvFileName = (typeof file !== "undefined" && file.name) ? file.name : "latest_upload.csv";

// ✅ Case 1: If same data (skip)
if (lastMeta.hash === currentHash && lastMeta.user === loggedUser) {
  console.warn("⏸ Same CSV data detected — skipping upload.");
} else {
  console.log("🚀 Uploading new or updated CSV to Firebase...");

  // Save new hash for comparison next time
  localStorage.setItem("lastCsvMeta", JSON.stringify({
    name: csvFileName,
    hash: currentHash,
    user: loggedUser,
    time: new Date().toISOString()
  }));

  // ✅ Upload processed data to Realtime DB
  saveCSVToFirebase(uploadRows);

  // ✅ Optional: upload raw file (only if available)
  try {
    if (typeof firebase !== "undefined" && firebase.storage && typeof file !== "undefined") {
      firebase.storage().ref('csvUploads/' + csvFileName).put(file)
        .then(() => console.log('✅ CSV uploaded successfully!'))
        .catch(err => console.error('❌ Firebase upload failed:', err));
    }
  } catch (err) {
    console.error("⚠️ Firebase Storage skipped:", err);
  }
}


}



let excelData = [];
let invoices = [];
let doneTargets = [];
let customers = [];
let customerCodes = [];
let items = [];
let customerTargets = {};
let isLoggedIn = false;
let bonusDeals = {};
let lastRenderedCustomerCode = null;

function getActiveDataUser() {
  const logged = (getLoggedUser() || "").toString().trim().toUpperCase();
  if (logged === "ADMIN" && typeof document !== "undefined") {
    const selected = (document.getElementById("userSelect")?.value || "").toString().trim().toUpperCase();
    if (selected) {
      localStorage.setItem("activeDataUser", selected);
      return selected;
    }
  }
  return (localStorage.getItem("activeDataUser") || logged || "").toString().trim().toUpperCase();
}

function setActiveDataUser(user) {
  const clean = (user || getLoggedUser() || "").toString().trim().toUpperCase();
  if (clean) localStorage.setItem("activeDataUser", clean);
  return clean;
}

/**
 * Save processed CSV rows (mapped array) online.
 * Uses: 1) if window.FIREBASE_UPLOAD_ENDPOINT set -> POST there
 *       2) else if DATABASE_URL set -> upload to Firebase Realtime DB via REST
 *       3) else -> fallback: save to localStorage and console.warn
 *
 * Expects `data` = array of objects (mapped rows)
 */


function saveCSVToFirebase(data) {
  try {
    if (!data) return;
    const loggedUser = getLoggedUser();
    if (!loggedUser) {
      console.warn("⚠️ No logged-in user — saving locally instead.");
      localStorage.setItem("excelData", JSON.stringify(data));
      return;
    }

    const payload = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: getLoggedUser() || loggedUser,
      rows: data
    };


    const targetUploadUser = getActiveDataUser() || loggedUser.toUpperCase();
    const url = `${DATABASE_URL}/csvUploads/${targetUploadUser}/latest.json`;

    fetch(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    })
      .then(res => {
        if (!res.ok) throw new Error("HTTP " + res.status);
        console.log("✅ Firebase updated successfully!");
      })
      .catch(err => {
        console.error("❌ Upload failed:", err);
        localStorage.setItem("excelData", JSON.stringify(data));
      });
  } catch (err) {
    console.error("❌ saveCSVToFirebase error:", err);
  }
}

async function deleteUserData(userToDelete) {
    const loggedUser = getLoggedUser();

    if (!loggedUser) {
        alert("No logged in user!");
        return;
    }

    if (loggedUser.toUpperCase() !== "ADMIN") {
        alert("Only ADMIN can delete data!");
        return;
    }

    userToDelete = (userToDelete || "").toString().trim().toUpperCase();
    if (!userToDelete || userToDelete === "ADMIN") {
        alert("Please select a valid user to delete.");
        return;
    }

    const isAllDelete = userToDelete === "ALL";
    const confirmText = isAllDelete
        ? "Are you sure you want to DELETE ALL Firebase data for every user/booker? This will delete central data, user wise data, and My Sale data."
        : `Are you sure you want to DELETE all Firebase data of: ${userToDelete}?`;

    if (!confirm(confirmText)) {
        return;
    }

    try {
        const paths = isAllDelete
            ? [
                `${DATABASE_URL}/csvUploads.json`,
                `${DATABASE_URL}/mySales.json`
            ]
            : [
                `${DATABASE_URL}/csvUploads/${userToDelete}.json`,
                `${DATABASE_URL}/mySales/${userToDelete}.json`,
                `${DATABASE_URL}/csvUploads/${userToDelete}/mySales.json`
            ];
        const results = await Promise.all(paths.map(url => fetch(url, { method: "DELETE" })));
        const failed = results.find(res => !res.ok);
        if (failed) throw new Error("Failed: " + failed.status);

        if (!isAllDelete) {
            await removeUserRowsFromAllUpload(userToDelete);
        }

        if (isAllDelete || getActiveDataUser() === userToDelete) {
            excelData = [];
            invoices = [];
            mySaleData = [];
            localStorage.removeItem("excelData");
            localStorage.removeItem("excelDataAll");
            localStorage.removeItem("invoices");
            localStorage.removeItem("mySaleData");
            localStorage.removeItem("lastCsvUploadRef");
            buildCustomerTargets();
            renderInvoiceTable();
            renderMySaleTable?.();
        }
        showAppNotification(isAllDelete ? "All Firebase data deleted successfully." : `Firebase data deleted for ${userToDelete}.`, "success");
    } catch (err) {
        console.error("Delete error:", err);
        showAppNotification("Error deleting data from Firebase.", "error");
    }
}

async function removeUserRowsFromAllUpload(userToDelete) {
    try {
        if (typeof DATABASE_URL !== "string" || !DATABASE_URL) return;
        const cleanUser = (userToDelete || "").toString().trim().toUpperCase();
        if (!cleanUser || cleanUser === "ALL" || cleanUser === "ADMIN") return;

        const url = `${DATABASE_URL}/csvUploads/ALL/latest.json`;
        const res = await fetch(url);
        if (!res.ok) return;
        const payload = await res.json();
        if (!payload || !Array.isArray(payload.rows)) return;

        const keptRows = payload.rows.filter(row => {
            const user1 = (row.User1 || row.user1 || row.USER || row.user || "").toString().trim().toUpperCase();
            const user2 = (row.User2 || row.user2 || "").toString().trim().toUpperCase();
            return user1 !== cleanUser && user2 !== cleanUser;
        });

        if (keptRows.length === payload.rows.length) return;
        const updatedPayload = {
            ...payload,
            rows: keptRows,
            updatedAt: new Date().toISOString(),
            updatedBy: getLoggedUser() || "ADMIN"
        };

        await fetch(url, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(updatedPayload)
        });
    } catch (err) {
        console.warn("Could not remove user rows from ALL upload:", err);
    }
}
document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("deleteUserBtn");
    if (!btn) return;

    btn.addEventListener("click", () => {
        const user = document.getElementById("userSelect").value;
        deleteUserData(user);
    });
});





function syncUserDataFromFirebase(onDone) {
  const loggedUser = getLoggedUser();
  if (!loggedUser) {
    console.warn('No logged-in user. Cannot sync data.');
    if (onDone) onDone([]);
    return;
  }

  if (typeof DATABASE_URL !== 'string' || DATABASE_URL.length === 0) {
    console.warn('No Firebase DATABASE_URL configured. Using local data.');
    const localData = JSON.parse(localStorage.getItem('excelData') || '[]');
    processCSVData(localData, onDone);
    return;
  }

  const userPath = `csvUploads/${loggedUser.toUpperCase()}`;
  const url = `${DATABASE_URL}/${userPath}.json`;

  fetch(url)
    .then(res => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    })
    .then(data => {
      let allRows = [];
      if (data) {
        // Flatten all uploads for the user
        Object.values(data).forEach(upload => {
          if (upload.rows && Array.isArray(upload.rows)) {
            allRows = allRows.concat(upload.rows);
          }
        });
      }
      console.log(`✅ Fetched ${allRows.length} rows for user ${loggedUser}`);
      processCSVData(allRows, onDone);
    })
    .catch(err => {
      console.error('❌ Failed to fetch user data from Firebase:', err);
      const localData = JSON.parse(localStorage.getItem('excelData') || '[]');
      processCSVData(localData, onDone);
    });
}




function buildCustomerTargets() {
    console.log('Building customer targets from excelData:', excelData);
    customerTargets = {};
    customers = [];
    customerCodes = [];
    items = [];
    bonusDeals = {};
    
    excelData.forEach(row => {
        const customerCode = (row.CustomerCode || '').trim().toUpperCase();
        const customer = (row.Customer || '').trim();
        const city = (row.City || '').trim();
        const item = (row.Item1 || '').trim().toUpperCase();
        const target = Number(row.Target1 || 0);
        const dealQty = row.DealQty;
        const dealBonus = row.DealBonus;

        if (!customer || !customerCode || !city) {
            console.warn('Skipping row due to missing customer data:', row);
            return;
        }
        if (!customerCodes.includes(customerCode)) {
            customerCodes.push(customerCode);
            customers.push({ code: customerCode, name: customer, city: city });
        }
        if (!customerTargets[customerCode]) {
            customerTargets[customerCode] = { name: customer, city: city, items: {} };
        }
        if (item && target >= 0) {
            customerTargets[customerCode].items[item] = (customerTargets[customerCode].items[item] || 0) + target;
            if (!items.includes(item)) items.push(item);
        }

        if (item && dealQty > 0 && dealBonus > 0) {
            if (!bonusDeals[item]) bonusDeals[item] = [];
            bonusDeals[item].push({ qty: dealQty, bonus: dealBonus });
        }
    });

    console.log('Customer targets built:', customerTargets);
    console.log('Items extracted:', items);
    console.log('Bonus deals built:', bonusDeals);
    localStorage.setItem('items', JSON.stringify(items));
    localStorage.setItem('customers', JSON.stringify(customers));
    localStorage.setItem('customerCodes', JSON.stringify(customerCodes));
    localStorage.setItem('bonusDeals', JSON.stringify(bonusDeals));
    updateCityDropdown();
    renderBonusDeals();
    populateBonusItems();
}

function updateCityDropdown() {
    const citySelect = document.getElementById('citySelect');
    if (!citySelect) return;
    const cities = [...new Set(excelData.map(row => row.City?.trim()))].filter(city => city);
    console.log('Cities for dropdown:', cities);
    citySelect.innerHTML = '<option value="">Select a city</option>';
    cities.forEach(city => {
        const option = document.createElement('option');
        option.value = city;
        option.textContent = city;
        citySelect.appendChild(option);
    });
}

function generateUnlockCode() {
    const randomCode = Math.floor(100000 + Math.random() * 900000);
    localStorage.setItem('displayCode', randomCode);
    const finalCode = (randomCode * 2) + 985973;
    return finalCode;
}

function checkLockStatus() {
    const codeSection = document.getElementById('codeSection');
    const loginPage = document.getElementById('loginPage');
    const mainPage = document.getElementById('mainPage');
    const sidebar = document.getElementById('sidebar');
    const hamburgerContainer = document.getElementById('hamburgerContainer');
    const displayCodeElement = document.getElementById('displayCode');

    if (!codeSection || !displayCodeElement || !loginPage || !mainPage || !sidebar || !hamburgerContainer) {
        console.error('Critical DOM elements missing:', { codeSection, displayCodeElement, loginPage, mainPage, sidebar, hamburgerContainer });
        return;
    }

    codeSection.classList.add('fixed', 'top-1/2', 'left-1/2', 'transform', '-translate-x-1/2', '-translate-y-1/2', 'z-50', 'bg-white', 'p-6', 'rounded-lg', 'shadow-lg', 'w-80', 'max-w-[90%]');

    let overlay = document.getElementById('lockOverlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'lockOverlay';
        overlay.classList.add('fixed', 'inset-0', 'bg-black', 'bg-opacity-50', 'z-40', 'hidden');
        document.body.appendChild(overlay);
    }

    const today = new Date();
    const currentYearMonth = `${today.getFullYear()}-${today.getMonth() + 1}`;
    const lastUnlockMonth = localStorage.getItem('lastUnlockMonth');
    const isNewDevice = !localStorage.getItem('deviceInitialized');
    const isFirstOfMonth = today.getDate() === 1;
    const storedIsAppLocked = localStorage.getItem('isAppLocked') === 'true';

    // Initialize device on first run
    if (isNewDevice) {
        localStorage.setItem('deviceInitialized', 'true');
        localStorage.setItem('isAppLocked', 'true');
    }

    // Check if app should be locked
    if (isNewDevice || (isFirstOfMonth && lastUnlockMonth !== currentYearMonth) || storedIsAppLocked) {
        isAppLocked = true;
        localStorage.setItem('isAppLocked', 'true');
        unlockCode = generateUnlockCode();
        localStorage.setItem('unlockCode', unlockCode);
        localStorage.setItem('lastLockCheck', today.toISOString());
        displayCodeElement.textContent = localStorage.getItem('displayCode');
        codeSection.classList.remove('hidden');
        overlay.classList.remove('hidden');
        loginPage.classList.add('hidden');
        mainPage.classList.add('hidden');
        sidebar.classList.add('hidden', '-translate-x-full');
        hamburgerContainer.classList.add('hidden');
        console.log('Lock popup shown with code:', localStorage.getItem('displayCode'));
    } else {
        isAppLocked = false;
        localStorage.setItem('isAppLocked', 'false');
        codeSection.classList.add('hidden');
        overlay.classList.add('hidden');
        const loggedUser = getLoggedUser();
        if (loggedUser) {
            isLoggedIn = true;
            loginPage.classList.add('hidden');
            mainPage.classList.remove('hidden');
            sidebar.classList.add('md:block');
            hamburgerContainer.classList.remove('hidden');
            initSidebarNav();
            renderInvoiceTable();
        } else {
            loginPage.classList.remove('hidden');
            mainPage.classList.add('hidden');
            sidebar.classList.add('hidden', '-translate-x-full');
            hamburgerContainer.classList.add('hidden');
        }
        console.log('App is unlocked, showing login or main page');
    }
}

function unlockApp() {
    const unlockCodeInput = document.getElementById('unlockCode');
    const codeError = document.getElementById('codeError');
    const codeSection = document.getElementById('codeSection');
    const overlay = document.getElementById('lockOverlay');
    if (!unlockCodeInput || !codeError || !codeSection || !overlay) {
        console.error('Unlock DOM elements missing:', { unlockCodeInput, codeError, codeSection, overlay });
        return;
    }

    const enteredCode = unlockCodeInput.value.trim();
    const storedUnlockCode = parseInt(localStorage.getItem('unlockCode'));

    // ✅ Admin Master Code
    const adminCode = "123ND";

    if (enteredCode === adminCode || parseInt(enteredCode) === storedUnlockCode) {
        isAppLocked = false;
        localStorage.setItem('isAppLocked', 'false');
        const today = new Date();
        const currentYearMonth = `${today.getFullYear()}-${today.getMonth() + 1}`;
        localStorage.setItem('lastUnlockMonth', currentYearMonth);
        localStorage.removeItem('unlockCode');
        localStorage.removeItem('displayCode');
        localStorage.setItem('lastLockCheck', today.toISOString());
        codeSection.classList.add('hidden');
        overlay.classList.add('hidden');
        document.getElementById('loginPage').classList.remove('hidden');
        codeError.classList.add('hidden');
        unlockCodeInput.value = '';
        console.log('✅ App unlocked successfully for month:', currentYearMonth);
    } else {
        codeError.classList.remove('hidden');
        console.error('❌ Invalid unlock code entered:', enteredCode);
    }
}


function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        sidebar.classList.toggle('hidden');
        sidebar.classList.toggle('-translate-x-full');
    }
}

function initSidebarNav() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) {
        console.error('Sidebar element not found');
        return;
    }
    const buttons = sidebar.querySelectorAll('button');
    buttons.forEach(button => {
        button.removeEventListener('click', handleSidebarClick);
        button.addEventListener('click', handleSidebarClick);
    });
    const loggedUserName = getLoggedUser();
    const userNameEls = document.querySelectorAll('#loggedUserName');
    userNameEls.forEach(el => {
        el.textContent = loggedUserName || 'User';
    });
}

function safeTargetInputId(customerCode, item) {
  return `zt_${encodeURIComponent(customerCode)}_${encodeURIComponent(item)}`.replace(/%/g, "_");
}

function getUserForDataRow(row) {
  const active = getActiveDataUser();
  if (active && active !== "ADMIN" && active !== "ALL") return active;
  return ((row?.User1 || row?.User2 || getLoggedUser() || "")).toString().trim().toUpperCase();
}

async function saveRowsToFirebaseUser(rows, user) {
  const targetUser = (user || "").toString().trim().toUpperCase();
  if (!targetUser || targetUser === "ALL" || typeof DATABASE_URL !== "string" || !DATABASE_URL) return;
  const payload = {
    uploadedAt: new Date().toISOString(),
    uploadedBy: getLoggedUser() || targetUser,
    rows: rows || []
  };
  await fetch(`${DATABASE_URL}/csvUploads/${targetUser}/latest.json`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
}

async function saveTargetUpdateToFirebase() {
  const rows = JSON.parse(localStorage.getItem("excelData") || "[]");
  const byUser = {};
  rows.forEach(row => {
    const user = getUserForDataRow(row);
    if (!user || user === "ALL") return;
    if (!byUser[user]) byUser[user] = [];
    byUser[user].push(row);
  });
  await Promise.all(Object.entries(byUser).map(([user, userRows]) => saveRowsToFirebaseUser(userRows, user)));
}

function setTargetForZeroItem(customerCode, item) {
  const input = document.getElementById(safeTargetInputId(customerCode, item));
  const newTarget = Number(input?.value || 0);
  if (!newTarget || newTarget <= 0) {
    alert("Please enter target greater than 0.");
    return;
  }
  excelData = (excelData || []).map(row => {
    const code = (row.CustomerCode || "").toString().trim().toUpperCase();
    const rowItem = (row.Item1 || "").toString().trim().toUpperCase();
    if (code === customerCode.toUpperCase() && rowItem === item.toUpperCase()) {
      return { ...row, Target1: newTarget };
    }
    return row;
  });
  localStorage.setItem("excelData", JSON.stringify(excelData));
  saveTargetUpdateToFirebase();
  buildCustomerTargets();
  renderInvoiceTable();
}

function applyTargetToAllZeroItems() {
  const input = document.getElementById("zeroTargetApplyAllValue");
  const newTarget = Number(input?.value || 0);
  if (!newTarget || newTarget <= 0) {
    alert("Please enter target greater than 0.");
    return;
  }
  let updated = 0;
  excelData = (excelData || []).map(row => {
    const hasItem = (row.CustomerCode || "").toString().trim() && (row.Item1 || "").toString().trim();
    if (hasItem && Number(row.Target1 || 0) === 0) {
      updated++;
      return { ...row, Target1: newTarget };
    }
    return row;
  });
  localStorage.setItem("excelData", JSON.stringify(excelData));
  saveTargetUpdateToFirebase();
  buildCustomerTargets();
  renderInvoiceTable();
  alert(`Target applied to ${updated} zero target rows.`);
}

function getRankDisplay(level) {
  if (level === "Golden") return "Golden";
  if (level === "Silver") return "Silver";
  if (level === "Bronze") return "Bronze";
  return level || "";
}

function getRankColor(level) {
  if (level === "Golden") return "#F59E0B";
  if (level === "Silver") return "#94A3B8";
  if (level === "Bronze") return "#B45309";
  return "#2563EB";
}

function getCustomerRankings(sourceTargets = customerTargets) {
  const allCustomers = Object.entries(sourceTargets || {}).map(([code, data]) => {
    const items = data.items || {};
    const totalTargetQty = Object.values(items).reduce((sum, qty) => sum + (Number(qty) || 0), 0);
    const itemsCount = Object.keys(items).length;
    const achievedValue = (invoices || [])
      .filter(inv => inv.customerCode?.toUpperCase() === code.toUpperCase())
      .reduce((sum, inv) => sum + ((Number(inv.quantity) || 0) * (Number(inv.rate) || 0)), 0);
    return {
      code,
      name: data.name || "Unknown",
      itemsCount,
      totalTargetQty,
      achievedValue,
      rankScore: totalTargetQty + itemsCount + achievedValue
    };
  }).sort((a, b) =>
    a.rankScore - b.rankScore ||
    a.totalTargetQty - b.totalTargetQty ||
    a.itemsCount - b.itemsCount ||
    a.name.localeCompare(b.name)
  );

  const uniqueScores = [...new Set(allCustomers.map(cust => cust.rankScore))];
  return allCustomers.map(cust => {
    const groupIndex = uniqueScores.indexOf(cust.rankScore);
    const fromTop = uniqueScores.length - 1 - groupIndex;
    let level = `Rank ${groupIndex + 1}`;
    if (fromTop === 0) level = "Golden";
    else if (fromTop === 1) level = "Silver";
    else if (fromTop === 2) level = "Bronze";
    return {
      ...cust,
      level,
      displayLevel: getRankDisplay(level),
      levelColor: getRankColor(level)
    };
  });
}

function populateRankFilter(rankedCustomers) {
  const rankFilter = document.getElementById("rankFilter");
  if (!rankFilter) return;
  const current = rankFilter.value || "all";
  const levels = [...new Set((rankedCustomers || []).map(c => c.level).filter(Boolean))];
  rankFilter.innerHTML = `<option value="all">All Ranks</option>` + levels
    .map(level => `<option value="${level}">${getRankDisplay(level)}</option>`)
    .join("");
  rankFilter.value = levels.includes(current) ? current : "all";
}
function getSelectedItems() {
  const checks = Array.from(document.querySelectorAll("#itemFilterMenu .item-filter-check"));
  const selected = checks.filter(ch => ch.checked).map(ch => ch.value);
  if (!checks.length || selected.includes("all") || selected.length === 0) return ["all"];
  return selected;
}

function itemFilterAllows(item, selectedItems = getSelectedItems()) {
  return selectedItems.includes("all") || selectedItems.includes(item);
}

function getAllTargetItems() {
  const items = new Set();
  Object.values(customerTargets || {}).forEach(customer => {
    Object.keys(customer.items || {}).forEach(item => items.add(item));
  });
  return items;
}

function updateItemFilterOptions(visibleItems) {
  const menu = document.getElementById("itemFilterMenu");
  const label = document.getElementById("itemFilterLabel");
  if (!menu) return;
  const previous = getSelectedItems();
  const allItems = getAllTargetItems();
  const sortedItems = Array.from(allItems.size ? allItems : (visibleItems || [])).sort((a, b) => a.localeCompare(b));
  const useAll = previous.includes("all");
  const checkedItems = useAll ? new Set(["all"]) : new Set(previous.filter(item => sortedItems.includes(item)));
  if (!checkedItems.size) checkedItems.add("all");
  const selectedCount = checkedItems.has("all") ? sortedItems.length : checkedItems.size;
  menu.innerHTML = `
    <div class="item-filter-header">
      <div>
        <div class="item-filter-title">Filter Items</div>
        <div class="item-filter-count">${selectedCount} selected</div>
      </div>
      <div class="item-filter-actions">
        <button type="button" data-item-action="all">All</button>
        <button type="button" data-item-action="clear">Clear</button>
      </div>
    </div>
    <input id="itemFilterSearch" class="item-filter-search" type="text" placeholder="Search item..." autocomplete="off">
    <div class="item-filter-list">
      <label class="item-filter-option font-semibold">
        <input type="checkbox" class="item-filter-check" value="all" ${checkedItems.has("all") ? "checked" : ""}>
        <span>All Items</span>
      </label>
      ${sortedItems.map(item => `
        <label class="item-filter-option" data-item-name="${item.toLowerCase()}">
          <input type="checkbox" class="item-filter-check" value="${item}" ${checkedItems.has(item) ? "checked" : ""}>
          <span>${item}</span>
        </label>
      `).join("")}
    </div>
  `;
  if (label) {
    label.textContent = checkedItems.has("all") ? "All Items" : `${checkedItems.size} Items`;
  }
  positionItemFilterMenu();
}

function filterItemDropdownList(searchText = "") {
  const query = searchText.trim().toLowerCase();
  document.querySelectorAll("#itemFilterMenu .item-filter-option[data-item-name]").forEach(option => {
    option.style.display = option.dataset.itemName.includes(query) ? "flex" : "none";
  });
}

function setItemFilterSelection(mode) {
  const checks = Array.from(document.querySelectorAll("#itemFilterMenu .item-filter-check"));
  checks.forEach(ch => {
    ch.checked = mode === "all" ? ch.value === "all" : false;
  });
  const all = checks.find(ch => ch.value === "all");
  if (mode === "clear" && all) all.checked = true;
  renderInvoiceTable();
}

function handleItemFilterChange(event) {
  const target = event.target;
  if (!target?.classList?.contains("item-filter-check")) return;
  const checks = Array.from(document.querySelectorAll("#itemFilterMenu .item-filter-check"));
  if (target.value === "all" && target.checked) {
    checks.forEach(ch => { if (ch.value !== "all") ch.checked = false; });
  } else if (target.value !== "all" && target.checked) {
    const all = checks.find(ch => ch.value === "all");
    if (all) all.checked = false;
  }
  if (!checks.some(ch => ch.checked)) {
    const all = checks.find(ch => ch.value === "all");
    if (all) all.checked = true;
  }
  renderInvoiceTable();
}

function positionItemFilterMenu() {
  const box = document.getElementById("itemFilterBox");
  const label = document.getElementById("itemFilterLabel");
  const menu = document.getElementById("itemFilterMenu");
  if (!box || !label || !menu || !box.open) return;

  const rect = label.getBoundingClientRect();
  const width = Math.min(Math.max(rect.width, 280), window.innerWidth - 16);
  const left = Math.min(Math.max(8, rect.left), window.innerWidth - width - 8);
  const spaceBelow = window.innerHeight - rect.bottom - 12;
  const spaceAbove = rect.top - 12;
  const openAbove = spaceBelow < 180 && spaceAbove > spaceBelow;
  const maxHeight = Math.max(160, Math.min(320, openAbove ? spaceAbove : spaceBelow));

  menu.style.position = "fixed";
  menu.style.left = `${left}px`;
  menu.style.width = `${width}px`;
  menu.style.maxHeight = `${maxHeight}px`;
  menu.style.zIndex = "99999";
  menu.style.top = openAbove ? "auto" : `${rect.bottom + 6}px`;
  menu.style.bottom = openAbove ? `${window.innerHeight - rect.top + 6}px` : "auto";
}

function setupItemFilterDropdownPosition() {
  const box = document.getElementById("itemFilterBox");
  const menu = document.getElementById("itemFilterMenu");
  if (!box) return;
  box.addEventListener("toggle", positionItemFilterMenu);
  if (menu && !menu.dataset.enhanced) {
    menu.dataset.enhanced = "true";
    menu.addEventListener("input", event => {
      if (event.target?.id === "itemFilterSearch") filterItemDropdownList(event.target.value);
    });
    menu.addEventListener("click", event => {
      const action = event.target?.dataset?.itemAction;
      if (!action) return;
      event.preventDefault();
      setItemFilterSelection(action);
    });
  }
  if (window.__itemFilterPositionBound) return;
  window.__itemFilterPositionBound = true;
  window.addEventListener("resize", positionItemFilterMenu);
  window.addEventListener("scroll", positionItemFilterMenu, true);
}

function getCityWiseSummary(statusFilter = "all", selectedItems = ["all"]) {
  const cityMap = {};
  const visibleItems = new Set();
  Object.entries(customerTargets || {}).forEach(([customerCode, customer]) => {
    Object.entries(customer.items || {}).forEach(([item, targetQty]) => {
      if (!itemFilterAllows(item, selectedItems)) return;
      const matchingInvoices = invoices.filter(inv =>
        inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
        inv.item?.toUpperCase() === item.toUpperCase()
      );
      const achievedQty = matchingInvoices.reduce((sum, inv) => sum + Number(inv.quantity || 0), 0);
      const achievedValue = matchingInvoices.reduce((sum, inv) => sum + (Number(inv.quantity || 0) * Number(inv.rate || 0)), 0);
      const targetQtyNum = Number(targetQty) || 0;
      const remainingQty = targetQtyNum - achievedQty;
      let statusType = "normal";
      if (targetQtyNum === 0) statusType = "zeroTarget";
      else if (remainingQty < 0) statusType = "red";
      else if (remainingQty === 0 && achievedQty > 0) statusType = "green";
      if (statusFilter === "red" && statusType !== "red") return;
      if (statusFilter === "green" && statusType !== "green") return;

      const city = (customer.city || "Unknown City").toString().trim() || "Unknown City";
      const key = `${city}||${item}`;
      if (!cityMap[key]) cityMap[key] = { city, item, customers: new Set(), items: 0, target: 0, achieved: 0, remaining: 0, value: 0 };
      cityMap[key].customers.add(customerCode);
      cityMap[key].items += 1;
      cityMap[key].target += targetQtyNum;
      cityMap[key].achieved += achievedQty;
      cityMap[key].remaining += remainingQty;
      cityMap[key].value += achievedValue;
      visibleItems.add(item);
    });
  });
  const rows = Object.values(cityMap).sort((a, b) => a.city.localeCompare(b.city) || a.item.localeCompare(b.item));
  const totals = rows.reduce((acc, row) => {
    row.customers.forEach(code => acc.customers.add(code));
    acc.items += row.items;
    acc.target += row.target;
    acc.achieved += row.achieved;
    acc.remaining += row.remaining;
    acc.value += row.value;
    return acc;
  }, { customers: new Set(), items: 0, target: 0, achieved: 0, remaining: 0, value: 0 });
  return { rows, totals, visibleItems };
}

function getCityWisePivot(report) {
  const cities = [...new Set(report.rows.map(row => row.city))].sort((a, b) => a.localeCompare(b));
  const itemMap = {};
  report.rows.forEach(row => {
    if (!itemMap[row.item]) itemMap[row.item] = { item: row.item, cities: {}, totalQty: 0 };
    itemMap[row.item].cities[row.city] = (itemMap[row.item].cities[row.city] || 0) + row.achieved;
    itemMap[row.item].totalQty += row.achieved;
  });
  const items = Object.values(itemMap).sort((a, b) => a.item.localeCompare(b.item));
  const cityValueTotals = {};
  cities.forEach(city => cityValueTotals[city] = 0);
  report.rows.forEach(row => {
    cityValueTotals[row.city] = (cityValueTotals[row.city] || 0) + row.value;
  });
  const grandValueTotal = Object.values(cityValueTotals).reduce((sum, value) => sum + value, 0);
  return { cities, items, cityValueTotals, grandValueTotal };
}

function renderCityWisePivotRows(report) {
  const pivot = getCityWisePivot(report);
  if (!pivot.items.length) return '<tr><td colspan="9" class="p-2 text-center">No city wise data available.</td></tr>';
  const colCount = pivot.cities.length + 2;
  let html = `
    <tr class="bg-emerald-100 font-bold text-xs sm:text-sm sticky top-0 z-10">
      <td class="border p-2">ITEM NAME</td>
      ${pivot.cities.map(city => `<td class="border p-2 text-right">${city}</td>`).join("")}
      <td class="border p-2 text-right">TOTAL</td>
    </tr>`;
  html += pivot.items.map(row => `
    <tr class="bg-white hover:bg-blue-50 text-xs sm:text-sm">
      <td class="border p-2 font-semibold">${row.item}</td>
      ${pivot.cities.map(city => `<td class="border p-2 text-right">${(row.cities[city] || 0).toLocaleString()}</td>`).join("")}
      <td class="border p-2 text-right font-bold">${row.totalQty.toLocaleString()}</td>
    </tr>`).join("");
  html += `
    <tr class="bg-indigo-100 font-bold text-xs sm:text-sm">
      <td class="border p-2">VALUES</td>
      ${pivot.cities.map(city => `<td class="border p-2 text-right">${pivot.cityValueTotals[city].toLocaleString()}</td>`).join("")}
      <td class="border p-2 text-right">${pivot.grandValueTotal.toLocaleString()}</td>
    </tr>`;
  return html;
}

function renderCityWisePivotHead(report) {
  const pivot = getCityWisePivot(report);
  return `
    <thead class="bg-emerald-100 sticky top-0 z-40">
      <tr>
        <th class="border p-2">ITEM NAME</th>
        ${pivot.cities.map(city => `<th class="border p-2 text-right">${city}</th>`).join("")}
        <th class="border p-2 text-right">TOTAL</th>
      </tr>
    </thead>`;
}

function renderCityWiseRows(report, label = "City Wise") {
  if (!report.rows.length) return '<tr><td colspan="9" class="p-2 text-center">No city wise data available.</td></tr>';
  let html = report.rows.map(row => `
    <tr class="bg-blue-50 hover:bg-blue-100 transition text-xs sm:text-sm">
      <td class="border p-1 sm:p-2 font-semibold">${row.city}</td>
      <td class="border p-1 sm:p-2">${row.customers.size}</td>
      <td class="border p-1 sm:p-2 font-semibold">${row.item}</td>
      <td class="border p-1 sm:p-2">${row.items}</td>
      <td class="border p-1 sm:p-2">${row.target.toLocaleString()}</td>
      <td class="border p-1 sm:p-2">${row.achieved.toLocaleString()}</td>
      <td class="border p-1 sm:p-2">${row.remaining.toLocaleString()}</td>
      <td class="border p-1 sm:p-2 font-bold">${row.target > 0 ? ((row.achieved / row.target) * 100).toFixed(1) : 0}%</td>
      <td class="border p-1 sm:p-2 font-bold">${row.value.toLocaleString()}</td>
    </tr>`).join("");
  html += `
    <tr class="bg-indigo-100 font-bold text-xs sm:text-sm">
      <td class="border p-2">TOTAL</td>
      <td class="border p-2">${report.totals.customers.size}</td>
      <td class="border p-2">${label}</td>
      <td class="border p-2">${report.totals.items}</td>
      <td class="border p-2">${report.totals.target.toLocaleString()}</td>
      <td class="border p-2">${report.totals.achieved.toLocaleString()}</td>
      <td class="border p-2">${report.totals.remaining.toLocaleString()}</td>
      <td class="border p-2">${report.totals.target > 0 ? ((report.totals.achieved / report.totals.target) * 100).toFixed(1) : 0}%</td>
      <td class="border p-2">${report.totals.value.toLocaleString()}</td>
    </tr>`;
  return html;
}

function handleSidebarClick(event) {
    const buttonId = event.target.id;
    console.log('Sidebar button clicked:', buttonId);
    const sidebar = document.getElementById('sidebar');
    if (sidebar && window.innerWidth < 768) {
        sidebar.classList.add('hidden', '-translate-x-full');
    }
    if (buttonId === 'navInvoiceEntry') {
        showMainPage();
    } else if (buttonId === 'navAllocation') {
        showAllocationPage();
    } else if (buttonId === 'navDoneTargets') {
        showDoneTargetPage();
    } else if (buttonId === 'navBonus') {
        showBonusPage();
    } else if (buttonId === 'navLogout') {
        logout();
    }
}

function initHamburger() {
    const hamburger = document.getElementById('hamburger');
    if (hamburger) {
        hamburger.removeEventListener('click', toggleSidebar);
        hamburger.addEventListener('click', toggleSidebar);
    }
}

function showMainPage() {
    document.getElementById('mainPage').classList.remove('hidden');
    document.getElementById('allocationPage').classList.add('hidden');
    document.getElementById('doneTargetPage').classList.add('hidden');
    document.getElementById('bonusPage').classList.add('hidden');
    document.getElementById('mySalePage').classList.add('hidden');

    document.getElementById('navInvoiceEntry').classList.add('bg-primary', 'text-white');
    document.getElementById('navAllocation').classList.remove('bg-primary', 'text-white');
    document.getElementById('navDoneTargets').classList.remove('bg-primary', 'text-white');
    document.getElementById('navBonus').classList.remove('bg-primary', 'text-white');
    document.getElementById('navMySale')?.classList.remove('bg-yellow-600', 'text-white');

    renderInvoiceTable();
}

function showAllocationPage() {
    document.getElementById('mainPage').classList.add('hidden');
    document.getElementById('allocationPage').classList.remove('hidden');
    document.getElementById('doneTargetPage').classList.add('hidden');
    document.getElementById('bonusPage').classList.add('hidden');
    document.getElementById('mySalePage').classList.add('hidden');

    document.getElementById('navInvoiceEntry').classList.remove('bg-primary', 'text-white');
    document.getElementById('navAllocation').classList.add('bg-primary', 'text-white');
    document.getElementById('navDoneTargets').classList.remove('bg-primary', 'text-white');
    document.getElementById('navBonus').classList.remove('bg-primary', 'text-white');
    document.getElementById('navMySale')?.classList.remove('bg-yellow-600', 'text-white');

    const tablesContainer = document.getElementById('allocationTables');
    if (tablesContainer) {
        tablesContainer.innerHTML = '<p class="text-center text-gray-500">Please search for a customer to view report.</p>';
        lastRenderedCustomerCode = null;
    }
    console.log('Allocation page shown, allocation tables cleared');
}

function showDoneTargetPage() {
    document.getElementById('mainPage').classList.add('hidden');
    document.getElementById('allocationPage').classList.add('hidden');
    document.getElementById('doneTargetPage').classList.remove('hidden');
    document.getElementById('bonusPage').classList.add('hidden');
    document.getElementById('mySalePage').classList.add('hidden');

    document.getElementById('navInvoiceEntry').classList.remove('bg-primary', 'text-white');
    document.getElementById('navAllocation').classList.remove('bg-primary', 'text-white');
    document.getElementById('navDoneTargets').classList.add('bg-primary', 'text-white');
    document.getElementById('navBonus').classList.remove('bg-primary', 'text-white');
    document.getElementById('navMySale')?.classList.remove('bg-yellow-600', 'text-white');

    renderDoneTargetTables();
}

function showBonusPage() {
    document.getElementById('mainPage').classList.add('hidden');
    document.getElementById('allocationPage').classList.add('hidden');
    document.getElementById('doneTargetPage').classList.add('hidden');
    document.getElementById('bonusPage').classList.remove('hidden');
    document.getElementById('mySalePage').classList.add('hidden');

    document.getElementById('navInvoiceEntry').classList.remove('bg-primary', 'text-white');
    document.getElementById('navAllocation').classList.remove('bg-primary', 'text-white');
    document.getElementById('navDoneTargets').classList.remove('bg-primary', 'text-white');
    document.getElementById('navBonus').classList.add('bg-primary', 'text-white');
    document.getElementById('navMySale')?.classList.remove('bg-yellow-600', 'text-white');

    renderBonusDeals();
}

function showMySalePage() {
    document.getElementById('mainPage').classList.add('hidden');
    document.getElementById('allocationPage').classList.add('hidden');
    document.getElementById('doneTargetPage').classList.add('hidden');
    document.getElementById('bonusPage').classList.add('hidden');
    document.getElementById('mySalePage').classList.remove('hidden');

    document.getElementById('navInvoiceEntry').classList.remove('bg-primary', 'text-white');
    document.getElementById('navAllocation').classList.remove('bg-primary', 'text-white');
    document.getElementById('navDoneTargets').classList.remove('bg-primary', 'text-white');
    document.getElementById('navBonus').classList.remove('bg-primary', 'text-white');
    document.getElementById('navMySale')?.classList.add('bg-yellow-600', 'text-white');

    renderMySaleTable();
}


function login() {
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const loginError = document.getElementById('loginError');
    const loginPage = document.getElementById('loginPage');
    const mainPage = document.getElementById('mainPage');
    const sidebar = document.getElementById('sidebar');
    const hamburgerContainer = document.getElementById('hamburgerContainer');

    if (!usernameInput || !passwordInput || !loginError || !loginPage || !mainPage || !sidebar || !hamburgerContainer) {
        console.error('Login DOM elements missing');
        return;
    }

    const username = usernameInput.value.trim();
    const password = passwordInput.value;

    if (!username || !password) {
        loginError.classList.remove('hidden');
        loginError.textContent = 'Please enter both username and password.';
        return;
    }

    if (demoLogin(username, password)) {
        isLoggedIn = true;
        localStorage.setItem('isLoggedIn', 'true');
        loginPage.classList.add('hidden');
        mainPage.classList.remove('hidden');
        sidebar.classList.remove('hidden');
        sidebar.classList.add('md:block', '-translate-x-full');
        hamburgerContainer.classList.remove('hidden');
        loginError.classList.add('hidden');
        usernameInput.value = '';
        passwordInput.value = '';
        initSidebarNav();
        renderInvoiceTable();
    } else {
        loginError.classList.remove('hidden');
        loginError.textContent = 'Invalid credentials!';
    }
}

function logout() {
    isLoggedIn = false;
    localStorage.setItem('isLoggedIn', 'false');
    logoutDemo();
}

function autoFillCity() {
    const customerInput = document.getElementById('customer');
    const cityInput = document.getElementById('city');
    const suggestionsDiv = document.getElementById('customerSuggestions');
    if (!customerInput || !cityInput || !suggestionsDiv) return;

    const query = customerInput.value.trim().toLowerCase();
    suggestionsDiv.innerHTML = '';
    suggestionsDiv.classList.add('hidden');

    if (!query) {
        cityInput.value = '';
        return;
    }

    const filteredCustomers = customers.filter(c => 
        c.name.toLowerCase().includes(query) || 
        c.code.toLowerCase().includes(query)
    );

    if (filteredCustomers.length > 0) {
        suggestionsDiv.classList.remove('hidden');
        filteredCustomers.forEach(customer => {
            const suggestion = document.createElement('div');
            suggestion.className = 'p-2 hover:bg-teal-500 cursor-pointer';
            suggestion.textContent = `${customer.name} (${customer.code}) - ${customer.city}`;
            suggestion.addEventListener('click', () => {
                customerInput.value = `${customer.name} (${customer.code})`;
                cityInput.value = customer.city;
                suggestionsDiv.classList.add('hidden');
                document.getElementById('item').value = '';
                document.getElementById('target').value = '';
                document.getElementById('remaining').value = '';
                document.getElementById('itemSuggestions').classList.add('hidden');
            });
            suggestionsDiv.appendChild(suggestion);
        });
    } else {
        cityInput.value = '';
    }
}

function addInvoice() {
    const customerInput = document.getElementById('customer')?.value.trim();
    const itemInput = document.getElementById('item')?.value.trim();
    const quantityInput = document.getElementById('quantity')?.value.trim();
    const cityInput = document.getElementById('city')?.value.trim();
    const errorDiv = document.getElementById('invoiceError');

    if (!customerInput || !itemInput || !quantityInput || !cityInput || !errorDiv) {
        console.error('Invoice input fields missing or invalid:', { customerInput, itemInput, quantityInput, cityInput });
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Please fill all fields.';
        return;
    }

    const customerMatch = customerInput.match(/(.+)\s*\((.+)\)/);
    if (!customerMatch) {
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Invalid customer format. Use: Name (Code)';
        console.error('Invalid customer format:', customerInput);
        return;
    }

    const customerCode = customerMatch[2].trim().toUpperCase();
    const customerName = customerMatch[1].trim();
    const quantity = Number(quantityInput);
    const item = itemInput.trim().toUpperCase();

    if (isNaN(quantity) || quantity <= 0) {
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Invalid quantity.';
        console.error('Invalid quantity:', quantityInput);
        return;
    }

    if (!customerTargets[customerCode]) {
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Customer not found in targets.';
        console.error('Customer not found:', customerCode);
        return;
    }

    if (!customerTargets[customerCode].items[item]) {
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Item not found for this customer.';
        console.error('Item not found for customerCode:', customerCode, 'item:', item);
        return;
    }

    const target = Number(customerTargets[customerCode].items[item] || 0);
    const achieved = invoices
        .filter(inv => inv && inv.customerCode?.toUpperCase() === customerCode && inv.item?.toUpperCase() === item && !isNaN(Number(inv.quantity)))
        .reduce((sum, inv) => sum + Number(inv.quantity), 0);
    const remaining = target - (achieved + quantity);

    if (doneTargets.some(dt => dt.customerCode?.toUpperCase() === customerCode && dt.item?.toUpperCase() === item)) {
        errorDiv.classList.remove('hidden');
        errorDiv.textContent = 'Target completed, cannot create new invoice for this customer and item.';
        console.error('Target completed for customerCode:', customerCode, 'item:', item);
        return;
    }

    if (remaining < 0 && !confirm('Adding this quantity will make Remaining negative. Proceed?')) {
        console.log('User cancelled invoice addition due to negative remaining:', remaining);
        return;
    }

    const newInvoice = {
        city: cityInput,
        customerCode: customerCode,
        customer: customerName,
        item: item,
        quantity: quantity,
        user: getLoggedUser() || ''
    };
    invoices.push(newInvoice);
    localStorage.setItem('invoices', JSON.stringify(invoices));
    console.log('Added invoice:', newInvoice);
    console.log('Current invoices:', invoices);
    document.getElementById('quantity').value = '';
    document.getElementById('remaining').value = String(remaining);
    errorDiv.classList.add('hidden');
    renderInvoiceTable();
    renderAllocationTables();
    syncManualAchievementToFirebase(newInvoice, target).catch(err => {
        console.error('Manual achievement Firebase sync failed:', err);
    });
}

async function syncManualAchievementToFirebase(invoice, targetQty) {
    if (!invoice || !invoice.customerCode || !invoice.item) return;
    const cleanUser = (getActiveDataUser() || getLoggedUser() || "").toString().trim().toUpperCase();
    const rowUser = cleanUser && cleanUser !== "ALL" && cleanUser !== "ADMIN"
        ? cleanUser
        : (getLoggedUser() || "").toString().trim().toUpperCase();
    const today = new Date().toISOString().slice(0, 10);
    let updated = false;

    excelData = (excelData || []).map(row => {
        const sameCustomer = (row.CustomerCode || "").toString().trim().toUpperCase() === invoice.customerCode;
        const sameItem = (row.Item1 || "").toString().trim().toUpperCase() === invoice.item;
        const users = [row.User1, row.User2].map(user => (user || "").toString().trim().toUpperCase()).filter(Boolean);
        const sameUser = !rowUser || rowUser === "ADMIN" || users.length === 0 || users.includes(rowUser);
        if (!updated && sameCustomer && sameItem && sameUser) {
            updated = true;
            return {
                ...row,
                City: row.City || invoice.city || "",
                Customer: row.Customer || invoice.customer || "",
                Target1: Number(row.Target1 || targetQty || 0),
                Achieve1: Number(row.Achieve1 || 0) + Number(invoice.quantity || 0),
                User1: row.User1 || rowUser,
                Date: today
            };
        }
        return row;
    });

    if (!updated) {
        excelData.push({
            City: invoice.city || "",
            CustomerCode: invoice.customerCode,
            Customer: invoice.customer || "",
            Item1: invoice.item,
            Target1: Number(targetQty || 0),
            Achieve1: Number(invoice.quantity || 0),
            User1: rowUser,
            User2: "",
            DealQty: 0,
            DealBonus: 0,
            SummaryNumber: "",
            CompanyName: "",
            Value: 0,
            Date: today,
            ItemRate: 0
        });
    }

    localStorage.setItem("excelData", JSON.stringify(excelData));
    if (cleanUser === "ALL") localStorage.setItem("excelDataAll", JSON.stringify(excelData));
    if (typeof saveCSVToFirebase === "function") await saveCSVToFirebase(excelData);
}

function renderInvoiceTable() {
    const tbody = document.getElementById('invoiceTableBody');
    const thead = document.getElementById('invoiceTableHead');
    if (!tbody) {
        console.error('invoiceTableBody element not found');
        return;
    }

    // --- Add filters only once ---
    if (thead && !document.getElementById("statusFilter")) {
        const filterRow = document.createElement("tr");
        filterRow.innerHTML = `
            <th colspan="9" class="p-0">
                <div class="filter-toolbar sticky top-0 z-20">
                    <div class="filter-toolbar-inner">
                        <div class="filter-control">
                            <label>Filter by Status</label>
                            <select id="statusFilter">
                                <option value="all">🌍 All</option>
                                <option value="green">✅ Completed</option>
                                <option value="cityWiseGreen">City Wise Completed</option>
                                <option value="red">🔴 Red Zone</option>
                                <option value="cityWiseRed">City Wise Red Zone</option>
                                <option value="normal">⏳ Pending</option>
                                <option value="zeroTarget">Zero Target</option>
                                <option value="nonProductive">🚫 Non Productive</option>
                                <option value="top10">🏆 Top 10 Customers</option>
                                <option value="cityWise">City Wise Report</option>
                                <option value="itemSummary">📊 Item Summary</option>
                                 <option value="nonProductiveItemSummary">🚫 Non Productive Item</option>
                            </select>
                        </div>
                        <div class="filter-control zero-target-control">
                            <label>Apply All Zero Target</label>
                            <div class="filter-inline">
                                <input id="zeroTargetApplyAllValue" type="number" min="1" placeholder="Target">
                                <button type="button" onclick="applyTargetToAllZeroItems()">Apply All</button>
                            </div>
                        </div>
                        <div class="filter-control item-filter-control">
                            <label>Filter by Item</label>
                            <details id="itemFilterBox">
                                <summary id="itemFilterLabel">All Items</summary>
                                <div id="itemFilterMenu">
                                    <label class="flex items-center gap-2 px-2 py-1 font-semibold"><input type="checkbox" class="item-filter-check" value="all" checked> All Items</label>
                                </div>
                            </details>
                        </div>
                        <div class="filter-control">
                            <label>Filter by Rank</label>
                            <select id="rankFilter">
                                <option value="all">All Ranks</option>
                            </select>
                        </div>
                    </div>
                </div>
            </th>
        `;
        thead.prepend(filterRow);

        document.getElementById("statusFilter").addEventListener("change", renderInvoiceTable);
        document.getElementById("itemFilterMenu").addEventListener("change", handleItemFilterChange);
        setupItemFilterDropdownPosition();
        document.getElementById("rankFilter").addEventListener("change", renderInvoiceTable);
    }

    const selectedFilter = document.getElementById("statusFilter")?.value || "all";
    const selectedItems = getSelectedItems();
    let selectedRank = document.getElementById("rankFilter")?.value || "all";

    const rankedCustomers = getCustomerRankings();
    populateRankFilter(rankedCustomers);
    selectedRank = document.getElementById("rankFilter")?.value || "all";

    // --- Top 10 customers by totalTarget (QTY) ---
    const customerTotals = Object.entries(customerTargets).map(([code, cust]) => {
        const totalTargetQty = Object.values(cust.items).reduce((a, b) => a + Number(b), 0);
        return { code, name: cust.name || code, totalTargetQty };
    });
    const top10Customers = customerTotals.sort((a,b)=>b.totalTargetQty-a.totalTargetQty).slice(0,10).map(c=>c.code);

    let rowsHtml = '';
    let visibleItems = new Set();
    let zeroAchieveCustomers = [];

    // --- Summary counters ---
    let totalCustomers=0, nonProductive=0, completed=0, progress=0;
    let overallAchievedValue=0, overallTargetValue=0, overallRemainingValue=0;
    let overallAchievedQty=0, overallTargetQty=0;

    const customerShades = ["bg-gray-50","bg-blue-50","bg-purple-50","bg-pink-50","bg-yellow-50","bg-teal-50"];
    let customerIndex = 0;

    // --- Item summary (QTY-based) ---
    let itemSummary = {};
    Object.entries(customerTargets).forEach(([customerCode, customer]) => {
        Object.entries(customer.items).forEach(([item, targetQty]) => {
            if (!itemSummary[item]) itemSummary[item] = { totalTargetQty:0, totalAchievedQty:0, totalRemainingQty:0, totalValue:0, customerCount:0, achievedCustomerCount:0 };

            const matchingInvoices = invoices.filter(inv =>
                inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                inv.item?.toUpperCase() === item.toUpperCase()
            );

            const achievedQty = matchingInvoices.reduce((sum, inv) => sum + Number(inv.quantity || 0), 0);
            const achievedValue = matchingInvoices.reduce((sum, inv) => sum + (Number(inv.quantity || 0) * Number(inv.rate || 0)), 0);

            itemSummary[item].totalTargetQty += Number(targetQty);
            itemSummary[item].totalAchievedQty += achievedQty;
            itemSummary[item].totalRemainingQty += (Number(targetQty) - achievedQty);
            itemSummary[item].totalValue += achievedValue;
            itemSummary[item].customerCount += 1;
            if (achievedQty > 0) itemSummary[item].achievedCustomerCount += 1;
        });
    });

    if (selectedFilter === "itemSummary" || selectedFilter === "nonProductiveItemSummary") {

        // --- Render Item Summary Table (QTY-based) ---
        Object.entries(itemSummary).sort((a, b) => a[0].localeCompare(b[0])).forEach(([item, data]) => {
           // 🚫 Non Productive Item Summary ONLY
if (
    selectedFilter === "nonProductiveItemSummary" &&
    data.totalTargetQty > 0 &&
    data.totalAchievedQty > 0
) return;

            if (!itemFilterAllows(item, selectedItems)) return;
            const perc = data.totalTargetQty>0?((data.totalAchievedQty/data.totalTargetQty)*100).toFixed(1):0;
            let rowClass = "bg-gray-50";
            if(data.totalRemainingQty<0) rowClass="bg-red-500 text-white";
            else if(data.totalRemainingQty===0 && data.totalAchievedQty>0) rowClass="bg-green-500 text-white";

            rowsHtml += `<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
                <td class="border p-1 sm:p-2"></td>
                <td class="border p-1 sm:p-2">${data.achievedCustomerCount} Productivity</td>
                <td class="border p-1 sm:p-2">${data.customerCount} Customers</td>
                <td class="border p-1 sm:p-2">${item}</td>
                <td class="border p-1 sm:p-2">${data.totalTargetQty.toLocaleString()}</td>
                <td class="border p-1 sm:p-2">${data.totalAchievedQty.toLocaleString()}</td>
                <td class="border p-1 sm:p-2">${data.totalRemainingQty.toLocaleString()}</td>
                <td class="border p-1 sm:p-2 font-bold">${perc}%</td>
                <td class="border p-1 sm:p-2 font-bold">${data.totalValue.toLocaleString()}</td>
            </tr>`;
            visibleItems.add(item);
        });
    } else if (selectedFilter === "cityWise" || selectedFilter === "cityWiseRed" || selectedFilter === "cityWiseGreen") {
        const cityStatus = selectedFilter === "cityWiseRed" ? "red" : (selectedFilter === "cityWiseGreen" ? "green" : "all");
        const cityLabel = selectedFilter === "cityWiseRed" ? "City Wise Red Zone" : (selectedFilter === "cityWiseGreen" ? "City Wise Completed" : "City Wise");
        const cityReport = getCityWiseSummary(cityStatus, selectedItems);
        cityReport.visibleItems.forEach(item => visibleItems.add(item));
        rowsHtml = renderCityWisePivotRows(cityReport);
    } else {
        // --- Customer Table Rendering (QTY-based) ---
        Object.entries(customerTargets).forEach(([customerCode, customer]) => {
            // --- Apply rank filter ---
            const rankInfo = rankedCustomers.find(c => c.code === customerCode);
            if (!rankInfo) {
                console.warn(`No rank info found for customer: ${customerCode}`);
                return;
            }
            if (selectedRank !== "all" && rankInfo.level !== selectedRank) return;

            if(selectedFilter==="top10" && !top10Customers.includes(customerCode)) return;

            totalCustomers++;
            const customerShade = customerShades[customerIndex % customerShades.length];
            customerIndex++;

            let allCompleted=true, anyAchieved=false;

            Object.entries(customer.items).forEach(([item, targetQty]) => {
                if(!itemFilterAllows(item, selectedItems)) return;

                const matchingInvoices = invoices.filter(inv =>
                    inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                    inv.item?.toUpperCase() === item.toUpperCase()
                );

                const achievedQty = matchingInvoices.reduce((sum, inv)=>sum+Number(inv.quantity||0),0);
                const achievedValue = matchingInvoices.reduce((sum, inv)=>sum+(Number(inv.quantity||0)*Number(inv.rate||0)),0);

                let avgRate = 0;
                if(matchingInvoices.length>0){
                    avgRate = matchingInvoices.reduce((s,inv)=>s+Number(inv.rate||0),0) / matchingInvoices.length;
                }

                const targetQtyNum = Number(targetQty);
                const remainingQty = targetQtyNum - achievedQty;
                const perc = targetQtyNum>0?((achievedQty/targetQtyNum)*100).toFixed(1):0;

                overallAchievedQty += achievedQty;
                overallTargetQty += targetQtyNum;

                const targetValue = targetQtyNum * avgRate;
                const remainingValue = targetValue - achievedValue;

                overallAchievedValue += achievedValue;
                overallTargetValue += targetValue;
                overallRemainingValue += remainingValue;

                if(achievedQty<targetQtyNum) allCompleted=false;
                if(achievedQty>0) anyAchieved=true;

                let rowClass = customerShade;
                let statusType="normal";
                if(targetQtyNum === 0){ rowClass="bg-orange-100"; statusType="zeroTarget"; }
                else if(remainingQty<0){ rowClass="bg-red-500 text-white"; statusType="red"; }
                else if(remainingQty===0 && achievedQty>0){ rowClass="bg-green-500 text-white"; statusType="green"; }

                // --- Non-Productive Filter ---
if (selectedFilter === "nonProductive" && anyAchieved) return;

// --- Other Status Filters ---
if (
    selectedFilter !== "all" &&
    selectedFilter !== "top10" &&
    selectedFilter !== "nonProductive" &&     // allow nonProductive
    selectedFilter !== statusType
) return;

                const targetCell = targetQtyNum === 0
                    ? `<div class="flex flex-col sm:flex-row gap-1"><input id="${safeTargetInputId(customerCode, item)}" type="number" min="1" class="border rounded px-2 py-1 w-24 text-black" placeholder="Target"><button onclick="setTargetForZeroItem('${String(customerCode).replace(/'/g, "\\'")}', '${String(item).replace(/'/g, "\\'")}')" class="bg-orange-500 hover:bg-orange-600 text-white px-2 py-1 rounded">Set</button></div>`
                    : targetQtyNum.toLocaleString();

                visibleItems.add(item);

                rowsHtml+=`<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
                    <td class="border p-1 sm:p-2">${customer.city||''}</td>
                    <td class="border p-1 sm:p-2">${customerCode}</td>
                    <td class="border p-1 sm:p-2">${customer.name||''} (${rankInfo.displayLevel || rankInfo.level})</td>
                    <td class="border p-1 sm:p-2">${item}</td>
                    <td class="border p-1 sm:p-2">${targetCell}</td>
                    <td class="border p-1 sm:p-2">${achievedQty.toLocaleString()}</td>
                    <td class="border p-1 sm:p-2">${remainingQty.toLocaleString()}</td>
                    <td class="border p-1 sm:p-2 font-bold">${perc}%</td>
                    <td class="border p-1 sm:p-2 font-bold">${achievedValue.toLocaleString()}</td>
                </tr>`;
            });

            if(!anyAchieved) zeroAchieveCustomers.push({name:customer.name?.trim()||customerCode, code:customerCode});
            if(!anyAchieved) nonProductive++;
            else if(allCompleted) completed++;
            else progress++;
        });
    }

    if(!rowsHtml) rowsHtml='<tr><td colspan="9" class="p-2 text-center">No invoices available.</td></tr>';
    tbody.innerHTML = rowsHtml;

    // --- Summary boxes (Dashboard) ---
    document.getElementById("totalCustomersBox").lastElementChild.innerText = totalCustomers;
    document.getElementById("nonProductiveBox").lastElementChild.innerText = nonProductive;
    document.getElementById("progressBox").lastElementChild.innerText = progress;
    document.getElementById("completedBox").lastElementChild.innerText = completed;

  // ✅ Corrected Overall % calculation (based on quantities)
const smartOverall = calculateSmartPerformance();
document.getElementById("overallBox").lastElementChild.innerText =
    smartOverall + "% ";


    // --- Value Toggle System (Dashboard only) ---
    const totalValueBox = document.getElementById("totalValueBox").lastElementChild;
    window.totalValueData = { 
        achieved: overallAchievedValue,
        target: overallTargetValue,
        remaining: overallRemainingValue
    };
    if(!window.valueBoxState){ window.valueBoxState = 0; }

    const updateValueBox = ()=>{
        if(window.valueBoxState===0){
            totalValueBox.innerText = window.totalValueData.achieved.toLocaleString()+" (Achieved)";
            totalValueBox.style.color="green";
        } else if(window.valueBoxState===1){
            totalValueBox.innerText = window.totalValueData.target.toLocaleString()+" (Target)";
            totalValueBox.style.color="blue";
        } else {
            totalValueBox.innerText = window.totalValueData.remaining.toLocaleString()+" (Remaining)";
            totalValueBox.style.color="orange";
        }
    };
    updateValueBox();

    const totalValueBoxParent = document.getElementById("totalValueBox");
    if(totalValueBoxParent){
        totalValueBoxParent.onclick = ()=>{
            window.valueBoxState = (window.valueBoxState+1)%3;
            updateValueBox();
        };
    }    // --- Update Item Filter ---
    updateItemFilterOptions(selectedFilter === "nonProductive" ? new Set() : visibleItems);

   // --- Breaking News
const breakingNews = document.getElementById("breakingNews");
if (breakingNews) {
    if (zeroAchieveCustomers.length > 0) {
        breakingNews.innerHTML = `
            <marquee behavior="scroll" direction="left" scrollamount="5" class="flex items-center h-full">
                ${zeroAchieveCustomers.map(customer => {
                    // Rank dhoond lo (pehle se rankedCustomers mojood hai)
                    const rankInfo = rankedCustomers.find(rc => rc.code === customer.code);
                    const level = rankInfo?.displayLevel || rankInfo?.level || "Unknown";

                    return `
                        <span class="
                            inline-flex items-center 
                            mx-3 px-4 py-1.5 
                            bg-red-700 text-white font-bold 
                            rounded-full 
                            shadow-lg shadow-red-900/60 
                            border-2 border-yellow-300 
                            ring-1 ring-yellow-200/50 
                            hover:bg-red-800 transition-all duration-200 
                            cursor-pointer whitespace-nowrap
                        " onclick="openCustomerPopup('${customer.code}')">
                            🚨 ${customer.name} (${customer.code}) - ${level}
                        </span>
                    `;
                }).join('')}
            </marquee>
        `;
        
        // Container ko allocation jaisa gradient + styling do
        breakingNews.className = `
            relative overflow-hidden h-12 font-semibold text-sm 
            rounded-xl shadow-xl mb-6 
            bg-gradient-to-r from-red-600 via-yellow-400 to-red-600 
            border-2 border-red-700
        `;
    } 
    else {
        breakingNews.innerHTML = `
            <div class="flex items-center justify-center h-full text-gray-900 font-medium">
                No alerts at this time
            </div>
        `;
        
        breakingNews.className = `
            relative overflow-hidden h-12 font-semibold text-sm 
            rounded-xl shadow-xl mb-6 
            bg-gradient-to-r from-red-600 via-yellow-400 to-red-600 
            border-2 border-red-700
        `;
    }
}
}





// --- Popup function ---
function showFilteredPopup() {
    const selectedStatus = document.getElementById("statusFilter").value;
    const selectedPopupItems = getSelectedItems();

    // --- Compute Top 10 Customers by totalTarget ---
    let customerTotals = Object.entries(customerTargets).map(([code, cust]) => {
        let totalTarget = Object.values(cust.items).reduce((a, b) => a + Number(b), 0);
        return { code, name: cust.name || code, totalTarget };
    });

    let top10Customers = customerTotals
        .sort((a, b) => b.totalTarget - a.totalTarget)
        .slice(0, 10)
        .map(c => c.code);

    let popupRows = '';
    const customerShades = ["bg-gray-50", "bg-blue-50", "bg-purple-50", "bg-pink-50", "bg-yellow-50", "bg-teal-50", "bg-orange-50"];
    let customerIndex = 0;

    // --- Totals ---
    let totalCustomers = 0;
    let totalItems = 0;
    let totalTarget = 0;
    let totalAchieved = 0;
    let totalRemaining = 0;
    let totalValue = 0;
    let totalAchievedCustomers = 0;
    const citySummary = {};

    let popupThead = ""; // dynamic header

   if (selectedStatus === "cityWise" || selectedStatus === "cityWiseRed" || selectedStatus === "cityWiseGreen") {
        const cityStatus = selectedStatus === "cityWiseRed" ? "red" : (selectedStatus === "cityWiseGreen" ? "green" : "all");
        const cityLabel = selectedStatus === "cityWiseRed" ? "City Wise Red Zone" : (selectedStatus === "cityWiseGreen" ? "City Wise Completed" : "City Wise");
        const cityReport = getCityWiseSummary(cityStatus, selectedPopupItems);
        popupThead = renderCityWisePivotHead(cityReport);
        popupRows = renderCityWisePivotRows(cityReport);
    } else if (
    selectedStatus === "itemSummary" ||
    selectedStatus === "nonProductiveItemSummary"
) {

        // --- Item-based summary for popup ---
        let itemSummary = {};
        Object.entries(customerTargets).forEach(([customerCode, customer]) => {
            Object.entries(customer.items).forEach(([item, target]) => {
                if (!itemSummary[item]) {
                    itemSummary[item] = { totalTarget: 0, totalAchieved: 0, totalRemaining: 0, totalValue: 0, customerCount: 0, achievedCustomerCount: 0 };
                }
                const achieved = invoices
                    .filter(inv =>
                        inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                        inv.item?.toUpperCase() === item.toUpperCase() &&
                        !isNaN(Number(inv.quantity))
                    )
                    .reduce((sum, inv) => sum + Number(inv.quantity), 0);

                const value = invoices
                    .filter(inv =>
                        inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                        inv.item?.toUpperCase() === item.toUpperCase()
                    )
                    .reduce((sum, inv) => sum + (Number(inv.quantity || 0) * Number(inv.rate || 0)), 0);

                itemSummary[item].totalTarget += Number(target);
                itemSummary[item].totalAchieved += achieved;
                itemSummary[item].totalRemaining += Number(target) - achieved;
                itemSummary[item].totalValue += value;
                itemSummary[item].customerCount += 1;
                if (achieved > 0) itemSummary[item].achievedCustomerCount += 1;
            });
        });

        popupThead = `
            <thead class="bg-gray-100 sticky top-0 z-40">
                <tr>
                    <th class="border p-2">Item</th>
                    <th class="border p-2">Customers</th>
                    <th class="border p-2">Productivity</th>
                    <th class="border p-2">Target</th>
                    <th class="border p-2">Achieved</th>
                    <th class="border p-2">Remaining</th>
                    <th class="border p-2">%</th>
                    <th class="border p-2">Value</th>
                </tr>
            </thead>
        `;

        Object.entries(itemSummary).sort((a, b) => a[0].localeCompare(b[0])).forEach(([item, data]) => {
           // 🚫 Non Productive Item Summary (Popup)
if (
    selectedStatus === "nonProductiveItemSummary" &&
    data.totalTarget > 0 &&
    data.totalAchieved > 0
) return;

            if (!itemFilterAllows(item, selectedPopupItems)) return;

            const percentage = data.totalTarget > 0 ? ((data.totalAchieved / data.totalTarget) * 100).toFixed(1) : 0;
            let rowClass = customerShades[customerIndex % customerShades.length];
            if (data.totalRemaining < 0) {
                rowClass = "bg-red-500 text-white";
            } else if (data.totalRemaining === 0 && data.totalAchieved > 0) {
                rowClass = "bg-green-500 text-white";
            }

            popupRows += `<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
                <td class="border p-1 sm:p-2">${item}</td>
                <td class="border p-1 sm:p-2">${data.customerCount}</td>
                <td class="border p-1 sm:p-2">${data.achievedCustomerCount}</td>
                <td class="border p-1 sm:p-2">${data.totalTarget}</td>
                <td class="border p-1 sm:p-2">${data.totalAchieved}</td>
                <td class="border p-1 sm:p-2">${data.totalRemaining}</td>
                <td class="border p-1 sm:p-2 font-bold">${percentage}%</td>
                <td class="border p-1 sm:p-2 font-bold">${data.totalValue.toLocaleString()}</td>
            </tr>`;

            totalItems++;
            totalAchievedCustomers += data.achievedCustomerCount;
            totalTarget += data.totalTarget;
            totalAchieved += data.totalAchieved;
            totalRemaining += data.totalRemaining;
            totalValue += data.totalValue;
            customerIndex++;
        });

    } else {
        // --- Existing customer-based popup ---
        popupThead = `
            <thead class="bg-gray-100 sticky top-0 z-40">
                <tr>
                    <th class="border p-2">City</th>
                    <th class="border p-2">Customer Code</th>
                    <th class="border p-2">Name</th>
                    <th class="border p-2">Item</th>
                    <th class="border p-2">Target</th>
                    <th class="border p-2">Achieved</th>
                    <th class="border p-2">Remaining</th>
                    <th class="border p-2">%</th>
                     <th class="border p-2">Value</th>
                </tr>
            </thead>
        `;

        Object.entries(customerTargets).forEach(([customerCode, customer]) => {
            if (selectedStatus === "top10" && !top10Customers.includes(customerCode)) {
                return;
            }

            const customerShade = customerShades[customerIndex % customerShades.length];
            customerIndex++;

            let customerHasRow = false;

            Object.entries(customer.items).forEach(([item, target]) => {
                if (!itemFilterAllows(item, selectedPopupItems)) return;

                const achieved = invoices
                    .filter(inv =>
                        inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                        inv.item?.toUpperCase() === item.toUpperCase() &&
                        !isNaN(Number(inv.quantity))
                    )
                    .reduce((sum, inv) => sum + Number(inv.quantity), 0);

                const remaining = target - achieved;
                let statusType = "normal";
                let rowClass = customerShade;

                if (Number(target) === 0) {
                    rowClass = "bg-orange-100";
                    statusType = "zeroTarget";
                } else if (remaining < 0) {
                    rowClass = "bg-red-500 text-white";
                    statusType = "red";
                } else if (remaining <= 0) {
                    rowClass = "bg-green-500 text-white";
                    statusType = "green";
                }

                // 🚫 Non-Productive Filter → show only customers where achieved = 0
if (selectedStatus === "nonProductive") {
    if (achieved > 0) return;  // if any achievement → skip row
}
else {
    // Normal Filters (all, top10, red, green)
    if (
        selectedStatus !== "all" &&
        selectedStatus !== "top10" &&
        selectedStatus !== "nonProductive" &&
        selectedStatus !== statusType
    ) return;
}


               const value = invoices
    .filter(inv =>
        inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
        inv.item?.toUpperCase() === item.toUpperCase()
    )
    .reduce((sum, inv) => sum + (Number(inv.quantity || 0) * Number(inv.rate || 0)), 0);

totalValue += value;
const cityKey = (customer.city || "Unknown City").toString().trim() || "Unknown City";
if (!citySummary[cityKey]) citySummary[cityKey] = { customers: new Set(), items: 0, target: 0, achieved: 0, remaining: 0, value: 0 };
citySummary[cityKey].customers.add(customerCode);
citySummary[cityKey].items += 1;
citySummary[cityKey].target += Number(target) || 0;
citySummary[cityKey].achieved += Number(achieved) || 0;
citySummary[cityKey].remaining += Number(remaining) || 0;
citySummary[cityKey].value += Number(value) || 0;

popupRows += `<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
    <td class="border p-1 sm:p-2">${customer.city || ''}</td>
    <td class="border p-1 sm:p-2">${customerCode}</td>
    <td class="border p-1 sm:p-2">${customer.name || ''}</td>
    <td class="border p-1 sm:p-2">${item}</td>
    <td class="border p-1 sm:p-2">${target}</td>
    <td class="border p-1 sm:p-2">${achieved}</td>
    <td class="border p-1 sm:p-2">${remaining}</td>
    <td class="border p-1 sm:p-2 font-bold">${Number(target) > 0 ? (remaining <= 0 ? "100%" : ((achieved/target*100).toFixed(1)+"%")) : "0%"}</td>
    <td class="border p-1 sm:p-2 font-bold">${value.toLocaleString()}</td>
</tr>`;

                customerHasRow = true;
                totalItems++;
                totalTarget += target;
                totalAchieved += achieved;
                totalRemaining += remaining;
               
             
            });

            if (customerHasRow) totalCustomers++;
        });
    }

    if (!popupRows) return;

    // --- Summary Footer Row ---
 let summaryRow = "";

// 🔵 ITEM SUMMARY POPUP
if (selectedStatus === "cityWise" || selectedStatus === "cityWiseRed" || selectedStatus === "cityWiseGreen") {
    summaryRow = "";
}
else if (
    selectedStatus === "itemSummary" ||
    selectedStatus === "nonProductiveItemSummary"
) {
    summaryRow = `
    <tr class="bg-indigo-100 font-bold text-xs sm:text-sm">
        <td class="border p-2 text-center">TOTAL</td>
        <td class="border p-2 text-center">${totalItems}</td>
        <td class="border p-2 text-center">${totalAchievedCustomers}</td>
        <td class="border p-2">${totalTarget}</td>
        <td class="border p-2">${totalAchieved}</td>
        <td class="border p-2">${totalRemaining}</td>
        <td class="border p-2">${calculateSmartPerformance()}%</td>
        <td class="border p-2">${totalValue.toLocaleString()}</td>
    </tr>`;
}

// 🟢 CUSTOMER-BASED POPUP
else {
    summaryRow = `
    <tr class="bg-indigo-100 font-bold text-xs sm:text-sm">
        <td colspan="4" class="border p-2 text-center">
            TOTAL (${totalCustomers} Customers / ${totalItems} Items)
        </td>

        <td class="border p-2">${totalTarget}</td>
        <td class="border p-2">${totalAchieved}</td>
        <td class="border p-2">${totalRemaining}</td>
        <td class="border p-2">${calculateSmartPerformance()}%</td>
        <td class="border p-2">${totalValue.toLocaleString()}</td>
    </tr>`;
}


    let popup = document.getElementById("invoicePopup");

    // Function to attach copy functionality (har baar call karenge)
    function attachCopyFunctionality() {
        const copyBtn = document.getElementById("copyTableBtn");
        if (!copyBtn) return;

        // Remove previous listener if any (prevent duplicate)
        copyBtn.replaceWith(copyBtn.cloneNode(true));
        const newBtn = document.getElementById("copyTableBtn");

        newBtn.addEventListener("click", function() {
            const table = document.getElementById("popupTable");
            if (!table) {
                alert("Table not found!");
                return;
            }

            let text = "";

            // Header
            const headers = table.querySelectorAll("thead th");
            if (headers.length > 0) {
                text += Array.from(headers)
                    .map(th => th.innerText.trim().replace(/\s+/g, ' '))
                    .join("\t") + "\n";
            }

            // Body rows (including summary)
            const rows = table.querySelectorAll("tbody tr");
            rows.forEach(row => {
                const cells = row.querySelectorAll("td");
                text += Array.from(cells)
                    .map(td => td.innerText.trim().replace(/\s+/g, ' '))
                    .join("\t") + "\n";
            });

            navigator.clipboard.writeText(text).then(() => {
                const originalText = newBtn.innerHTML;
                newBtn.innerHTML = "✅ Copied to Clipboard!";
                newBtn.disabled = true;
                newBtn.classList.remove("bg-blue-600", "hover:bg-blue-700");
                newBtn.classList.add("bg-green-600");
                
                setTimeout(() => {
                    newBtn.innerHTML = originalText;
                    newBtn.disabled = false;
                    newBtn.classList.remove("bg-green-600");
                    newBtn.classList.add("bg-blue-600", "hover:bg-blue-700");
                }, 2000);
            }).catch(err => {
                console.error("Copy failed:", err);
                alert("Copy failed! Browser may not support or page is not secure (HTTPS needed).");
            });
        });
    }

    if (!popup) {
        // First time creation
        popup = document.createElement("div");
        popup.id = "invoicePopup";
        popup.className = "fixed inset-0 bg-black bg-opacity-50 flex justify-center items-start pt-2 z-50 hidden";
        popup.innerHTML = `
            <div class="bg-white rounded shadow-lg w-full h-full sm:w-[95%] sm:max-w-6xl sm:h-[80vh] flex flex-col overflow-hidden">
                <div class="overflow-auto p-2 flex-1">
                    <table id="popupTable" class="w-full border-collapse border text-xs sm:text-sm">
                        ${popupThead}
                        <tbody id="popupInvoiceBody">${popupRows}${summaryRow}</tbody>
                    </table>
                </div>
                <div class="p-3 border-t bg-gray-100 flex flex-col sm:flex-row gap-3 justify-between items-center">
                    <button id="copyTableBtn" class="bg-blue-600 text-white px-5 py-2 rounded hover:bg-blue-700 flex items-center gap-2 font-medium">
                        📋 Copy Table to Clipboard
                    </button>
                    <button id="closePopup" class="bg-red-600 text-white px-5 py-2 rounded hover:bg-red-700">
                        ✖ Close
                    </button>
                </div>
            </div>
        `;
        document.body.appendChild(popup);

        // Close button
        document.getElementById("closePopup").addEventListener("click", () => {
            popup.classList.add("hidden");
        });

        // Attach copy functionality
        attachCopyFunctionality();
    } else {
        // Update existing popup
        const table = popup.querySelector("#popupTable");
        if (table) {
            table.querySelector("thead").outerHTML = popupThead;
            document.getElementById("popupInvoiceBody").innerHTML = popupRows + summaryRow;
        }
        
        // Re-attach copy functionality (important!)
        attachCopyFunctionality();
    }

    popup.classList.remove("hidden");
}






function renderAllocationTables(customerCode = null) {
    const tablesContainer = document.getElementById('allocationTables');
    if (!tablesContainer) {
        console.error('allocationTables element not found');
        return;
    }

    if (customerCode === lastRenderedCustomerCode) {
        console.log('Skipping render: same customerCode already rendered:', customerCode);
        return;
    }

    tablesContainer.innerHTML = '';
    lastRenderedCustomerCode = customerCode;
    console.log('Rendering allocation table for customerCode:', customerCode);

    if (!customerCode) {
        tablesContainer.innerHTML = '<p class="text-center text-gray-500">Please search for a customer to view dashboard.</p>';
        return;
    }

    const customer = customerTargets[customerCode];
    if (!customer) {
        tablesContainer.innerHTML = '<p class="text-center text-gray-500">Customer not found.</p>';
        console.error('Customer not found for allocation:', customerCode);
        return;
    }

    const rankedCustomers = getCustomerRankings();
    const rankInfo = rankedCustomers.find(c => c.code === customerCode);
    const customerLevel = rankInfo ? rankInfo.displayLevel : "";
    const levelColor = rankInfo ? rankInfo.levelColor : "#888";

    // --- Table Calculation ---
    let rowsHtml = '';
    let totalTarget = 0, totalAchieved = 0, totalRemaining = 0, totalAchievedValue = 0;
    let totalItems = 0, nonProductive = 0, completed = 0, progress = 0;
    const zeroAchieveItems = [];

    const sortedItems = Object.keys(customer.items).sort((a, b) => a.localeCompare(b));

    sortedItems.forEach(item => {
        const target = Number(customer.items[item]);
        const matchingInvoices = invoices.filter(inv =>
            inv &&
            inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
            inv.item?.toUpperCase() === item.toUpperCase() &&
            !isNaN(Number(inv.quantity)) &&
            !isNaN(Number(inv.rate))
        );

        const achieved = matchingInvoices.reduce((sum, inv) => sum + Number(inv.quantity), 0);
        const achievedValue = matchingInvoices.reduce((sum, inv) => sum + (Number(inv.quantity) * Number(inv.rate)), 0);
        const cappedAchieved = Math.min(achieved, target);
        const remaining = target - achieved;

        totalTarget += target;
        totalAchieved += cappedAchieved;
        totalRemaining += Math.max(remaining, 0);
        totalAchievedValue += achievedValue;
        totalItems++;

        let rowStyle = "";
        if (remaining < 0) {
            rowStyle = "background-color: #dc2626; color: white;";
        } else if (achieved >= target) {
            rowStyle = "background-color: #16a34a; color: white;";
        } else if (achieved > 0) {
            const percent = Math.min((achieved / target) * 100, 100);
            rowStyle = `
                background: linear-gradient(
                    to right,
                    #16a34a ${percent}%,
                    #60a5fa ${percent}%
                );
                color: white;
                transition: background 0.6s ease;
            `;
        }

        if (achieved === 0) {
            nonProductive++;
            zeroAchieveItems.push(item);
        } else if (achieved >= target) {
            completed++;
        } else {
            progress++;
        }

        rowsHtml += `<tr style="${rowStyle}">
            <td class="border p-2">${item?.trim() || ''}</td>
            <td class="border p-2">${target.toLocaleString()}</td>
            <td class="border p-2">${achieved.toLocaleString()}</td>
            <td class="border p-2">${remaining.toLocaleString()}</td>
            <td class="border p-2 font-bold">${achievedValue.toLocaleString()}</td>
        </tr>`;
    });

    if (!rowsHtml) {
        rowsHtml = '<tr><td colspan="5" class="p-2 text-center">No items for this customer.</td></tr>';
    }

    const overallPercent = totalTarget > 0 ? ((totalAchieved / totalTarget) * 100).toFixed(1) : 0;

    // --- Final HTML Output ---
    tablesContainer.innerHTML = `
        <!-- Header -->
        <div class="mb-6 text-center p-6 rounded-2xl shadow-lg bg-gradient-to-r from-purple-700 via-purple-800 to-gray-900">
            <div class="flex justify-between items-center">
                <p class="text-sm font-bold px-3 py-1 rounded-full text-black" style="background-color:${levelColor}">
                    ${customerLevel}
                </p>
                <h2 class="text-lg font-extrabold text-white drop-shadow-lg flex-grow text-center">📊 Customer Dashboard</h2>
                <span></span>
            </div>
            <p class="text-3xl font-extrabold text-yellow-400 drop-shadow-lg mt-2">${customer.name || 'Unknown Name'}</p>
            <p class="text-gray-300 text-sm mt-1">${customer.city || 'Unknown City'} • ${customerCode}</p>
        </div>

        <!-- KPI Cards -->
        <div class="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
            <div class="p-5 rounded-2xl shadow-lg text-center bg-gradient-to-br from-blue-50 to-blue-100 hover:from-blue-100 hover:to-blue-200 transform hover:-translate-y-1 transition duration-300">
                <h3 class="text-lg font-bold text-blue-700">📦 Total Items</h3>
                <p class="text-3xl font-extrabold text-blue-900 mt-2">${totalItems}</p>
            </div>
            <div class="p-5 rounded-2xl shadow-lg text-center bg-gradient-to-br from-red-50 to-red-100 hover:from-red-100 hover:to-red-200 transform hover:-translate-y-1 transition duration-300">
                <h3 class="text-lg font-bold text-red-700">🚫 Non-Productive</h3>
                <p class="text-3xl font-extrabold text-red-900 mt-2">${nonProductive}</p>
            </div>
            <div class="p-5 rounded-2xl shadow-lg text-center bg-gradient-to-br from-yellow-50 to-yellow-100 hover:from-yellow-100 hover:to-yellow-200 transform hover:-translate-y-1 transition duration-300">
                <h3 class="text-lg font-bold text-yellow-700">⏳ In Progress</h3>
                <p class="text-3xl font-extrabold text-yellow-900 mt-2">${progress}</p>
            </div>
            <div class="p-5 rounded-2xl shadow-lg text-center bg-gradient-to-br from-green-50 to-green-100 hover:from-green-100 hover:to-green-200 transform hover:-translate-y-1 transition duration-300">
                <h3 class="text-lg font-bold text-green-700">✅ Completed</h3>
                <p class="text-3xl font-extrabold text-green-900 mt-2">${completed}</p>
            </div>
            <div class="p-5 rounded-2xl shadow-lg text-center bg-gradient-to-br from-purple-50 to-purple-100 hover:from-purple-100 hover:to-purple-200 transform hover:-translate-y-1 transition duration-300">
                <h3 class="text-lg font-bold text-purple-700">💰 Total Value</h3>
                <p class="text-3xl font-extrabold text-purple-900 mt-2">${totalAchievedValue.toLocaleString()}</p>
            </div>
        </div>

        <!-- Progress Bar -->
        <div class="mb-6">
            <h3 class="font-semibold mb-2">📈 Overall Achievement</h3>
            <div class="w-full bg-gray-200 rounded-full h-6 overflow-hidden">
                <div class="h-6 text-xs flex items-center justify-center font-bold text-white rounded-full"
                     style="width:${overallPercent}%; background: linear-gradient(to right, #60a5fa, #16a34a); transition: width 0.6s ease;">
                    ${overallPercent}%
                </div>
            </div>
        </div>

        <!-- Breaking News -->
        <div id="breakingNews" class="relative overflow-hidden h-10 font-semibold text-sm rounded-lg shadow-lg mb-6
                    bg-gradient-to-r from-red-500 via-yellow-400 to-red-500 border border-red-600">
            ${zeroAchieveItems.length > 0
                ? `<marquee behavior="scroll" direction="left" scrollamount="6">
                    ${zeroAchieveItems.map(it => `
                        <span class="text-white mx-4 bg-red-600 px-2 py-1 rounded-full shadow-md">
                            🚨 ${it}
                        </span>`).join("")}
                  </marquee>`
                : '<span class="text-gray-600 flex items-center justify-center h-full">No alerts at this time</span>'}
        </div>

        <!-- Table -->
        <div class="resizable-box" id="customerTableBox">
            <div class="customer-table scrollable-table">
                <table>
                    <thead>
                        <tr>
                            <th class="border p-2 bg-secondary">Item</th>
                            <th class="border p-2 bg-secondary">Target</th>
                            <th class="border p-2 bg-secondary">Achieved</th>
                            <th class="border p-2 bg-secondary">Remaining</th>
                            <th class="border p-2 bg-secondary">Achieved Value</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rowsHtml}
                    </tbody>
                    <tfoot>
                        <tr class="font-extrabold bg-gradient-to-r from-indigo-700 via-blue-700 to-indigo-800 text-white text-lg shadow-inner">
                            <td class="border-2 border-indigo-900 p-3 text-center">Total</td>
                            <td class="border-2 border-indigo-900 p-3 text-right">${totalTarget.toLocaleString()}</td>
                            <td class="border-2 border-indigo-900 p-3 text-right">${totalAchieved.toLocaleString()}</td>
                            <td class="border-2 border-indigo-900 p-3 text-right">${totalRemaining.toLocaleString()}</td>
                            <td class="border-2 border-indigo-900 p-3 text-right">${totalAchievedValue.toLocaleString()}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    `;
}








function renderDoneTargetTables() {
    const container = document.getElementById('doneTargetTables');
    if (!container) {
        console.error('doneTargetTables element not found');
        return;
    }

    let updatedDoneTargets = [];
    let extraAllocations = [];

    // ✅ Customer-wise check (all items completed)
    Object.entries(customerTargets).forEach(([customerCode, data]) => {
        let allCompleted = true;
        let customerAchievedItems = [];

        data.items && Object.entries(data.items).forEach(([item, target]) => {
            const achieved = invoices
                .filter(inv => inv && inv.customerCode?.toUpperCase() === customerCode.toUpperCase() && inv.item?.toUpperCase() === item.toUpperCase() && !isNaN(Number(inv.quantity)))
                .reduce((sum, inv) => sum + Number(inv.quantity), 0);

            if (achieved < target) {
                allCompleted = false;
            }

            customerAchievedItems.push({
                item,
                target,
                achieved,
                remaining: target - achieved
            });

            if (achieved > target) {
                extraAllocations.push({
                    customer: data.name,
                    item,
                    achieved,
                    target
                });
            }
        });

        if (allCompleted && customerAchievedItems.length > 0) {
            updatedDoneTargets.push({
                customerCode,
                customer: data.name,
                city: data.city,
                items: customerAchievedItems
            });
        }
    });

    // ✅ Dashboard counters
    const totalDone = updatedDoneTargets.length;
    const totalExtra = extraAllocations.length;
    const totalPending = Object.values(customerTargets).reduce((count, data) => {
        let pending = 0;
        Object.entries(data.items || {}).forEach(([item, target]) => {
            const achieved = invoices
                .filter(inv => inv.customerCode?.toUpperCase() === data.code?.toUpperCase() && inv.item?.toUpperCase() === item.toUpperCase())
                .reduce((sum, inv) => sum + Number(inv.quantity || 0), 0);
            if (achieved < target) pending++;
        });
        return count + pending;
    }, 0);

    let tablesHtml = `
        <!-- ✅ Dashboard -->
        <div class="grid grid-cols-3 gap-4 mb-4 text-center">
            <div class="p-4 bg-green-500 text-white font-bold rounded-lg shadow">Done Customers<br><span class="text-2xl">${totalDone}</span></div>
            <div class="p-4 bg-blue-500 text-white font-bold rounded-lg shadow">Extra Allocations<br><span class="text-2xl">${totalExtra}</span></div>
            <div class="p-4 bg-red-500 text-white font-bold rounded-lg shadow">Total Allocation<br><span class="text-2xl">${totalPending}</span></div>
        </div>
    `;

    // ✅ Breaking News directly under Dashboard
    if (extraAllocations.length > 0) {
        const newsItems = extraAllocations.map(ea =>
            `${ea.customer} → ${ea.item}: Achieved ${ea.achieved} (Target ${ea.target})`
        ).join(" ⚡ ");

        tablesHtml += `
            <div class="mb-6 bg-black text-yellow-300 p-2 rounded shadow">
                <marquee behavior="scroll" direction="left" scrollamount="6" class="font-bold text-sm">
                    🔥 Extra Allocations: ${newsItems}
                </marquee>
            </div>
        `;
    }

    updatedDoneTargets.forEach(data => {
        let rowsHtml = '';
        data.items.forEach((dt, index) => {
            let rowClass = "";
            if (dt.achieved > dt.target) {
                rowClass = "bg-purple-200"; // extra
            } else if (dt.achieved === dt.target) {
                rowClass = "bg-green-200"; // completed
            } else {
                rowClass = "bg-sky-200"; // partial
            }

            const extraStyle = index % 2 === 0 ? "bg-gray-50" : "bg-white";

            rowsHtml += `<tr class="${rowClass} ${extraStyle} hover:bg-yellow-100 transition">
                <td class="border p-2 text-sm font-medium">${dt.item}</td>
                <td class="border p-2 text-center">${dt.target}</td>
                <td class="border p-2 text-center">${dt.achieved}</td>
                <td class="border p-2 text-center">${dt.remaining}</td>
            </tr>`;
        });

        tablesHtml += `
            <div class="customer-table mb-6 shadow-lg rounded-xl overflow-hidden border border-gray-300">
                <h3 class="text-lg font-bold mb-2 bg-gradient-to-r from-green-600 to-green-800 text-white p-2 rounded-t-xl shadow">
                    ✅ ${data.customer} (${data.customerCode}) - ${data.city}
                </h3>
                <div class="overflow-x-auto scrollable-table">
                    <table class="min-w-full border-collapse">
                        <thead class="bg-gradient-to-r from-gray-800 via-gray-900 to-black text-white text-sm uppercase tracking-wider shadow-md sticky top-0">
                            <tr>
                                <th class="border p-3 text-left">Item</th>
                                <th class="border p-3 text-center">Target</th>
                                <th class="border p-3 text-center">Achieved</th>
                                <th class="border p-3 text-center">Remaining</th>
                            </tr>
                        </thead>
                        <tbody>${rowsHtml}</tbody>
                    </table>
                </div>
            </div>`;
    });

    if (!tablesHtml) {
        tablesHtml = '<p class="text-center text-gray-500">No customers have fully completed their targets yet.</p>';
    }

    container.innerHTML = tablesHtml;
}





function updateCustomerReport(customerCode) {
    const reportContainer = document.getElementById('customerReport');
    const reportTitle = document.getElementById('customerReportTitle');
    const tbody = document.getElementById('customerReportBody');
    const totT = document.getElementById('customerReportTotalTarget');
    const totA = document.getElementById('customerReportTotalAchieved');
    const totR = document.getElementById('customerReportTotalRemaining');

    if (!tbody || !totT || !totA || !totR || !reportContainer || !reportTitle) {
        console.error('Customer report DOM elements missing');
        return;
    }

    const customer = customerTargets[customerCode]?.name || '';
    const itemMap = customerTargets[customerCode]?.items || {};
    let rowsHtml = '';
    let totalTarget = 0, totalAchieved = 0, totalRemaining = 0;

    reportTitle.textContent = `${customer} (${customerCode}) Item-wise Summary`;
    console.log(`Generating customer report: customerCode=${customerCode}, items:`, itemMap);

    Object.entries(itemMap).forEach(([item, target]) => {
        const achieved = invoices
            .filter(inv => inv && inv.customerCode?.toUpperCase() === customerCode.toUpperCase() && inv.item?.toUpperCase() === item.toUpperCase() && !isNaN(Number(inv.quantity)))
            .reduce((sum, inv) => sum + Number(inv.quantity), 0);
        const remaining = Number(target) - achieved;
        totalTarget += Number(target);
        totalAchieved += achieved;
        totalRemaining += remaining;
        rowsHtml += `<tr>
            <td class="border p-2">${item}</td>
            <td class="border p-2">${target}</td>
            <td class="border p-2">${achieved}</td>
            <td class="border p-2">${remaining}</td>
        </tr>`;
        console.log(`Customer report: item=${item}, target=${target}, achieved=${achieved}, remaining=${remaining}`);
    });

    if (!rowsHtml) {
        rowsHtml = '<tr><td colspan="4" class="p-2 text-center">No items for this customer.</td></tr>';
        totalTarget = 0;
        totalAchieved = 0;
        totalRemaining = 0;
    }

    tbody.innerHTML = rowsHtml;
    totT.textContent = String(totalTarget);
    totA.textContent = String(totalAchieved);
    totR.textContent = String(totalRemaining);
    reportContainer.classList.remove('hidden');
}

function exportData(format = "csv") {
    console.log(`Exporting to ${format.toUpperCase()}...`);
    const csvData = [];

    // ----------------- Invoices Section (A–J) -----------------
    const invoiceHeaders = ['City', 'CustomerCode', 'Customer', 'Item', 'Target', 'Achieve', 'User1', 'User2', 'Qty', 'Bonus'];
    csvData.push(invoiceHeaders);

    Object.entries(customerTargets).forEach(([customerCode, customer]) => {
        Object.entries(customer.items).forEach(([item, target]) => {
            const achieved = invoices
                .filter(inv =>
                    inv &&
                    inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                    inv.item?.toUpperCase() === item.toUpperCase() &&
                    !isNaN(Number(inv.quantity))
                )
                .reduce((sum, inv) => sum + Number(inv.quantity), 0);

            let user1 = '';
            let user2 = '';
            invoices.forEach(inv => {
                if (inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                    inv.item?.toUpperCase() === item.toUpperCase()) {
                    if (!user1 && inv.user) user1 = inv.user;
                }
            });

            const deals = bonusDeals[item] || [];
            const dealQty = deals.length > 0 ? Math.min(...deals.map(d => d.qty)) : 0;
            const dealBonus = deals.length > 0 ? Math.max(...deals.map(d => d.bonus)) : 0;

            const row = [customer.city, customerCode, customer.name, item, target, achieved, user1, user2, dealQty, dealBonus];
            csvData.push(row);
        });
    });

    // ----------------- Bonus Deals Section -----------------
    if (bonusDeals && Object.keys(bonusDeals).length > 0) {
        csvData.push([]); // blank separator
        csvData.push(['Item', 'DealQty', 'DealBonus']); // headers

        Object.entries(bonusDeals).forEach(([item, deals]) => {
            deals.forEach(d => {
                const row = [item, d.qty, d.bonus];
                csvData.push(row);
            });
        });
    }

    // ----------------- My Sale Data Section (K–N) -----------------
    const mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
    if (mySaleData.length > 0) {
        csvData.push([]); // blank separator
        csvData.push(['SummaryNumber', 'CompanyName', 'Value', 'Date']); // headers

        mySaleData.forEach(sale => {
            const row = [
                sale.summary || '',
                sale.company || '',
                Number(sale.value) || 0,
                sale.date || ''
            ];
            csvData.push(row);
        });
    }

    // ----------------- Export Logic -----------------
    if (format === "csv") {
        // CSV Build
        const csvContent = csvData
            .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
            .join('\n');

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `export_${new Date().toISOString().split('T')[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

    } else if (format === "xlsx") {
        // Excel Build with SheetJS
        const ws = XLSX.utils.aoa_to_sheet(csvData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Export");

        XLSX.writeFile(wb, `export_${new Date().toISOString().split('T')[0]}.xlsx`);
    }

    console.log(`${format.toUpperCase()} exported successfully`);
}





function showResetExcelModal() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const resetExcelModal = document.getElementById('resetExcelModal');
    if (resetExcelModal) {
        resetExcelModal.classList.remove('hidden');
        toggleSidebar();
    }
}

function closeResetExcelModal() {
    const resetExcelModal = document.getElementById('resetExcelModal');
    const resetExcelError = document.getElementById('resetExcelError');
    if (resetExcelModal && resetExcelError) {
        resetExcelModal.classList.add('hidden');
        resetExcelError.classList.add('hidden');
    }
}

function resetExcel() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const password = document.getElementById('resetExcelPassword')?.value;
    const resetExcelError = document.getElementById('resetExcelError');
    if (!password || !resetExcelError) return;
    if (password === '123') {
        excelData = [];
        customers = [];
        customerCodes = [];
        items = [];
        customerTargets = {};
        buildCustomerTargets();
        localStorage.setItem('excelData', JSON.stringify(excelData));
        localStorage.setItem('customers', JSON.stringify(customers));
        localStorage.setItem('customerCodes', JSON.stringify(customerCodes));
        localStorage.setItem('items', JSON.stringify(items));
        document.getElementById('excelFile').value = '';
        document.getElementById('customer').value = '';
        document.getElementById('item').value = '';
        document.getElementById('city').value = '';
        document.getElementById('target').value = '';
        document.getElementById('quantity').value = '';
        document.getElementById('remaining').value = '';
        closeResetExcelModal();
        alert('Excel/CSV data reset successfully!');
        renderInvoiceTable();
    } else {
        resetExcelError.classList.remove('hidden');
    }
}

function showResetInvoicesModal() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const resetInvoicesModal = document.getElementById('resetInvoicesModal');
    if (resetInvoicesModal) {
        resetInvoicesModal.classList.remove('hidden');
        toggleSidebar();
    }
}

function closeResetInvoicesModal() {
    const resetInvoicesModal = document.getElementById('resetInvoicesModal');
    const resetInvoicesError = document.getElementById('resetInvoicesError');
    if (resetInvoicesModal && resetInvoicesError) {
        resetInvoicesModal.classList.add('hidden');
        resetInvoicesError.classList.add('hidden');
    }
}

function resetInvoices() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const password = document.getElementById('resetInvoicesPassword')?.value;
    const resetInvoicesError = document.getElementById('resetInvoicesError');
    if (!password || !resetInvoicesError) return;
    if (password === '123') {
        invoices = [];
        localStorage.setItem('invoices', JSON.stringify(invoices));
        renderAllocationTables();
        renderInvoiceTable();
        closeResetInvoicesModal();
        alert('All invoices reset successfully!');
    } else {
        resetInvoicesError.classList.remove('hidden');
    }
}

function showResetDoneModal() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const resetDoneModal = document.getElementById('resetDoneModal');
    if (resetDoneModal) {
        resetDoneModal.classList.remove('hidden');
        toggleSidebar();
    }
}

function closeResetDoneModal() {
    const resetDoneModal = document.getElementById('resetDoneModal');
    const resetDoneError = document.getElementById('resetDoneError');
    if (resetDoneModal && resetDoneError) {
        resetDoneModal.classList.add('hidden');
        resetDoneError.classList.add('hidden');
    }
}

function resetDoneTargets() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const password = document.getElementById('resetDonePassword')?.value;
    const resetDoneError = document.getElementById('resetDoneError');
    if (!password || !resetDoneError) return;
    if (password === '123') {
        doneTargets = [];
        localStorage.setItem('doneTargets', JSON.stringify(doneTargets));
        renderDoneTargetTables();
        closeResetDoneModal();
        alert('All done targets reset successfully!');
    } else {
        resetDoneError.classList.remove('hidden');
    }
}

function showResetAppModal() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const resetAppModal = document.getElementById('resetAppModal');
    if (resetAppModal) {
        resetAppModal.classList.remove('hidden');
        toggleSidebar();
    }
}

function closeResetAppModal() {
    const resetAppModal = document.getElementById('resetAppModal');
    const resetAppError = document.getElementById('resetAppError');
    if (resetAppModal && resetAppError) {
        resetAppModal.classList.add('hidden');
        resetAppError.classList.add('hidden');
    }
}

function resetApp() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const password = document.getElementById('resetAppPassword')?.value;
    const resetAppError = document.getElementById('resetAppError');
    if (!password || !resetAppError) return;

    if (password === '123') {
        // reset arrays / objects
        excelData = [];
        invoices = [];
        doneTargets = [];
        customers = [];
        customerCodes = [];
        items = [];
        customerTargets = {};
        bonusDeals = {};
        mySaleData = [];   // ✅ My Sale reset

        // update localStorage
        localStorage.setItem('excelData', JSON.stringify(excelData));
        localStorage.setItem('invoices', JSON.stringify(invoices));
        localStorage.setItem('doneTargets', JSON.stringify(doneTargets));
        localStorage.setItem('customers', JSON.stringify(customers));
        localStorage.setItem('customerCodes', JSON.stringify(customerCodes));
        localStorage.setItem('items', JSON.stringify(items));
        localStorage.setItem('customerTargets', JSON.stringify(customerTargets));
        localStorage.setItem('bonusDeals', JSON.stringify(bonusDeals));
        localStorage.setItem('mySaleData', JSON.stringify(mySaleData)); // ✅ save empty sale data

        // clear inputs
        document.getElementById('excelFile').value = '';
        document.getElementById('customer').value = '';
        document.getElementById('item').value = '';
        document.getElementById('city').value = '';
        document.getElementById('target').value = '';
        document.getElementById('quantity').value = '';
        document.getElementById('remaining').value = '';
        document.getElementById('customerSuggestions').classList.add('hidden');
        document.getElementById('itemSuggestions').classList.add('hidden');
        document.getElementById('customerSearch').value = '';
        document.getElementById('citySelect').value = '';
        document.getElementById('customerReport').classList.add('hidden');
        document.getElementById('customerSearchSuggestions').classList.add('hidden');
        document.getElementById('bonusItemSelect').value = '';
        document.getElementById('bonusQty').value = '';
        document.getElementById('bonusValue').value = '';

        // re-render UI
        renderAllocationTables();
        renderDoneTargetTables();
        renderBonusDeals();
        if (typeof renderMySaleTable === "function") {
            renderMySaleTable(); // ✅ refresh My Sale page
        }

        closeResetAppModal();
        showMainPage();
        alert('App reset successfully!');
    } else {
        resetAppError.classList.remove('hidden');
    }
}


function filterCustomersByCity() {
    const citySelect = document.getElementById('citySelect');
    const customerSearch = document.getElementById('customerSearch');
    const suggestionsDiv = document.getElementById('customerSearchSuggestions');
    if (!citySelect || !customerSearch || !suggestionsDiv) return;
    customerSearch.value = '';
    suggestionsDiv.classList.add('hidden');
    suggestionsDiv.innerHTML = '';
    document.getElementById('customerReport').classList.add('hidden');
    document.getElementById('customerReportTitle').textContent = 'Customer Item-wise Summary';
    document.getElementById('customerReportBody').innerHTML = '';
    document.getElementById('customerReportTotalTarget').textContent = '0';
    document.getElementById('customerReportTotalAchieved').textContent = '0';
    document.getElementById('customerReportTotalRemaining').textContent = '0';
    renderAllocationTables();
}

function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => {
            console.log('Debounced function called with args:', args);
            func.apply(this, args);
        }, wait);
    };
}

function handleCustomerSearch() {
    const customerSearch = document.getElementById('customerSearch');
    const citySelect = document.getElementById('citySelect');
    const suggestionsDiv = document.getElementById('customerSearchSuggestions');
    if (!customerSearch || !citySelect || !suggestionsDiv) {
        console.error('Customer search DOM elements missing');
        return;
    }

    const q = customerSearch.value.trim().toLowerCase();
    suggestionsDiv.innerHTML = '';
    suggestionsDiv.classList.add('hidden');
    renderAllocationTables();
    console.log('Handling customer search for query:', q);

    if (!q) {
        console.log('Empty search query, cleared tables');
        return;
    }

    let filteredCustomers = customers;
    if (citySelect.value.trim()) {
        filteredCustomers = customers.filter(c => c.city.toLowerCase() === citySelect.value.trim().toLowerCase());
    }

    filteredCustomers = filteredCustomers.filter(c => 
        c.code.toLowerCase() === q || 
        c.name.toLowerCase() === q ||
        c.code.toLowerCase().includes(q) || 
        c.name.toLowerCase().includes(q)
    );

    if (filteredCustomers.length > 0) {
        suggestionsDiv.classList.remove('hidden');
        filteredCustomers.forEach(customer => {
            const suggestion = document.createElement('div');
            suggestion.className = 'p-2 hover:bg-teal-500 cursor-pointer';
            suggestion.textContent = `${customer.name} (${customer.code}) - ${customer.city}`;
            suggestion.addEventListener('click', () => {
                customerSearch.value = `${customer.code} - ${customer.name}`;
                customerSearch.select();
                suggestionsDiv.classList.add('hidden');
                renderAllocationTables(customer.code);
                console.log('Customer selected for allocation:', customer.code);
            });
            suggestionsDiv.appendChild(suggestion);
        });
    }

    let customerCode = null;
    const exactMatch = filteredCustomers.find(c => c.code.toLowerCase() === q || c.name.toLowerCase() === q);
    if (exactMatch) {
        customerCode = exactMatch.code;
    } else if (filteredCustomers.length === 1) {
        customerCode = filteredCustomers[0].code;
    }

    if (customerCode) {
        renderAllocationTables(customerCode);
        console.log('Rendering table for customerCode:', customerCode);
    } else {
        console.log('No customer found for search query:', q);
    }
}

function renderBonusDeals() {
    const tbody = document.querySelector('#bonusTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const uniqueItems = [...new Set(Object.keys(bonusDeals))].sort((a, b) => a.localeCompare(b));

    let totalItems = 0;
    let totalMinQty = 0;
    let totalMaxBonus = 0;

    uniqueItems.forEach(item => {
        let deals = bonusDeals[item];
        let minQty = Math.min(...deals.map(d => d.qty));
        let maxBonus = Math.max(...deals.map(d => d.bonus));

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="p-2">${item}</td>
            <td class="p-2 text-center">${minQty}</td>
            <td class="p-2 text-center">${maxBonus}</td>
        `;
        tbody.appendChild(tr);

        totalItems++;
        totalMinQty += minQty;
        totalMaxBonus += maxBonus;
    });

    const totalRow = document.createElement('tr');
    totalRow.innerHTML = `
        <td class="p-2 font-bold">Total (${totalItems} items)</td>
        <td class="p-2 text-center font-bold">${totalMinQty}</td>
        <td class="p-2 text-center font-bold">${totalMaxBonus}</td>
    `;
    tbody.appendChild(totalRow);
}

function populateBonusItems() {
    const sel = document.getElementById('bonusItemSelect');
    if (!sel) return;
    sel.innerHTML = '<option value="">Select Item</option>';

    const uniqueItems = [...new Set(Object.keys(bonusDeals))].sort((a, b) => a.localeCompare(b));

    uniqueItems.forEach(item => {
        let opt = document.createElement('option');
        opt.value = item;
        opt.textContent = item;
        sel.appendChild(opt);
    });
}

function calculateFromQty() {
    const item = document.getElementById('bonusItemSelect').value;
    const qty = parseInt(document.getElementById('bonusQty').value) || 0;
    if (item && bonusDeals[item]) {
        let totalBonus = 0;
        bonusDeals[item].forEach(d => {
            totalBonus += Math.floor(qty / d.qty) * d.bonus;
        });
        document.getElementById('bonusValue').value = totalBonus;
    } else {
        document.getElementById('bonusValue').value = '';
    }
}

function calculateFromBonus() {
    const item = document.getElementById('bonusItemSelect').value;
    const bonus = parseInt(document.getElementById('bonusValue').value) || 0;
    if (item && bonusDeals[item]) {
        let requiredQty = 0;
        bonusDeals[item].forEach(d => {
            let q = Math.ceil(bonus / d.bonus) * d.qty;
            if (q > requiredQty) requiredQty = q;
        });
        document.getElementById('bonusQty').value = requiredQty;
    } else {
        document.getElementById('bonusQty').value = '';
    }
}

function resetBonusPlan() {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        return;
    }
    const password = prompt('Enter password to reset bonus plan:');
    if (password === '123') {
        bonusDeals = {};
        localStorage.setItem('bonusDeals', JSON.stringify(bonusDeals));
        document.getElementById('bonusItemSelect').value = '';
        document.getElementById('bonusQty').value = '';
        document.getElementById('bonusValue').value = '';
        renderBonusDeals();
        populateBonusItems();
        alert('Bonus plan reset successfully!');
    } else {
        alert('Invalid password!');
    }
}

window.onload = () => {
    try {
        let codeSection = document.getElementById('codeSection');
        if (!codeSection) {
            console.warn('codeSection not found, creating dynamically');
            codeSection = document.createElement('div');
            codeSection.id = 'codeSection';
            codeSection.classList.add('hidden');
            codeSection.innerHTML = `
                <div class="text-center">
                    <h2 class="text-lg font-bold mb-4">App Locked</h2>
                    <p class="mb-2">Code: <span id="displayCode"></span></p>
                    <input id="unlockCode" type="text" placeholder="Enter Unlock Code" class="border p-2 w-full mb-2">
                    <p id="codeError" class="hidden text-red-500 mb-2">Invalid Code!</p>
                    <button onclick="unlockApp()" class="bg-primary text-white p-2 rounded">Unlock</button>
                </div>
            `;
            document.body.appendChild(codeSection);
        }

        // Always check lock status first
        checkLockStatus();
        if (isAppLocked) {
            console.log('App is locked, skipping further initialization');
            return;
        }

        const storedExcelData = localStorage.getItem('excelData');
        if (storedExcelData) {
            excelData = JSON.parse(storedExcelData);
            customers = JSON.parse(localStorage.getItem('customers') || '[]');
            customerCodes = JSON.parse(localStorage.getItem('customerCodes') || '[]');
            items = JSON.parse(localStorage.getItem('items') || '[]');
            bonusDeals = JSON.parse(localStorage.getItem('bonusDeals') || '{}');
            buildCustomerTargets();
            renderBonusDeals();
            populateBonusItems();
        } else {
            excelData = [];
            customers = [];
            customerCodes = [];
            items = [];
            customerTargets = {};
            bonusDeals = {};
            localStorage.setItem('excelData', JSON.stringify(excelData));
            localStorage.setItem('customers', JSON.stringify(customers));
            localStorage.setItem('customerCodes', JSON.stringify(customerCodes));
            localStorage.setItem('items', JSON.stringify(items));
            localStorage.setItem('bonusDeals', JSON.stringify(bonusDeals));
        }

        const storedInvoices = localStorage.getItem('invoices');
        if (storedInvoices) {
            invoices = JSON.parse(storedInvoices);
            invoices = invoices.filter(inv => inv && inv.customerCode?.trim() && inv.item?.trim() && !isNaN(Number(inv.quantity)));
            localStorage.setItem('invoices', JSON.stringify(invoices));
            console.log('Loaded and validated invoices:', invoices);
        } else {
            invoices = [];
            localStorage.setItem('invoices', JSON.stringify(invoices));
        }

        const storedDoneTargets = localStorage.getItem('doneTargets');
        if (storedDoneTargets) {
            doneTargets = JSON.parse(storedDoneTargets);
            console.log('Loaded doneTargets:', doneTargets);
        } else {
            doneTargets = [];
            localStorage.setItem('doneTargets', JSON.stringify(doneTargets));
        }

        initHamburger();
        const customerInput = document.getElementById('customer');
        if (customerInput) {
            customerInput.addEventListener('input', autoFillCity);
        }
        const customerSearch = document.getElementById('customerSearch');
        if (customerSearch) {
            customerSearch.addEventListener('input', debounce(handleCustomerSearch, 300));
        }
        const itemInput = document.getElementById('item');
        if (itemInput) {
            itemInput.addEventListener('keydown', (event) => {
                if (event.key === 'Backspace' && itemInput.value.trim() !== '') {
                    itemInput.select();
                    console.log('Backspace pressed, item input selected');
                }
            });
            itemInput.addEventListener('input', () => {
                const suggestionsDiv = document.getElementById('itemSuggestions');
                if (!suggestionsDiv) return;

                const query = itemInput.value.trim().toLowerCase();
                suggestionsDiv.innerHTML = '';
                suggestionsDiv.classList.add('hidden');

                if (!query) return;

                const filteredItems = items.filter(item => item.toLowerCase().includes(query));
                if (filteredItems.length > 0) {
                    suggestionsDiv.classList.remove('hidden');
                    filteredItems.forEach(item => {
                        const suggestion = document.createElement('div');
                        suggestion.className = 'p-2 hover:bg-teal-500 cursor-pointer';
                        suggestion.textContent = item;
                        suggestion.addEventListener('click', () => {
                            itemInput.value = item;
                            itemInput.select();
                            suggestionsDiv.classList.add('hidden');
                            const customerInput = document.getElementById('customer').value.trim();
                            const customerMatch = customerInput.match(/(.+)\s*\((.+)\)/);
                            if (customerMatch) {
                                const customerCode = customerMatch[2].trim().toUpperCase();
                                const target = customerTargets[customerCode]?.items[item.toUpperCase()] || 0;
                                const achieved = invoices
                                    .filter(inv => inv && inv.customerCode?.toUpperCase() === customerCode && inv.item?.toUpperCase() === item.toUpperCase() && !isNaN(Number(inv.quantity)))
                                    .reduce((sum, inv) => sum + Number(inv.quantity), 0);
                                document.getElementById('target').value = String(target);
                                document.getElementById('remaining').value = String(target - achieved);
                                console.log(`Item selected: customerCode=${customerCode}, item=${item}, target=${target}, achieved=${achieved}`);
                            }
                        });
                        suggestionsDiv.appendChild(suggestion);
                    });
                }
            });
        }

        const qtyInput = document.getElementById('bonusQty');
        const bonusInput = document.getElementById('bonusValue');
        if (qtyInput) qtyInput.addEventListener('input', calculateFromQty);
        if (bonusInput) bonusInput.addEventListener('input', calculateFromBonus);

        initSidebarNav();
        console.log('Initial Invoices:', invoices);
        console.log('Initial Done Targets:', doneTargets);
        console.log('Initial Customers:', customers);
        console.log('Initial Bonus Deals:', bonusDeals);
        renderInvoiceTable();
    } catch (error) {
        console.error('Initialization error:', error);
        alert('Failed to initialize app. Please clear cache and try again.');
    }
};

document.getElementById('excelFile')?.addEventListener('change', (event) => {
    if (isAppLocked) {
        alert('App is locked. Please unlock using the code.');
        event.target.value = '';
        return;
    }
    if (!getLoggedUser()) {
        alert('Please log in to upload a file.');
        event.target.value = '';
        return;
    }
    if (excelData.length > 0 && !confirm('Existing data will be replaced. Continue?')) {
        event.target.value = '';
        return;
    }
    const file = event.target.files[0];
    if (!file) {
        alert('Please select a file.');
        return;
    }
    const fileExtension = file.name.split('.').pop().toLowerCase();
    if (fileExtension === 'csv') {
        parseCSVandFilter(file, (data) => {
            excelData = data;
            localStorage.setItem('excelData', JSON.stringify(excelData));
            buildCustomerTargets();
            renderInvoiceTable();
            renderBonusDeals();
            populateBonusItems();
        });
    } else {
        alert('Please upload a valid CSV file.');
        event.target.value = '';
    }
});

function downloadTableImage() {
    const tableElement = document.querySelector("#customerTableBox table");
    html2canvas(tableElement, { scale: 2 }).then(canvas => {
        const link = document.createElement("a");
        link.download = "customer_full_report.png";
        link.href = canvas.toDataURL("image/png");
        link.click();
    });
}

function shareTableOnWhatsApp() {
    const tableElement = document.querySelector("#customerTableBox table"); // full table
    html2canvas(tableElement, { scale: 2 }).then(canvas => {
        const imageUrl = canvas.toDataURL("image/png");
        const blob = dataURLtoBlob(imageUrl);
        const file = new File([blob], "customer_full_report.png", { type: "image/png" });

        if (navigator.canShare && navigator.canShare({ files: [file] })) {
            navigator.share({
                files: [file],
                title: "Customer Report",
                text: "Here is the full customer report."
            });
        } else {
            alert("WhatsApp share not supported on this browser. Please download image and share manually.");
        }
    });
}

function shareTableOnEmail() {
    const tableElement = document.querySelector("#customerTableBox table"); // full table
    html2canvas(tableElement, { scale: 2 }).then(canvas => {
        const imageUrl = canvas.toDataURL("image/png");

        // image کو base64 کے ساتھ mailto میں attach نہیں کیا جا سکتا
        // اس لیے ہم صرف body میں link ڈال دیتے ہیں
        const subject = encodeURIComponent("Customer Report");
        const body = encodeURIComponent("Attached is the full customer report.\n\n") + imageUrl;

        window.location.href = `mailto:?subject=${subject}&body=${body}`;
    });
}

function dataURLtoBlob(dataurl) {
    const arr = dataurl.split(','), mime = arr[0].match(/:(.*?);/)[1],
        bstr = atob(arr[1]), n = bstr.length, u8arr = new Uint8Array(n);
    for (let i = 0; i < n; i++) u8arr[i] = bstr.charCodeAt(i);
    return new Blob([u8arr], { type: mime });
}

// --------------- My Sale: online implementation ---------------
function toggleSections() {
  document.getElementById("formGrid")?.classList.toggle("hidden");
  document.getElementById("customContent")?.classList.toggle("hidden");
}

let mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");

function showMySalePage() {
  const pages = ["mainPage", "allocationPage", "doneTargetPage", "bonusPage"];
  pages.forEach(id => document.getElementById(id)?.classList.add("hidden"));
  const salePage = document.getElementById("mySalePage");
  if (salePage) salePage.classList.remove("hidden");
  ["navInvoiceEntry","navAllocation","navDoneTargets","navBonus","navMySale","navMysale"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("bg-primary","text-white","bg-yellow-600");
  });
  const nav = document.getElementById("navMySale") || document.getElementById("navMysale");
  if (nav) nav.classList.add("bg-yellow-600","text-white");
  syncMySaleFromFirebase();
  renderSaleUploadHistory();
}

function renderMySaleTable() {
  mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
  const salePage = document.getElementById("mySalePage");
  if (!salePage) return;
  const companyTbody = salePage.querySelector("#mySaleCompanyTableBody") || salePage.querySelector("#mySaleTableBody");
  const dateTbody = salePage.querySelector("#mySaleDateTableBody");
  const companyTotalEl = salePage.querySelector("#mySaleCompanyTotal") || salePage.querySelector("#mySaleTotal");
  const dateTotalEl = salePage.querySelector("#mySaleDateTotal");
  const legacyTotalEl = salePage.querySelector("#mySaleTotal");
  setupMySaleDateInputs();

  const companyRowsSource = getMySaleRowsFromMonthStartToToday();
  const dateRowsSource = getMySaleRowsForSelectedDateRange();
  const companyMap = {};
  const dateCompanyMap = {};
  let companyGrandTotal = 0;
  let dateGrandTotal = 0;

  companyRowsSource.map(normalizeSaleRecord).forEach(sale => {
    const value = Number(sale.value) || 0;
    const summary = sale.summary || "";
    const company = sale.company || "Unknown Company";
    const companyKey = `${summary}||${company}`;
    if (!companyMap[companyKey]) companyMap[companyKey] = { summary, company, value: 0 };
    companyMap[companyKey].value += value;
    companyGrandTotal += value;
  });

  dateRowsSource.map(normalizeSaleRecord).forEach(sale => {
    const value = Number(sale.value) || 0;
    const company = sale.company || "Unknown Company";
    const dateKey = normalizeDateValue(sale.date) || "No Date";
    const dateCompanyKey = `${dateKey}||${company}`;
    if (!dateCompanyMap[dateCompanyKey]) dateCompanyMap[dateCompanyKey] = { date: dateKey, company, value: 0 };
    dateCompanyMap[dateCompanyKey].value += value;
    dateGrandTotal += value;
  });

  const companyRows = Object.values(companyMap)
    .sort((a, b) => (a.summary || "").localeCompare(b.summary || "", undefined, { numeric: true }) || a.company.localeCompare(b.company));
  const dateRows = Object.values(dateCompanyMap)
    .sort((a, b) => a.date.localeCompare(b.date) || a.company.localeCompare(b.company));

  if (companyTbody) {
    companyTbody.innerHTML = companyRows.length
      ? companyRows.map(row => `<tr>
          <td class="border p-2">${escapeHtml(row.summary)}</td>
          <td class="border p-2">${escapeHtml(row.company)}</td>
          <td class="border p-2 text-right font-semibold">${formatNumber(row.value)}</td>
        </tr>`).join("")
      : `<tr><td colspan="3" class="text-center p-3 text-gray-500">No company wise sale found</td></tr>`;
  }

  if (dateTbody) {
    dateTbody.innerHTML = dateRows.length
      ? dateRows.map(row => `<tr>
          <td class="border p-2">${escapeHtml(row.date)}</td>
          <td class="border p-2">${escapeHtml(row.company)}</td>
          <td class="border p-2 text-right font-semibold">${formatNumber(row.value)}</td>
        </tr>`).join("")
      : `<tr><td colspan="3" class="text-center p-3 text-gray-500">No date wise sale found</td></tr>`;
  }

  if (companyTotalEl) companyTotalEl.textContent = formatNumber(companyGrandTotal);
  if (dateTotalEl) dateTotalEl.textContent = formatNumber(dateGrandTotal);
  if (legacyTotalEl) legacyTotalEl.textContent = formatNumber(companyGrandTotal);
}

function formatNumber(n){ return Number(n).toLocaleString(); }
function escapeHtml(s){ return (s===undefined || s===null) ? "" : String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function getMonthStartToTodayRange() {
  const now = new Date();
  const today = formatLocalDateInput(now);
  return {
    from: today,
    to: today
  };
}
function formatLocalDateInput(date) {
  const local = new Date(date.getTime() - (date.getTimezoneOffset() * 60000));
  return local.toISOString().slice(0, 10);
}
function getMonthStartDateInput() {
  const now = new Date();
  return formatLocalDateInput(new Date(now.getFullYear(), now.getMonth(), 1));
}
function setupMySaleDateInputs() {
  const fromEl = document.getElementById("mySaleDateFrom");
  const toEl = document.getElementById("mySaleDateTo");
  const defaults = getMonthStartToTodayRange();
  if (!fromEl || !toEl) return defaults;
  if (!fromEl.value) fromEl.value = defaults.from;
  if (!toEl.value) toEl.value = defaults.to;
  return { from: fromEl.value, to: toEl.value };
}
function getMySaleRowsForSelectedDateRange() {
  const range = setupMySaleDateInputs();
  return getMySaleRowsForDateRange(range.from, range.to);
}
function getMySaleRowsFromMonthStartToToday() {
  const defaults = getMonthStartToTodayRange();
  return getMySaleRowsForDateRange(getMonthStartDateInput(), defaults.to);
}
function getMySaleRowsForDateRange(from, to) {
  const fromTime = from ? Date.parse(from) : -Infinity;
  const toTime = to ? Date.parse(to) : Infinity;
  return (mySaleData || []).filter(sale => {
    const date = normalizeDateValue(sale.date);
    if (!date) return true;
    const saleTime = Date.parse(date);
    if (isNaN(saleTime)) return true;
    return saleTime >= fromTime && saleTime <= toTime;
  });
}
function applyMySaleDateFilter() {
  const fromEl = document.getElementById("mySaleDateFrom");
  const toEl = document.getElementById("mySaleDateTo");
  if (fromEl?.value) localStorage.setItem("mySaleDateFrom", fromEl.value);
  if (toEl?.value) localStorage.setItem("mySaleDateTo", toEl.value);
  renderMySaleTable();
}
function resetMySaleDateFilter() {
  const defaults = getMonthStartToTodayRange();
  localStorage.setItem("mySaleDateFrom", defaults.from);
  localStorage.setItem("mySaleDateTo", defaults.to);
  const fromEl = document.getElementById("mySaleDateFrom");
  const toEl = document.getElementById("mySaleDateTo");
  if (fromEl) fromEl.value = defaults.from;
  if (toEl) toEl.value = defaults.to;
  renderMySaleTable();
}
function normalizeDateValue(value) {
  const raw = (value || "").toString().trim();
  if (!raw) return "";
  const parsed = Date.parse(raw);
  if (!isNaN(parsed)) return new Date(parsed).toISOString().slice(0, 10);
  const parts = raw.split(/[\/\-\.]/).map(part => part.trim());
  if (parts.length === 3) {
    const [a,b,c] = parts;
    if (c.length === 4) {
      const iso = `${c}-${b.padStart(2,"0")}-${a.padStart(2,"0")}`;
      if (!isNaN(Date.parse(iso))) return iso;
    }
  }
  return raw;
}
function pickLatestDate(a,b){ return normalizeDateValue(b) || normalizeDateValue(a) || b || a || ""; }
function normalizeSaleRecord(sale) {
  return {
    id: sale.id || `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    summary: (sale.summary || "").toString().trim(),
    company: (sale.company || "").toString().trim(),
    value: Number(sale.value) || 0,
    date: normalizeDateValue(sale.date) || new Date().toISOString().slice(0, 10),
    user: (sale.user || getActiveDataUser() || getLoggedUser() || "").toString().trim().toUpperCase(),
    createdAt: sale.createdAt || new Date().toISOString()
  };
}
function upsertLocalSaleRecords(records) {
  const byKey = {};
  (JSON.parse(localStorage.getItem("mySaleData") || "[]") || []).forEach(sale => {
    const clean = normalizeSaleRecord(sale);
    byKey[`${clean.user}|${clean.date}|${clean.summary}|${clean.company}`] = clean;
  });
  records.forEach(sale => {
    const clean = normalizeSaleRecord(sale);
    byKey[`${clean.user}|${clean.date}|${clean.summary}|${clean.company}`] = clean;
  });
  mySaleData = Object.values(byKey);
  localStorage.setItem("mySaleData", JSON.stringify(mySaleData));
  renderMySaleTable();
}
async function saveMySaleToFirebase() {
  try {
    if (typeof DATABASE_URL !== "string" || !DATABASE_URL) return;
    const current = JSON.parse(localStorage.getItem("mySaleData") || "[]");
    const byUser = {};
    current.map(normalizeSaleRecord).forEach(row => {
      const user = (row.user || getActiveDataUser()).toString().trim().toUpperCase();
      if (!user || user === "ALL") return;
      if (!byUser[user]) byUser[user] = [];
      byUser[user].push(row);
    });
    if (!Object.keys(byUser).length) {
      const user = (getActiveDataUser() || getLoggedUser() || "").toString().trim().toUpperCase();
      if (!user || user === "ALL") return;
      byUser[user] = [];
    }
    await Promise.all(Object.entries(byUser).map(([user, rows]) => {
      const payload = { uploadedAt: new Date().toISOString(), uploadedBy: getLoggedUser() || user, rows };
      return Promise.all([
        fetch(`${DATABASE_URL}/mySales/${user}/latest.json`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) }),
        fetch(`${DATABASE_URL}/csvUploads/${user}/mySales/latest.json`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) })
      ]);
    }));
  } catch (err) { console.error("My Sale Firebase save failed:", err); }
}
function getSaleRowsFromMainRows(rows) {
  const saleMap = {};
  (rows || []).map(normalizeMainRow).forEach(row => {
    const value = Number(row.Value) || 0;
    const summary = (row.SummaryNumber || "").toString().trim();
    const company = (row.CompanyName || "").toString().trim();
    const user = (row.User1 || row.User2 || "").toString().trim().toUpperCase();
    if (!user || user === "ADMIN" || user === "ALL" || !summary || !value) return;
    const date = normalizeDateValue(row.Date) || new Date().toISOString().slice(0, 10);
    const key = `${user}|${date}|${summary}|${company}`;
    if (!saleMap[key]) saleMap[key] = { summary, company, value: 0, date, user };
    saleMap[key].value += value;
  });
  return Object.values(saleMap);
}
function mergeSaleRecordArrays(...groups) {
  const byKey = {};
  groups.flat().map(normalizeSaleRecord).forEach(row => {
    const user = (row.user || "").toString().trim().toUpperCase();
    if (!user || user === "ADMIN" || user === "ALL") return;
    if (!row.summary && !row.company && !row.value) return;
    const key = `${user}|${row.date}|${row.summary}|${row.company}`;
    byKey[key] = row;
  });
  return Object.values(byKey);
}
async function saveMainCsvSalesToFirebase(rows, uploadedBy) {
  const saleRows = getSaleRowsFromMainRows(rows);
  if (!saleRows.length) return { rows: 0, users: 0 };
  const previous = JSON.parse(localStorage.getItem("mySaleData") || "[]");
  upsertLocalSaleRecords(saleRows);
  await saveMySaleToFirebase();
  const users = [...new Set(saleRows.map(row => row.user).filter(Boolean))];
  addSaleUploadHistory("main_csv_sales", saleRows);
  console.log(`Main CSV sales saved by ${uploadedBy || getLoggedUser() || "ADMIN"}: ${saleRows.length} rows / ${users.length} users.`);
  if ((getActiveDataUser() || "").toString().trim().toUpperCase() === "ALL") {
    mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
  } else if (previous.length) {
    mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
  }
  return { rows: saleRows.length, users: users.length };
}
async function fetchMySalePayloadForUser(user) {
  for (const url of [`${DATABASE_URL}/mySales/${user}/latest.json`, `${DATABASE_URL}/csvUploads/${user}/mySales/latest.json`]) {
    try {
      const res = await fetch(url);
      if (!res.ok) continue;
      const json = await res.json();
      if (json && Array.isArray(json.rows)) return json;
    } catch (err) { console.warn("My Sale fetch path skipped:", err); }
  }
  return null;
}
function collectMySaleRowsFromNode(node, output) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node.rows)) {
    output.push(...node.rows);
    return;
  }
  Object.values(node).forEach(child => collectMySaleRowsFromNode(child, output));
}
async function fetchAllMySaleRows() {
  const byKey = {};
  const addRows = (rows, fallbackUser = "") => {
    (rows || []).forEach(row => {
      const clean = normalizeSaleRecord({ ...row, user: row.user || fallbackUser });
      if (!clean.user || clean.user === "ADMIN" || clean.user === "ALL") return;
      if (!clean.summary && !clean.company && !clean.value) return;
      const key = `${clean.user}|${clean.date}|${clean.summary}|${clean.company}`;
      byKey[key] = clean;
    });
  };
  const collectByUser = (json, nodeSelector) => {
    if (!json || typeof json !== "object") return;
    Object.entries(json).forEach(([user, node]) => {
      const cleanUser = (user || "").toString().trim().toUpperCase();
      if (!cleanUser || cleanUser === "ADMIN" || cleanUser === "ALL") return;
      const saleNode = nodeSelector(node);
      const rows = [];
      collectMySaleRowsFromNode(saleNode, rows);
      addRows(rows, cleanUser);
    });
  };

  try {
    const res = await fetch(`${DATABASE_URL}/mySales.json`);
    if (res.ok) collectByUser(await res.json(), node => node);
  } catch (err) {
    console.warn("All My Sale /mySales fetch skipped:", err);
  }

  try {
    const res = await fetch(`${DATABASE_URL}/csvUploads.json`);
    if (res.ok) collectByUser(await res.json(), node => node?.mySales);
  } catch (err) {
    console.warn("All My Sale /csvUploads mySales fetch skipped:", err);
  }

  try {
    const res = await fetch(`${DATABASE_URL}/csvUploads/ALL/latest.json`);
    if (res.ok) {
      const json = await res.json();
      if (json && Array.isArray(json.rows)) addRows(getSaleRowsFromMainRows(json.rows));
    }
  } catch (err) {
    console.warn("All My Sale central CSV sale fallback skipped:", err);
  }
  return Object.values(byKey);
}
async function syncMySaleFromFirebase(onDone, forceUser = null) {
  try {
    const currentUser = (getLoggedUser() || "").toString().trim().toUpperCase();
    if (!currentUser || typeof DATABASE_URL !== "string" || !DATABASE_URL) { renderMySaleTable(); if(onDone) onDone(mySaleData); return; }
    const targetUser = currentUser === "ADMIN" ? (forceUser || getActiveDataUser() || currentUser).toString().trim().toUpperCase() : currentUser;
    if (currentUser === "ADMIN" && targetUser === "ALL") {
      const firebaseSales = await fetchAllMySaleRows();
      const localAllRows = [
        ...(JSON.parse(localStorage.getItem("excelDataAll") || "[]") || []),
        ...(JSON.parse(localStorage.getItem("excelData") || "[]") || []),
        ...(Array.isArray(excelData) ? excelData : [])
      ];
      const localSales = getSaleRowsFromMainRows(localAllRows);
      mySaleData = mergeSaleRecordArrays(firebaseSales, localSales);
      localStorage.setItem("mySaleData", JSON.stringify(mySaleData));
    } else {
      const json = await fetchMySalePayloadForUser(targetUser);
      if (json && Array.isArray(json.rows)) {
        mySaleData = json.rows.map(normalizeSaleRecord);
        localStorage.setItem("mySaleData", JSON.stringify(mySaleData));
      } else {
        mySaleData = [];
        localStorage.setItem("mySaleData", JSON.stringify(mySaleData));
      }
    }
  } catch (err) { console.warn("My Sale Firebase sync skipped:", err); }
  renderMySaleTable();
  if (onDone) onDone(mySaleData);
}
function addManualSale() {
  const summary = document.getElementById("manualSaleNumber")?.value || "";
  const company = document.getElementById("manualSaleCompany")?.value || "";
  const value = Number(document.getElementById("manualSaleValue")?.value || 0);
  const date = document.getElementById("manualSaleDate")?.value || new Date().toISOString().slice(0, 10);
  if (!summary.trim() || !company.trim() || !value) { alert("Please enter serial/company number, company and sale value."); return; }
  const record = normalizeSaleRecord({ summary, company, value, date });
  upsertLocalSaleRecords([record]);
  saveMySaleToFirebase();
  showAppNotification("Manual sale saved successfully.", "success");
  ["manualSaleNumber", "manualSaleCompany", "manualSaleValue"].forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
}
function normalizeSaleHeader(value) { return (value || "").toString().trim().toLowerCase().replace(/\s+/g, "").replace(/[^a-z0-9]/g, ""); }
function getSaleCell(row, headerMap, names, fallbackIndex) {
  for (const name of names) { const index = headerMap[normalizeSaleHeader(name)]; if (index !== undefined && row[index] !== undefined) return row[index]; }
  return row[fallbackIndex] ?? "";
}
function parseSaleCsvRecords(rows) {
  if (!rows || !rows.length) return [];
  const firstRow = rows[0] || [];
  const headerMap = {};
  firstRow.forEach((cell, index) => { const key = normalizeSaleHeader(cell); if (key) headerMap[key] = index; });
  const hasHeader = headerMap.value !== undefined || headerMap.company !== undefined || headerMap.summery !== undefined || headerMap.summary !== undefined || headerMap.serialnum !== undefined || headerMap.serial !== undefined;
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const summaryFallback = hasHeader ? 4 : (firstRow.length >= 14 ? 10 : 4);
  const companyFallback = hasHeader ? 5 : (firstRow.length >= 14 ? 11 : 5);
  const valueFallback = hasHeader ? 6 : (firstRow.length >= 14 ? 12 : 6);
  const dateFallback = hasHeader ? 7 : (firstRow.length >= 14 ? 13 : 7);
  return dataRows.map(row => {
    const summary = getSaleCell(row, headerMap, ["serialnum","serialnumber","serial","srno","sr","companynumber","summarynumber","summery","summary"], summaryFallback).toString().trim();
    const company = getSaleCell(row, headerMap, ["company","companyname"], companyFallback).toString().trim();
    const valueRaw = getSaleCell(row, headerMap, ["value","sale"], valueFallback).toString().trim();
    const date = getSaleCell(row, headerMap, ["date","tilldate","tiltodate"], dateFallback).toString().trim();
    const user1 = getSaleCell(row, headerMap, ["user","user1"], 0).toString().trim();
    const user2 = getSaleCell(row, headerMap, ["user2"], 1).toString().trim();
    const value = parseFloat(valueRaw.replace(/,/g, ""));
    if (!summary || isNaN(value)) return null;
    return { summary, company, value, date, user: user1 || user2 || getActiveDataUser() };
  }).filter(Boolean);
}
function addSaleUploadHistory(fileName, records) {
  const history = JSON.parse(localStorage.getItem("saleUploadHistory") || "[]");
  const total = records.reduce((sum, item) => sum + (Number(item.value) || 0), 0);
  history.unshift({ fileName: fileName || "sale.csv", rows: records.length, total, uploadedAt: new Date().toLocaleString() });
  localStorage.setItem("saleUploadHistory", JSON.stringify(history.slice(0,25)));
  renderSaleUploadHistory();
}
function renderSaleUploadHistory() {
  const box = document.getElementById("saleUploadHistory");
  if (!box) return;
  const history = JSON.parse(localStorage.getItem("saleUploadHistory") || "[]");
  if (!history.length) { box.innerHTML = "No upload record"; return; }
  box.innerHTML = history.map(item => `<div class="flex flex-wrap justify-between gap-2 border-b py-2 last:border-b-0"><span class="font-semibold">${escapeHtml(item.fileName)}</span><span>${escapeHtml(item.uploadedAt)}</span><span>Rows: ${formatNumber(item.rows)}</span><span>Total: ${formatNumber(item.total)}</span></div>`).join("");
}
function clearSaleUploadHistory() { if (!confirm("Clear CSV upload record? Sale table data will remain.")) return; localStorage.removeItem("saleUploadHistory"); renderSaleUploadHistory(); }
function processSaleCsvRows(rows) {
  const records = parseSaleCsvRecords(rows);
  const batch = {};
  records.forEach(record => {
    const key = `${(record.user || getActiveDataUser() || "").toString().trim().toUpperCase()}|${record.date || ""}|${record.summary}|${record.company}`;
    if (!batch[key]) batch[key] = { summary: record.summary, company: record.company, value: 0, date: record.date, user: record.user || getActiveDataUser() };
    batch[key].value += Number(record.value);
    batch[key].company = record.company || batch[key].company;
    batch[key].date = pickLatestDate(batch[key].date, record.date);
  });
  upsertLocalSaleRecords(Object.values(batch));
  saveMySaleToFirebase();
  return records;
}
function handleSaleCsvFileChange(e) {
  const file = e?.target?.files?.[0];
  if (!file) return;
  Papa.parse(file, { skipEmptyLines: true, complete: function(results) { const records = processSaleCsvRows(results.data || []); addSaleUploadHistory(file.name, records || []); showAppNotification("My Sale CSV uploaded successfully.", "success"); }, error: function(err){ showAppNotification("CSV parse error: " + err.message, "error"); } });
  e.target.value = "";
}
function resetMySale() {
  const password = prompt("Enter password to reset My Sale data:");
  if (password !== "985973") { alert("Wrong password! Reset cancelled."); return; }
  if (!confirm("Are you sure you want to reset My Sale data?")) return;
  mySaleData = [];
  localStorage.removeItem("mySaleData");
  renderMySaleTable();
  saveMySaleToFirebase();
}
document.addEventListener("DOMContentLoaded", () => {
  const saleInput = document.getElementById("saleCsvFile");
  if (saleInput) { saleInput.removeEventListener("change", handleSaleCsvFileChange); saleInput.addEventListener("change", handleSaleCsvFileChange); }
  const nav = document.getElementById("navMySale") || document.getElementById("navMysale");
  if (nav) { nav.removeEventListener("click", showMySalePage); nav.addEventListener("click", showMySalePage); }
  const refreshBtn = document.getElementById("refreshMySale");
  if (refreshBtn) refreshBtn.addEventListener("click", () => syncMySaleFromFirebase(() => { renderSaleUploadHistory(); showAppNotification("My Sale data refreshed.", "success"); }));
  renderMySaleTable();
  renderSaleUploadHistory();
});
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
const firebaseConfig = {
  apiKey: "AIzaSyCY3r6YelbzubGJvw_dv5ZMM-7bZ7ebsw0",
  authDomain: "saniapp-7b4be.firebaseapp.com",
  databaseURL: "https://saniapp-7b4be-default-rtdb.firebaseio.com",
  projectId: "saniapp-7b4be",
  storageBucket: "saniapp-7b4be.firebasestorage.app",
  messagingSenderId: "928447805197",
  appId: "1:928447805197:web:c83681948e2e52e15da244",
  measurementId: "G-FTWSNVGVY1"
};
// example: put this near your firebaseConfig object
const DATABASE_URL = "https://saniapp-7b4be-default-rtdb.firebaseio.com"; // <-- replace with your Realtime DB URL (no trailing slash)


function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = JSON.parse(e.postData.contents);
  data.forEach(row => sheet.appendRow(Object.values(row)));
  return ContentService.createTextOutput("OK");
}

function saveCSVOnline(csvData) {
  fetch("https://api.jsonbin.io/v3/b", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Master-Key": "YOUR_JSONBIN_KEY"
    },
    body: JSON.stringify(csvData)
  })
  .then(res => res.json())
  .then(data => console.log("✅ Saved online at:", data.metadata.id))
  .catch(err => console.error(err));
}

function uploadToFirebase(data) {
  fetch("https://YOUR_PROJECT_ID.firebaseio.com/csvData.json", {
    method: "PUT",
    body: JSON.stringify(data)
  })
  .then(res => res.json())
  .then(() => showAppNotification("Data uploaded to Firebase.", "success"))
  .catch(console.error);
}
document.getElementById('excelFile')?.addEventListener('change', (event) => {
  if (isAppLocked) {
    alert('App is locked. Please unlock using the code.');
    event.target.value = '';
    return;
  }
  if (!getLoggedUser()) {
    alert('Please log in to upload a file.');
    event.target.value = '';
    return;
  }
  if (excelData.length > 0 && !confirm('Existing data will be replaced. Continue?')) {
    event.target.value = '';
    return;
  }
  const file = event.target.files[0];
  if (!file) {
    alert('Please select a file.');
    return;
  }
  const fileExtension = file.name.split('.').pop().toLowerCase();
  if (fileExtension === 'csv') {
    parseCSVandFilter(file, (data) => {
      excelData = data;
      localStorage.setItem('excelData', JSON.stringify(excelData));
      buildCustomerTargets();
      renderInvoiceTable();
      renderBonusDeals();
      populateBonusItems();
      // Sync data after upload
      const uploadSyncUser = getActiveDataUser();
      syncUserDataFromFirebase(() => {
        console.log('Data synced after CSV upload');
      }, uploadSyncUser);
    });
  } else {
    alert('Please upload a valid CSV file.');
    event.target.value = '';
  }
});



/* -----------------------------------------------------------------
   ✅ FUNCTION: Process Firebase JSON like CSV upload
------------------------------------------------------------------*/

/* ================================================================
   ✅ FIREBASE SYNC SYSTEM v3.5 (Custom CSV Structure)
   Works with: City, CustomerCode, Customer, Item1, Target1, Achieve1, ...
================================================================ */


// ✅ MERGE UPDATE MODE — keeps old Target, only updates Achieve/Value
async function saveCSVToFirebase(data) {
  try {
    if (!Array.isArray(data) || data.length === 0) {
      console.warn("⚠️ No data to upload.");
      return;
    }

    const loggedUser = getLoggedUser();
    if (!loggedUser) return alert("⚠️ Please log in first!");
    if (!DATABASE_URL) {
      console.warn("⚠️ DATABASE_URL missing, saving locally only.");
      localStorage.setItem("excelData", JSON.stringify(data));
      return;
    }

    const targetUploadUser = getActiveDataUser() || loggedUser.toUpperCase();
    const path = `csvUploads/${targetUploadUser}/latest`;
    const url = `${DATABASE_URL}/${path}.json`;

    // --- Step 1: Fetch old Firebase data ---
    let oldRows = [];
    try {
      const res = await fetch(url);
      if (res.ok) {
        const json = await res.json();
        if (json && Array.isArray(json.rows)) oldRows = json.rows;
      }
    } catch {
      console.warn("ℹ️ No previous Firebase data found.");
    }

    // --- Step 2: Merge by CustomerCode + Item1 ---
    const merged = [...oldRows];
    for (const newRow of data) {
      const code = (newRow.CustomerCode || "").trim().toUpperCase();
      const item = (newRow.Item1 || "").trim().toUpperCase();

      const idx = merged.findIndex(
        r =>
          (r.CustomerCode || "").trim().toUpperCase() === code &&
          (r.Item1 || "").trim().toUpperCase() === item
      );

      const clean = {
        ...newRow,
        Target1: parseInt(newRow.Target1) || 0,
        Achieve1: parseInt(newRow.Achieve1) || 0,
        DealQty: parseInt(newRow.DealQty) || 0,
        DealBonus: parseInt(newRow.DealBonus) || 0,
        Value: parseFloat((newRow.Value || "0").toString().replace(/,/g, "")) || 0,
        ItemRate: parseFloat((newRow.ItemRate || "0").toString().replace(/,/g, "")) || 0
      };

      if (idx >= 0) {
        // 🔄 Only update Achieve/Value fields, keep old Target
        merged[idx] = {
          ...merged[idx],
          Achieve1: clean.Achieve1 || merged[idx].Achieve1,
          Value: clean.Value || merged[idx].Value,
          DealQty: clean.DealQty || merged[idx].DealQty,
          DealBonus: clean.DealBonus || merged[idx].DealBonus,
          Date: clean.Date || merged[idx].Date
        };
      } else {
        // New row — add completely
        merged.push(clean);
      }
    }

    const payload = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: getLoggedUser() || loggedUser,
      rows: merged,
    };

    const putRes = await fetch(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (putRes.ok) {
      console.log(`✅ Merged ${merged.length} rows into Firebase.`);
      localStorage.setItem("excelData", JSON.stringify(merged));
    } else {
      console.error("❌ Upload failed:", putRes.status);
    }
  } catch (err) {
    console.error("❌ saveCSVToFirebase Error:", err);
  }
}




// 🔹 2. PROCESS FIREBASE JSON — fixes missing Achieve values
// ✅ FIX: Firebase JSON -> proper object reading
/* ================================================================
   ✅ ROBUST FIREBASE SYNC & CSV PROCESSING (FINAL)
   Place this block at the END of script.js (replace old funcs)
================================================================ */

// --- Utility: safe int/float parsers
function parseSafeInt(v) {
  if (v === null || v === undefined) return 0;
  const s = typeof v === "number" ? String(v) : v.toString();
  const n = parseInt(s.replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}
function parseSafeFloat(v) {
  if (v === null || v === undefined) return 0;
  const s = typeof v === "number" ? String(v) : v.toString();
  const n = parseFloat(s.replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function normalizeMainRow(row) {
  return {
    City: row.City || "",
    CustomerCode: (row.CustomerCode || "").toString().trim().toUpperCase(),
    Customer: row.Customer || "",
    Item1: (row.Item1 || "").toString().trim().toUpperCase(),
    Target1: parseSafeInt(row.Target1),
    Achieve1: parseSafeInt(row.Achieve1),
    User1: (row.User1 || "").toString().trim().toUpperCase(),
    User2: (row.User2 || "").toString().trim().toUpperCase(),
    DealQty: parseSafeInt(row.DealQty),
    DealBonus: parseSafeInt(row.DealBonus),
    SummaryNumber: row.SummaryNumber || "",
    CompanyName: row.CompanyName || "",
    Value: parseSafeFloat(row.Value),
    Date: row.Date || "",
    ItemRate: parseSafeFloat(row.ItemRate)
  };
}

function getRowUsers(row) {
  return [...new Set([row.User1, row.User2]
    .map(user => (user || "").toString().trim().toUpperCase())
    .filter(user => user && user !== "ADMIN" && user !== "ALL"))];
}

function filterRowsForUser(rows, user) {
  const cleanUser = (user || "").toString().trim().toUpperCase();
  const cleanRows = (rows || []).map(normalizeMainRow);
  if (!cleanUser || cleanUser === "ADMIN" || cleanUser === "ALL") return cleanRows;
  return cleanRows.filter(row => getRowUsers(row).includes(cleanUser));
}

function getUsersFromRows(rows) {
  const users = new Set();
  (rows || []).forEach(row => getRowUsers(row).forEach(user => users.add(user)));
  return [...users].sort((a, b) => a.localeCompare(b));
}

async function putCsvRowsForUser(user, rows, uploadedBy) {
  const cleanUser = (user || "").toString().trim().toUpperCase();
  if (!cleanUser || typeof DATABASE_URL !== "string" || DATABASE_URL.length === 0) return false;
  const payload = {
    uploadedAt: new Date().toISOString(),
    uploadedBy: uploadedBy || getLoggedUser() || cleanUser,
    rows: rows.map(normalizeMainRow)
  };
  const url = `${DATABASE_URL}/csvUploads/${cleanUser}/latest.json`;
  const res = await fetch(url, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!res.ok) throw new Error(`Upload failed for ${cleanUser}: ${res.status}`);
  return true;
}

function getMainRowMergeKey(row) {
  const clean = normalizeMainRow(row);
  const users = getRowUsers(clean).join(",");
  return `${clean.CustomerCode}|${clean.Item1}|${users}`;
}

function mergeCsvRowsKeepTargets(existingRows, incomingRows) {
  const merged = {};
  (existingRows || []).map(normalizeMainRow).forEach(row => {
    const key = getMainRowMergeKey(row);
    if (key.replace(/\|/g, "")) merged[key] = row;
  });

  (incomingRows || []).map(normalizeMainRow).forEach(row => {
    const key = getMainRowMergeKey(row);
    if (!key.replace(/\|/g, "")) return;
    const existing = merged[key];
    if (existing) {
      merged[key] = {
        ...existing,
        City: row.City || existing.City,
        Customer: row.Customer || existing.Customer,
        Target1: row.Target1 > 0 ? row.Target1 : (Number(existing.Target1) || 0),
        Achieve1: row.Achieve1,
        User1: row.User1 || existing.User1,
        User2: row.User2 || existing.User2,
        DealQty: row.DealQty,
        DealBonus: row.DealBonus,
        SummaryNumber: row.SummaryNumber || existing.SummaryNumber,
        CompanyName: row.CompanyName || existing.CompanyName,
        Value: row.Value,
        Date: row.Date || existing.Date,
        ItemRate: row.ItemRate || existing.ItemRate || 0
      };
    } else {
      merged[key] = row;
    }
  });
  return Object.values(merged);
}

async function fetchExistingCsvRowsForUser(user) {
  try {
    if (typeof DATABASE_URL !== "string" || !DATABASE_URL) return [];
    const cleanUser = (user || "").toString().trim().toUpperCase();
    if (!cleanUser) return [];
    const res = await fetch(`${DATABASE_URL}/csvUploads/${cleanUser}/latest.json`);
    if (!res.ok) return [];
    const json = await res.json();
    return Array.isArray(json?.rows) ? json.rows : [];
  } catch (err) {
    console.warn("Existing CSV rows fetch skipped:", err);
    return [];
  }
}

async function saveAdminFullCsvToFirebase(rows, loggedUser) {
  const cleanRows = rows.map(normalizeMainRow);
  const existingAllRows = await fetchExistingCsvRowsForUser("ALL");
  const mergedAllRows = mergeCsvRowsKeepTargets(existingAllRows, cleanRows);
  await putCsvRowsForUser("ALL", mergedAllRows, loggedUser);
  const users = getUsersFromRows(mergedAllRows);
  for (const user of users) {
    const userIncoming = filterRowsForUser(cleanRows, user);
    const userExisting = await fetchExistingCsvRowsForUser(user);
    const userMerged = mergeCsvRowsKeepTargets(userExisting, userIncoming);
    await putCsvRowsForUser(user, userMerged, loggedUser);
  }
  await saveMainCsvSalesToFirebase(mergedAllRows, loggedUser);
  localStorage.setItem("excelDataAll", JSON.stringify(mergedAllRows));
  localStorage.setItem("excelData", JSON.stringify(mergedAllRows));
  localStorage.setItem("lastCsvUploadRef", `${DATABASE_URL}/csvUploads/ALL/latest.json`);
  console.log(`Central CSV uploaded/merged: ${mergedAllRows.length} rows, ${users.length} user copies.`);
  return { rows: mergedAllRows.length, users: users.length };
}

// ----------------- SAVE (MERGE MODE -> keeps old Target1) -----------------
async function saveCSVToFirebase(data) {
  try {
    if (!Array.isArray(data) || data.length === 0) {
      console.warn("⚠️ saveCSVToFirebase: No data to save.");
      return;
    }

    const loggedUser = getLoggedUser();
    if (!loggedUser) {
      console.warn("⚠️ saveCSVToFirebase: No logged-in user. Saving local only.");
      localStorage.setItem("excelData", JSON.stringify(data));
      return;
    }

    if (typeof DATABASE_URL !== "string" || DATABASE_URL.length === 0) {
      console.warn("⚠️ saveCSVToFirebase: DATABASE_URL missing. Saving local only.");
      localStorage.setItem("excelData", JSON.stringify(data));
      return;
    }

    if (loggedUser.toString().trim().toUpperCase() === "ADMIN") {
      const result = await saveAdminFullCsvToFirebase(data, loggedUser);
      console.log(`Admin central upload complete: ${result.rows} rows / ${result.users} users.`);
      return;
    }
    const targetUploadUser = getActiveDataUser() || loggedUser.toUpperCase();
    const path = `csvUploads/${targetUploadUser}/latest`;
    const url = `${DATABASE_URL}/${path}.json`;

    // 1) Fetch existing latest (if any)
    let existingRows = [];
    try {
      const res = await fetch(url);
      if (res.ok) {
        const json = await res.json();
        if (json && Array.isArray(json.rows)) existingRows = json.rows;
      }
    } catch (err) {
      console.warn("ℹ️ saveCSVToFirebase: No existing latest found or fetch error.", err);
    }

    const mergedArray = mergeCsvRowsKeepTargets(existingRows, data);

    // 4) Upload into latest.json (overwrite safely)
    const payload = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: getLoggedUser() || loggedUser,
      rows: mergedArray
    };

    const putRes = await fetch(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (putRes.ok) {
      console.log(`✅ saveCSVToFirebase: Uploaded ${mergedArray.length} rows to ${url}`);
      showAppNotification("CSV data uploaded successfully.", "success");
      localStorage.setItem("excelData", JSON.stringify(mergedArray));
      // also keep lastCsvUploadRef for debugging
      localStorage.setItem("lastCsvUploadRef", url);
    } else {
      console.error("❌ saveCSVToFirebase: Upload failed:", putRes.status);
      // fallback: save locally
      localStorage.setItem("excelData", JSON.stringify(mergedArray));
    }
  } catch (err) {
    console.error("❌ saveCSVToFirebase Error:", err);
    // fallback: save locally
    try { localStorage.setItem("excelData", JSON.stringify(data)); } catch(e){}
  }
}

// --- PROCESS Firebase JSON into CSV-like rows (object-based mapping) ---
function processJSONFromFirebase(jsonData, onDone) {
  try {
    if (!jsonData || !Array.isArray(jsonData.rows)) {
      console.warn("⚠️ processJSONFromFirebase: Invalid or empty firebase JSON.");
      if (onDone) onDone([]);
      return;
    }

    const rows = jsonData.rows.map(r => ({
      City: r.City || "",
      CustomerCode: (r.CustomerCode || "").toString().trim().toUpperCase(),
      Customer: r.Customer || "",
      Item1: (r.Item1 || "").toString().trim().toUpperCase(),
      Target1: parseSafeInt(r.Target1),
      Achieve1: parseSafeInt(r.Achieve1),
      User1: r.User1 || "",
      User2: r.User2 || "",
      DealQty: parseSafeInt(r.DealQty),
      DealBonus: parseSafeInt(r.DealBonus),
      SummaryNumber: r.SummaryNumber || "",
      CompanyName: r.CompanyName || "",
      Value: parseSafeFloat(r.Value),
      Date: r.Date || "",
      ItemRate: parseSafeFloat(r.ItemRate)
    }));

    console.log(`✅ processJSONFromFirebase: cleaned ${rows.length} rows.`);
    // Persist local backup
    localStorage.setItem("excelData", JSON.stringify(rows));
    // Pass to common processing to render
    processCSVData(rows, onDone);
  } catch (err) {
    console.error("❌ processJSONFromFirebase Error:", err);
    if (onDone) onDone([]);
  }
}

// ----------------- PROCESS CSV DATA (common for CSV & Firebase) -----------------
function processCSVData(data, onDone) {
  try {
    if (!Array.isArray(data)) data = [];

    // Normalize fields types and defaults
    excelData = data.map(r => ({
      City: r.City || "",
      CustomerCode: (r.CustomerCode || "").toString().trim().toUpperCase(),
      Customer: r.Customer || "",
      Item1: (r.Item1 || "").toString().trim().toUpperCase(),
      Target1: parseSafeInt(r.Target1),
      Achieve1: parseSafeInt(r.Achieve1),
      User1: r.User1 || "",
      User2: r.User2 || "",
      DealQty: parseSafeInt(r.DealQty),
      DealBonus: parseSafeInt(r.DealBonus),
      SummaryNumber: r.SummaryNumber || "",
      CompanyName: r.CompanyName || "",
      Value: parseSafeFloat(r.Value),
      Date: r.Date || "",
      ItemRate: parseSafeFloat(r.ItemRate)
    }));

    // Save backup
    localStorage.setItem("excelData", JSON.stringify(excelData));

    // Recompute invoices (these are based on Achieve1)
   // ✅ Recompute invoices (include all rows so Target Value works)
invoices = excelData
  .filter(r => r.CustomerCode && r.Item1)
  .map(r => ({
    city: r.City,
    customerCode: r.CustomerCode,
    customer: r.Customer,
    item: r.Item1,
    target: r.Target1,
    quantity: r.Achieve1 || 0, // zero allowed
    rate: r.ItemRate || 0,
    user: r.User1 || r.User2 || getLoggedUser() || ""
  }));

    localStorage.setItem("invoices", JSON.stringify(invoices));

    // Recompute bonusDeals
    bonusDeals = {};
    excelData.forEach(row => {
      const item = row.Item1;
      if (!item) return;
      if (!bonusDeals[item]) bonusDeals[item] = [];
      if (row.DealQty > 0 || row.DealBonus > 0) {
        bonusDeals[item].push({ qty: row.DealQty, bonus: row.DealBonus });
      }
    });
    localStorage.setItem("bonusDeals", JSON.stringify(bonusDeals));

    // Build customer targets and UI data
    buildCustomerTargets();

    // Render UI pieces (ensure these functions exist)
    if (typeof renderInvoiceTable === "function") renderInvoiceTable();
    if (typeof renderMySaleTable === "function") renderMySaleTable();
    if (typeof renderBonusDeals === "function") renderBonusDeals();
    if (typeof populateBonusItems === "function") populateBonusItems();

    if (onDone) onDone(excelData);
  } catch (err) {
    console.error("❌ processCSVData Error:", err);
    if (onDone) onDone([]);
  }
}

// ----------------- SYNC from Firebase (loads only latest.json) -----------------
// =======================
// Upgraded sync with ADMIN override
// Replace existing syncUserDataFromFirebase with this version
// =======================
async function syncUserDataFromFirebase(onDone, forceUser = null) {
  try {
    const loggedUser = getLoggedUser();
    if (!loggedUser) {
      console.warn('⚠️ syncUserDataFromFirebase: No logged-in user.');
      if (onDone) onDone([]);
      return;
    }

    // Normalize usernames to uppercase for paths
    const currentUser = loggedUser.toString().trim().toUpperCase();

    // If caller provided a forceUser AND current logged user is ADMIN -> allow override
    let targetUser = null;
    if (forceUser && currentUser === 'ADMIN') {
      targetUser = forceUser.toString().trim().toUpperCase();
      setActiveDataUser(targetUser);
      console.log(`🔐 ADMIN override: syncing data for user ${targetUser}`);
    } else {
      // default: sync for the logged-in user only
      targetUser = currentUser;
      setActiveDataUser(targetUser);
    }

    if (typeof DATABASE_URL !== "string" || DATABASE_URL.length === 0) {
      console.warn("⚠️ syncUserDataFromFirebase: DATABASE_URL missing. Loading local data.");
      const local = JSON.parse(localStorage.getItem("excelData") || "[]");
      processCSVData(local, onDone);
      return;
    }

    // Fetch all uploads for the target user (not latest.json only)
    const url = `${DATABASE_URL}/csvUploads/${targetUser}.json`;
    console.log("🔄 syncUserDataFromFirebase: fetching", url);

    const res = await fetch(url);
    if (!res.ok) {
      console.warn("⚠️ syncUserDataFromFirebase: fetch returned", res.status);
      const local = JSON.parse(localStorage.getItem("excelData") || "[]");
      processCSVData(local, onDone);
      return;
    }

    const json = await res.json();
    // If ADMIN override used, json may contain multiple upload nodes; flatten them into rows
    let allRows = [];
    if (json) {
      // If structure is { latest: { rows: [...] }, otherUpload: { rows: [...] }, ... }
      const values = Object.values(json);
      values.forEach(v => {
        if (!v) return;
        if (Array.isArray(v.rows)) {
          allRows = allRows.concat(v.rows);
        } else if (Array.isArray(v)) {
          // sometimes direct array stored
          allRows = allRows.concat(v);
        } else if (v.rows && Array.isArray(v.rows)) {
          allRows = allRows.concat(v.rows);
        }
      });
    }

    if (!allRows || allRows.length === 0) {
      console.warn('⚠️ syncUserDataFromFirebase: No rows found for', targetUser, '- falling back to local.');
      const local = JSON.parse(localStorage.getItem("excelData") || "[]");
      processCSVData(local, onDone);
      return;
    }

    console.log(`✅ syncUserDataFromFirebase: fetched ${allRows.length} rows for user ${targetUser}`);
    // processCSVData will normalize and render; it expects an array of row objects
    processCSVData(allRows, onDone);
  } catch (err) {
    console.error("❌ syncUserDataFromFirebase Error:", err);
    const local = JSON.parse(localStorage.getItem("excelData") || "[]");
    processCSVData(local, onDone);
  }
}

// =======================
// Helper to sync currently-selected user in UI (works without changing HTML)
// - tries common select element ids, otherwise asks via prompt
// Attach this to your Sync button if you want admin to pick from UI.
// =======================
function syncSelectedUser() {
  const btn = document.getElementById('syncBtn');
  if (btn) {
    btn.innerText = "⏳ Syncing...";
    btn.disabled = true;
  }

  // try different possible element ids that might exist in your HTML
  const possibleIds = ['userSelect', 'userFilter', 'userDropdown', 'userList', 'user'];
  let selected = null;
  for (const id of possibleIds) {
    const el = document.getElementById(id);
    if (el) {
      selected = (el.value || el.options?.[el.selectedIndex]?.value || '').toString().trim();
      if (selected) break;
    }
  }

  // fallback: if not found or empty, ask admin
  if (!selected) {
    // only prompt if current user is ADMIN (otherwise we should not allow override)
    const currentUser = (getLoggedUser() || '').toString().trim().toUpperCase();
    if (currentUser === 'ADMIN') {
      selected = prompt('Enter username to sync (e.g. ND, ASIF). Leave empty to sync ADMIN data:');
      if (!selected) selected = null;
    }
  }

  // Call sync with override only when ADMIN selected someone
  const currentUser = (getLoggedUser() || '').toString().trim().toUpperCase();
  if (currentUser === 'ADMIN' && selected) {
    syncUserDataFromFirebase(async () => {
      await syncMySaleFromFirebase?.(null, selected);
      if (btn) {
        showAppNotification('Data synced successfully for ' + selected.toUpperCase() + '.', 'success');
        btn.innerText = "🔄 Sync Data";
        btn.disabled = false;
        renderInvoiceTable();
      }
    }, selected);
  } else {
    // normal sync (no override)
    syncUserDataFromFirebase(async () => {
      await syncMySaleFromFirebase?.();
      if (btn) {
        showAppNotification('Data synced successfully.', 'success');
        btn.innerText = "🔄 Sync Data";
        btn.disabled = false;
        renderInvoiceTable();
      }
    });
  }
}

let startupSyncPromptShown = false;

function getSelectedSyncUserForPrompt() {
  const ids = ["userSelect", "userFilter", "userDropdown", "userList", "user"];
  for (const id of ids) {
    const el = document.getElementById(id);
    const value = (el?.value || el?.options?.[el.selectedIndex]?.value || "").toString().trim();
    if (value) return value.toUpperCase();
  }
  return "";
}

function showStartupSyncPrompt() {
  if (startupSyncPromptShown || document.getElementById("startupSyncPrompt")) return;
  if (!(getLoggedUser && getLoggedUser())) return;
  startupSyncPromptShown = true;

  const modal = document.createElement("div");
  modal.id = "startupSyncPrompt";
  modal.className = "startup-sync-overlay";
  modal.innerHTML = `
    <div class="startup-sync-card">
      <h2>SYNC NOW</h2>
      <p>Fresh Firebase data load karna hai? Yes press karen to selected user data abhi sync ho jayega.</p>
      <div class="startup-sync-actions">
        <button type="button" class="startup-sync-yes" id="startupSyncYes">Yes</button>
        <button type="button" class="startup-sync-no" id="startupSyncNo">No</button>
      </div>
    </div>`;
  document.body.appendChild(modal);

  const close = () => modal.remove();
  document.getElementById("startupSyncNo")?.addEventListener("click", close);
  document.getElementById("startupSyncYes")?.addEventListener("click", () => {
    const yesBtn = document.getElementById("startupSyncYes");
    if (yesBtn) {
      yesBtn.textContent = "Syncing...";
      yesBtn.disabled = true;
    }
    const currentUser = (getLoggedUser() || "").toString().trim().toUpperCase();
    const selected = getSelectedSyncUserForPrompt();
    const finish = () => {
      renderInvoiceTable?.();
      renderMySaleTable?.();
      close();
    };
    if (currentUser === "ADMIN" && selected) {
      syncUserDataFromFirebase(() => {
        syncMySaleFromFirebase?.(finish, selected);
      }, selected);
    } else {
      syncUserDataFromFirebase(() => {
        syncMySaleFromFirebase?.(finish);
      });
    }
  });
}

function showLegalPage(page) {
  const modal = document.getElementById("legalModal");
  if (!modal) return;
  const sections = {
    terms: document.getElementById("legalTerms"),
    privacy: document.getElementById("legalPrivacy"),
    about: document.getElementById("legalAbout")
  };
  Object.values(sections).forEach(section => section?.classList.add("hidden"));
  const active = sections[page] || sections.terms;
  active?.classList.remove("hidden");
  const title = document.getElementById("legalModalTitle");
  if (title) {
    title.textContent = page === "privacy" ? "Privacy Policy" : page === "about" ? "About Me" : "Terms and Conditions";
  }
  modal.classList.remove("hidden");
}

function closeLegalPage() {
  document.getElementById("legalModal")?.classList.add("hidden");
}

function showAppNotification(message, type = "success") {
  let stack = document.getElementById("appToastStack");
  if (!stack) {
    stack = document.createElement("div");
    stack.id = "appToastStack";
    stack.className = "app-toast-stack";
    document.body.appendChild(stack);
  }
  const toast = document.createElement("div");
  toast.className = `app-toast ${type}`;
  toast.textContent = message;
  stack.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add("show"));
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 260);
  }, 3200);
}


/* ================================================================
   ✅ END Robust Sync System
================================================================ */

function calculateSmartPerformance() {
    let totalCustomerScore = 0;
    let customerCount = 0;

    Object.entries(customerTargets).forEach(([customerCode, customer]) => {
        const items = customer.items;
        const totalItems = Object.keys(items).length;
        if (totalItems === 0) return;

        let totalTargetQty = 0;
        let totalAchievedQty = 0;
        let completedItems = 0;

        Object.entries(items).forEach(([item, targetQty]) => {
            const achievedQty = invoices
                .filter(inv =>
                    inv.customerCode?.toUpperCase() === customerCode.toUpperCase() &&
                    inv.item?.toUpperCase() === item.toUpperCase()
                )
                .reduce((sum, inv) => sum + Number(inv.quantity || 0), 0);

            totalTargetQty += Number(targetQty);
            totalAchievedQty += achievedQty;

            if (achievedQty >= targetQty) completedItems++;
        });

        // --- Achieved% ---
        const achievedPercent = totalTargetQty > 0
            ? (totalAchievedQty / totalTargetQty) * 100
            : 0;

        // --- Item Completion Score ---
        const itemCompletionPercent = (completedItems / totalItems) * 100;

        // --- FINAL SMART SCORE (70% + 30%) ---
        const finalScore = (achievedPercent * 0.7) + (itemCompletionPercent * 0.3);

        totalCustomerScore += finalScore;
        customerCount++;
    });

    // --- RETURN OVERALL PERFORMANCE ---
    return customerCount > 0
        ? (totalCustomerScore / customerCount).toFixed(1)
        : 0;
}

function openCustomerPopup(customerCode) {

    const customer = customerTargets[customerCode];
    if (!customer) {
        alert("Customer not found!");
        return;
    }

    const ranked = getCustomerRankings();
    const rankInfo = ranked.find(c => c.code === customerCode);
    const customerLevel = rankInfo?.displayLevel || "";
    const levelColor = rankInfo?.levelColor || "#999";

    // ---------- KPI + TABLE ----------
    let rowsHtml = "";
    let totalItems = 0, nonProductive = 0, progress = 0, completed = 0;
    let totalTarget = 0, totalAchieved = 0, totalRemaining = 0, totalAchievedValue = 0;
    const zeroItems = [];

    const sortedItems = Object.keys(customer.items).sort();

    sortedItems.forEach(item => {
        const target = Number(customer.items[item]);

        const inv = invoices.filter(x =>
            x.customerCode?.toUpperCase() === customerCode &&
            x.item?.toUpperCase() === item
        );

        const achieved = inv.reduce((a, b) => a + Number(b.quantity || 0), 0);
        const achievedValue = inv.reduce((a, b) => a + (Number(b.quantity) * Number(b.rate)), 0);
        const capped = Math.min(achieved, target);
        const remaining = target - achieved;

        totalItems++;
        totalTarget += target;
        totalAchieved += capped;
        totalRemaining += Math.max(remaining, 0);
        totalAchievedValue += achievedValue;

        if (achieved === 0) {
            nonProductive++;
            zeroItems.push(item);
        }

        let rowStyle = "";
        if (remaining < 0) rowStyle = "background:#dc2626;color:white;";
        else if (achieved >= target) {
            completed++;
            rowStyle = "background:#16a34a;color:white;";
        } else if (achieved > 0) {
            progress++;
            const percent = Math.min((achieved / target) * 100, 100);
            rowStyle = `
                background: linear-gradient(to right, #16a34a ${percent}%, #60a5fa ${percent}%);
                color:white;
            `;
        }

        rowsHtml += `
            <tr style="${rowStyle}">
                <td class="border p-2">${item}</td>
                <td class="border p-2">${target}</td>
                <td class="border p-2">${achieved}</td>
                <td class="border p-2">${remaining}</td>
                <td class="border p-2 font-bold">${achievedValue.toLocaleString()}</td>
            </tr>
        `;
    });

    const overall = totalTarget > 0 ? ((totalAchieved / totalTarget) * 100).toFixed(1) : 0;

    // ---------- FINAL POPUP UI ----------
    const popup = `
    <div id="allocPopup" class="fixed inset-0 bg-black bg-opacity-50 flex justify-center items-center z-50">
        <div class="bg-white w-11/12 md:w-4/6 lg:w-1/2 rounded-2xl shadow-2xl p-6 max-h-[90vh] overflow-auto">

            <!-- HEADER -->
            <div class="mb-6 text-center p-6 rounded-2xl shadow-lg bg-gradient-to-r from-purple-700 via-purple-800 to-gray-900">

                <div class="flex justify-between items-center">
                    
                    <!-- LEFT: LEVEL BADGE -->
                    <p class="text-sm font-bold px-3 py-1 rounded-full text-black"
                       style="background:${levelColor}">
                        ${customerLevel}
                    </p>

                    <!-- CENTER TITLE -->
                    <h2 class="text-lg font-extrabold text-white drop-shadow-lg text-center flex-grow">
                        📊 Customer Dashboard
                    </h2>

                    <!-- RIGHT: NON-PRODUCTIVE BADGE -->
                    ${nonProductive > 0 ? `
                        <p class="text-sm font-bold px-3 py-1 rounded-full bg-red-600 text-white ml-3">
                            🚫 Non-Productive
                        </p>
                    ` : `<span></span>`}
                </div>

                <p class="text-3xl font-extrabold text-yellow-400 drop-shadow-lg mt-2">${customer.name}</p>
                <p class="text-gray-300 text-sm mt-1">${customer.city} • ${customerCode}</p>
            </div>

            <!-- KPIs -->
            <div class="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
                <div class="p-5 rounded-2xl shadow-lg text-center bg-blue-100">
                    <h3 class="text-lg font-bold text-blue-700">📦 Total Items</h3>
                    <p class="text-3xl font-extrabold text-blue-900 mt-2">${totalItems}</p>
                </div>

                <div class="p-5 rounded-2xl shadow-lg text-center bg-red-100">
                    <h3 class="text-lg font-bold text-red-700">🚫 Non-Buying</h3>
                    <p class="text-3xl font-extrabold text-red-900 mt-2">${nonProductive}</p>
                </div>

                <div class="p-5 rounded-2xl shadow-lg text-center bg-yellow-100">
                    <h3 class="text-lg font-bold text-yellow-700">⏳ Progress</h3>
                    <p class="text-3xl font-extrabold text-yellow-900 mt-2">${progress}</p>
                </div>

                <div class="p-5 rounded-2xl shadow-lg text-center bg-green-100">
                    <h3 class="text-lg font-bold text-green-700">✅ Completed</h3>
                    <p class="text-3xl font-extrabold text-green-900 mt-2">${completed}</p>
                </div>

                <div class="p-5 rounded-2xl shadow-lg text-center bg-purple-100">
                    <h3 class="text-lg font-bold text-purple-700">💰 Value</h3>
                    <p class="text-3xl font-extrabold text-purple-900 mt-2">${totalAchievedValue.toLocaleString()}</p>
                </div>
            </div>

            <!-- Progress Bar -->
            <div class="mb-6">
                <h3 class="font-semibold mb-2">📈 Overall Achievement</h3>
                <div class="w-full bg-gray-200 rounded-full h-6 overflow-hidden">
                    <div class="h-6 text-xs flex items-center justify-center font-bold text-white rounded-full"
                         style="width:${overall}%; background: linear-gradient(to right, #60a5fa, #16a34a);">
                        ${overall}%
                    </div>
                </div>
            </div>

            <!-- Breaking News -->
            <div class="relative overflow-hidden h-10 font-semibold text-sm rounded-lg shadow-lg mb-6
                        bg-gradient-to-r from-red-500 via-yellow-400 to-red-500 border border-red-600">

                ${
                    zeroItems.length > 0
                    ? `<marquee scrollamount="6">
                        ${zeroItems.map(it => `
                            <span class="text-white mx-4 bg-red-600 px-2 py-1 rounded-full">🚨 ${it}</span>
                        `).join("")}
                       </marquee>`
                    : `<span class="text-gray-100 flex items-center justify-center h-full">No Alerts</span>`
                }

            </div>

            <!-- Table -->
            <div class="overflow-auto max-h-80 border rounded">
                <table class="w-full text-sm border-collapse">
                    <thead class="bg-gray-200">
                        <tr>
                            <th class="border p-2">Item</th>
                            <th class="border p-2">Target</th>
                            <th class="border p-2">Achieved</th>
                            <th class="border p-2">Remaining</th>
                            <th class="border p-2">Value</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>

                    <tfoot>
                        <tr class="font-extrabold bg-indigo-800 text-white">
                            <td class="border p-2 text-center">Total</td>
                            <td class="border p-2 text-right">${totalTarget.toLocaleString()}</td>
                            <td class="border p-2 text-right">${totalAchieved.toLocaleString()}</td>
                            <td class="border p-2 text-right">${totalRemaining.toLocaleString()}</td>
                            <td class="border p-2 text-right">${totalAchievedValue.toLocaleString()}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>

            <div class="text-center mt-4">
                <button onclick="document.getElementById('allocPopup').remove()"
                        class="bg-red-600 text-white px-6 py-2 rounded-lg">
                    Close
                </button>
            </div>

        </div>
    </div>
    `;

    document.body.insertAdjacentHTML("beforeend", popup);
}


function searchCustomerFromMain() {
  const input = document.getElementById("mainCustomerSearch");
  const list = document.getElementById("mainCustomerSuggestions");
  const query = input.value.trim().toLowerCase();

  list.innerHTML = "";
  list.classList.add("hidden");

  if (!query) return;

  const matches = customers.filter(c =>
    c.name.toLowerCase().includes(query) ||
    c.code.toLowerCase().includes(query)
  );

  if (matches.length === 0) return;

  list.classList.remove("hidden");

  matches.forEach(c => {
    const div = document.createElement("div");
    div.className = "main-search-result";
    div.innerText = `${c.name} (${c.code}) - ${c.city}`;

    div.onclick = () => {
      input.value = "";
      list.innerHTML = "";
      list.classList.add("hidden");

      // 🔗 LINK TO ALLOCATION PAGE
      openCustomerFromMain(c.code);
      setTimeout(() => input.focus(), 0);
    };

    list.appendChild(div);
  });
}
function openCustomerFromMain(customerCode) {
  // Allocation page show
  showAllocationPage();

  // Allocation table render
  renderAllocationTables(customerCode);
}







