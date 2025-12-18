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
    const filtered = logged
        ? rows.filter(
              (r) =>
                  (r[6] || "").toString().trim().toUpperCase() === logged.toUpperCase() ||
                  (r[7] || "").toString().trim().toUpperCase() === logged.toUpperCase()
          )
        : rows;

    const mapped = filtered.map((row) => ({
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
    }));

    console.log("✅ Total CSV Rows:", lines.length);
    console.log("✅ Filtered Rows (after user filter):", filtered.length);
    console.log("✅ Final Mapped Rows:", mapped.length);

    allCSVData = mapped; // ✅ Save globally

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

    // ✅ My Sale Data
    let mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
    mapped.forEach((row) => {
        if (!row.SummaryNumber) return;
        const existing = mySaleData.find((s) => s.summary === row.SummaryNumber);
        if (existing) {
            existing.value = (Number(existing.value) || 0) + row.Value;
            existing.company = row.CompanyName || existing.company;
            existing.date = row.Date || existing.date;
        } else {
            mySaleData.push({
                summary: row.SummaryNumber,
                company: row.CompanyName,
                value: row.Value,
                date: row.Date
            });
        }
    });
    localStorage.setItem("mySaleData", JSON.stringify(mySaleData));

    // ✅ Render updates
    if (typeof renderInvoiceTable === "function") renderInvoiceTable(mapped);
    if (typeof renderMySaleTable === "function") renderMySaleTable();

   if (onDone) onDone(mapped);

// ✅ Create a unique hash from current data to detect duplicate uploads
const currentHash = btoa(JSON.stringify(mapped)).slice(0, 100);
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
  saveCSVToFirebase(mapped);

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
      uploadedBy: loggedUser,
      rows: data
    };

    const url = `${DATABASE_URL}/csvUploads/${loggedUser.toUpperCase()}/latest.json`;

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

function deleteUserData(userToDelete) {
    const loggedUser = getLoggedUser();

    if (!loggedUser) {
        alert("❌ No logged in user!");
        return;
    }

    // Only ADMIN can delete
    if (loggedUser.toUpperCase() !== "ADMIN") {
        alert("❌ Only ADMIN can delete data!");
        return;
    }

    if (!userToDelete) {
        alert("❌ Please select a user!");
        return;
    }

    if (!confirm(`Are you sure you want to DELETE all data of: ${userToDelete}?`)) {
        return;
    }

    // CORRECT PATH — EXACT PATH FROM YOUR FILE
    const finalPath = `${DATABASE_URL}/csvUploads/${userToDelete}/latest.json`;

    fetch(finalPath, { method: "DELETE" })
        .then(res => {
            if (!res.ok) throw new Error("Failed: " + res.status);
            alert(`✅ Deleted all data of ${userToDelete}`);
        })
        .catch(err => {
            console.error("❌ Delete error:", err);
            alert("Error deleting data!");
        });
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
        if (item && target > 0) {
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
                <div class="sticky top-0 z-20 bg-white border-b shadow-md">
                    <div class="p-2 bg-white">
                        <div class="flex flex-col sm:flex-row sm:flex-wrap items-start sm:items-center gap-2 sm:gap-4 mb-2">
                            <label class="font-bold">Filter by Status:</label>
                            <select id="statusFilter" class="w-full sm:w-auto border p-1 rounded text-sm">
                                <option value="all">🌍 All</option>
                                <option value="green">✅ Completed</option>
                                <option value="red">🔴 Red Zone</option>
                                <option value="normal">⏳ Pending</option>
                                <option value="nonProductive">🚫 Non Productive</option>
                                <option value="top10">🏆 Top 10 Customers</option>
                                <option value="itemSummary">📊 Item Summary</option>
                            </select>
                            <label class="font-bold">Filter by Item:</label>
                            <select id="itemFilter" class="w-full sm:w-auto border p-1 rounded text-sm">
                                <option value="all">📦 All Items</option>
                            </select>
                            <label class="font-bold">Filter by Rank:</label>
                            <select id="rankFilter" class="w-full sm:w-auto border p-1 rounded text-sm">
                                <option value="all">🏅 All Ranks</option>
                                <option value="Golden">🥇 Golden</option>
                                <option value="Silver">🥈 Silver</option>
                                <option value="Bronze">🥉 Bronze</option>
                                <option value="Level 1">Level 1</option>
                                <option value="Level 2">Level 2</option>
                                <option value="Level 3">Level 3</option>
                                <option value="Level 4">Level 4</option>
                                <option value="Level 5">Level 5</option>
                                <option value="Level 6">Level 6</option>
                                <option value="Level 7">Level 7</option>
                                <option value="Level 8">Level 8</option>
                                <option value="Level 9">Level 9</option>
                                <option value="Level 10">Level 10</option>
                                <option value="Level 15">Level 15</option>
                                 <option value="Level 20">Level 20</option>
                            </select>
                        </div>
                    </div>
                </div>
            </th>
        `;
        thead.prepend(filterRow);

        document.getElementById("statusFilter").addEventListener("change", renderInvoiceTable);
        document.getElementById("itemFilter").addEventListener("change", renderInvoiceTable);
        document.getElementById("rankFilter").addEventListener("change", renderInvoiceTable);
    }

    const selectedFilter = document.getElementById("statusFilter")?.value || "all";
    const selectedItem = document.getElementById("itemFilter")?.value || "all";
    const selectedRank = document.getElementById("rankFilter")?.value || "all";

    // --- Calculate customer ranks based on item count groups ---
    const allCustomers = Object.entries(customerTargets).map(([code, data]) => ({
        code,
        name: data.name || "Unknown",
        itemsCount: Object.keys(data.items || {}).length
    }));

    // Group customers by itemsCount
    const itemCountGroups = {};
    allCustomers.forEach(cust => {
        const count = cust.itemsCount;
        if (!itemCountGroups[count]) itemCountGroups[count] = [];
        itemCountGroups[count].push(cust);
    });

    // Sort groups by itemsCount in descending order
    const sortedGroups = Object.keys(itemCountGroups)
        .map(Number)
        .sort((a, b) => b - a)
        .map(count => itemCountGroups[count]);

    // Assign levels to groups
    sortedGroups.forEach((group, index) => {
        let level;
        if (index === 0) {
            level = "Golden";
        } else if (index === 1) {
            level = "Silver";
        } else if (index === 2) {
            level = "Bronze";
        } else {
            level = `Level ${index - 2}`;
        }
        group.forEach(cust => {
            cust.level = level;
        });
    });

    // Flatten back to allCustomers for filtering
    const rankedCustomers = sortedGroups.flat();

    // --- Debug: Log customer ranks ---
    console.log('Customer Ranks:', rankedCustomers.map(c => ({
        code: c.code,
        name: c.name,
        itemsCount: c.itemsCount,
        level: c.level
    })));

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
            if (!itemSummary[item]) itemSummary[item] = { totalTargetQty:0, totalAchievedQty:0, totalRemainingQty:0, totalValue:0, customerCount:0 };

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
        });
    });

    if (selectedFilter === "itemSummary" || selectedFilter === "nonProductiveItemSummary") {

        // --- Render Item Summary Table (QTY-based) ---
        Object.entries(itemSummary).forEach(([item, data]) => {
           // 🚫 Non Productive Item Summary ONLY
if (
    selectedFilter === "nonProductiveItemSummary" &&
    data.totalTargetQty > 0 &&
    data.totalAchievedQty > 0
) return;

            if (selectedItem !== "all" && selectedItem !== item) return;
            const perc = data.totalTargetQty>0?((data.totalAchievedQty/data.totalTargetQty)*100).toFixed(1):0;
            let rowClass = "bg-gray-50";
            if(data.totalRemainingQty<0) rowClass="bg-red-500 text-white";
            else if(data.totalRemainingQty===0 && data.totalAchievedQty>0) rowClass="bg-green-500 text-white";

            rowsHtml += `<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
                <td class="border p-1 sm:p-2"></td>
                <td class="border p-1 sm:p-2"></td>
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
                if(selectedItem!=="all" && selectedItem!==item) return;

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
                if(remainingQty<0){ rowClass="bg-red-500 text-white"; statusType="red"; }
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

                visibleItems.add(item);

                rowsHtml+=`<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
                    <td class="border p-1 sm:p-2">${customer.city||''}</td>
                    <td class="border p-1 sm:p-2">${customerCode}</td>
                    <td class="border p-1 sm:p-2">${customer.name||''} (${rankInfo.level})</td>
                    <td class="border p-1 sm:p-2">${item}</td>
                    <td class="border p-1 sm:p-2">${targetQtyNum.toLocaleString()}</td>
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
    }

    // --- Update Item Filter ---
    const itemFilter = document.getElementById("itemFilter");
    if(itemFilter){
        const currentValue = itemFilter.value;
// If Non Productive filter → no items should appear  
if (selectedFilter === "nonProductive") {
    itemFilter.innerHTML = `<option value="all">📦 All Items</option>`;
    itemFilter.value = "all";
    return;
}

        const sortedItems = Array.from(visibleItems).sort();
        itemFilter.innerHTML = `<option value="all">📦 All Items</option>` + sortedItems.map(it=>`<option value="${it}">${it}</option>`).join("");
        if([...sortedItems,"all"].includes(currentValue)) itemFilter.value=currentValue;
    }

    // --- Breaking News ---
    const breakingNews = document.getElementById("breakingNews");
    if(breakingNews){
        if(zeroAchieveCustomers.length>0){
            const colors=["text-red-600","text-blue-600","text-green-600","text-purple-600","text-pink-600","text-orange-600"];
            let index=0;
            const newsHtml = zeroAchieveCustomers.map(c=>{
                const color = colors[index%colors.length];
                index++;
                return `<span class="${color} mx-2 sm:mx-4 cursor-pointer underline" 
             onclick="openCustomerPopup('${c.code}')">
             🚨 ${c.name} [${c.code}]
        </span>`;

            }).join("");
            breakingNews.innerHTML=`<marquee behavior="scroll" direction="left" scrollamount="4">${newsHtml}</marquee>`;
        }else breakingNews.innerHTML='<span class="text-gray-600 flex items-center justify-center h-full">No alerts at this time</span>';
    }
}



// --- Popup function ---
function showFilteredPopup() {
    const selectedStatus = document.getElementById("statusFilter").value;
    const selectedItemValue = document.getElementById("itemFilter").value;

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

    let popupThead = ""; // dynamic header

   if (
    selectedStatus === "itemSummary" ||
    selectedStatus === "nonProductiveItemSummary"
) {

        // --- Item-based summary for popup ---
        let itemSummary = {};
        Object.entries(customerTargets).forEach(([customerCode, customer]) => {
            Object.entries(customer.items).forEach(([item, target]) => {
                if (!itemSummary[item]) {
                    itemSummary[item] = { totalTarget: 0, totalAchieved: 0, totalRemaining: 0, totalValue: 0, customerCount: 0 };
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
            });
        });

        popupThead = `
            <thead class="bg-gray-100 sticky top-0 z-40">
                <tr>
                    <th class="border p-2">Item</th>
                    <th class="border p-2">Customers</th>
                    <th class="border p-2">Target</th>
                    <th class="border p-2">Achieved</th>
                    <th class="border p-2">Remaining</th>
                    <th class="border p-2">%</th>
                    <th class="border p-2">Value</th>
                </tr>
            </thead>
        `;

        Object.entries(itemSummary).forEach(([item, data]) => {
           // 🚫 Non Productive Item Summary (Popup)
if (
    selectedStatus === "nonProductiveItemSummary" &&
    data.totalTarget > 0 &&
    data.totalAchieved > 0
) return;

            if (selectedItemValue !== "all" && selectedItemValue !== item) return;

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
                <td class="border p-1 sm:p-2">${data.totalTarget}</td>
                <td class="border p-1 sm:p-2">${data.totalAchieved}</td>
                <td class="border p-1 sm:p-2">${data.totalRemaining}</td>
                <td class="border p-1 sm:p-2 font-bold">${percentage}%</td>
                <td class="border p-1 sm:p-2 font-bold">${data.totalValue.toLocaleString()}</td>
            </tr>`;

            totalItems++;
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
                if (selectedItemValue !== "all" && selectedItemValue !== item) return;

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

                if (remaining < 0) {
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

popupRows += `<tr class="${rowClass} hover:bg-indigo-100 transition text-xs sm:text-sm">
    <td class="border p-1 sm:p-2">${customer.city || ''}</td>
    <td class="border p-1 sm:p-2">${customerCode}</td>
    <td class="border p-1 sm:p-2">${customer.name || ''}</td>
    <td class="border p-1 sm:p-2">${item}</td>
    <td class="border p-1 sm:p-2">${target}</td>
    <td class="border p-1 sm:p-2">${achieved}</td>
    <td class="border p-1 sm:p-2">${remaining}</td>
    <td class="border p-1 sm:p-2 font-bold">${remaining <= 0 ? "100%" : ((achieved/target*100).toFixed(1)+"%")}</td>
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
   const summaryRow = `
    <tr class="bg-indigo-100 font-bold text-xs sm:text-sm">
        <td colspan="2" class="border p-2 text-center">TOTAL (${totalCustomers} Customers)</td>
        <td class="border p-2">${totalItems} Items</td>
        <td class="border p-2">${totalTarget}</td>
        <td class="border p-2">${totalAchieved}</td>
        <td class="border p-2">${totalRemaining}</td>
        <td class="border p-2">  ${calculateSmartPerformance()}% </td> 
        <td class="border p-2">${totalValue.toLocaleString()}</td>
    </tr>`;

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

    // ✅ Calculate total items for each customer (for ranking)
    const allCustomers = Object.entries(customerTargets).map(([code, data]) => ({
        code,
        name: data.name || "Unknown",
        itemsCount: Object.keys(data.items || {}).length
    }));

    // ✅ Group customers by itemsCount
    const itemCountGroups = {};
    allCustomers.forEach(cust => {
        const count = cust.itemsCount;
        if (!itemCountGroups[count]) itemCountGroups[count] = [];
        itemCountGroups[count].push(cust);
    });

    // ✅ Sort groups by itemsCount in descending order
    const sortedGroups = Object.keys(itemCountGroups)
        .map(Number)
        .sort((a, b) => b - a)
        .map(count => itemCountGroups[count]);

    // ✅ Assign rank levels to groups
    sortedGroups.forEach((group, index) => {
        let level, levelColor;
        if (index === 0) {
            level = "🥇 Golden";
            levelColor = "#FFD700";
        } else if (index === 1) {
            level = "🥈 Silver";
            levelColor = "#C0C0C0";
        } else if (index === 2) {
            level = "🥉 Bronze";
            levelColor = "#CD7F32";
        } else {
            level = `Level ${index - 2}`;
            const shades = ["#E0FFFF", "#B0E0E6", "#ADD8E6", "#87CEEB", "#6495ED", "#4169E1", "#0000CD"];
            levelColor = shades[(index - 3) % shades.length];
        }
        group.forEach(cust => {
            cust.level = level;
            cust.levelColor = levelColor;
        });
    });

    // ✅ Flatten back to allCustomers for finding rank info
    const rankedCustomers = sortedGroups.flat();

    // ✅ Debug: Log customer ranks
    console.log('Customer Ranks:', rankedCustomers.map(c => ({
        code: c.code,
        name: c.name,
        itemsCount: c.itemsCount,
        level: c.level
    })));

    // ✅ Find current customer's rank info
    const rankInfo = rankedCustomers.find(c => c.code === customerCode);
    const customerLevel = rankInfo ? rankInfo.level : "";
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

// --------------- My Sale: clean implementation ---------------

// optional: keep this if you use toggleSections elsewhere
function toggleSections() {
  document.getElementById("formGrid")?.classList.toggle("hidden");
  document.getElementById("customContent")?.classList.toggle("hidden");
}

// load existing data (single declaration)
let mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");

// show / hide pages and mark active nav
function showMySalePage() {
  // hide all other pages
  const pages = ["mainPage", "allocationPage", "doneTargetPage", "bonusPage"];
  pages.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.add("hidden");
  });

  // show only My Sale Page
  const salePage = document.getElementById("mySalePage");
  if (salePage) salePage.classList.remove("hidden");

  // reset nav active states
  ["navInvoiceEntry","navAllocation","navDoneTargets","navBonus","navMySale","navMysale"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("bg-primary","text-white","bg-yellow-600");
  });

  // highlight nav for My Sale
  const nav = document.getElementById("navMySale") || document.getElementById("navMysale");
  if (nav) nav.classList.add("bg-yellow-600","text-white");

  // render table
  if (typeof renderMySaleTable === "function") {
    renderMySaleTable();
  }
}


// render table + total
function renderMySaleTable() {
  // reload from localStorage before rendering (refresh feature)
  mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");

  const salePage = document.getElementById("mySalePage");
  if (!salePage) return;

  const tbody = salePage.querySelector("#mySaleTableBody");
  const totalEl = salePage.querySelector("#mySaleTotal");

  if (!tbody) return;

  if (!mySaleData || mySaleData.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" class="text-center p-2 text-gray-500">No data yet</td></tr>`;
    if (totalEl) totalEl.textContent = "0";
    return;
  }

  let rows = "";
  let total = 0;
  mySaleData.forEach(sale => {
    const v = Number(sale.value) || 0;
    total += v;
    rows += `<tr>
      <td class="border p-2">${escapeHtml(sale.summary)}</td>
      <td class="border p-2">${escapeHtml(sale.company)}</td>
      <td class="border p-2 text-right">${formatNumber(v)}</td>
      <td class="border p-2 text-center">${escapeHtml(sale.date)}</td>
    </tr>`;
  });

  tbody.innerHTML = rows;
  if (totalEl) totalEl.textContent = formatNumber(total);
}

// small helpers
function formatNumber(n){ return Number(n).toLocaleString(); }
function escapeHtml(s){ return (s===undefined || s===null) ? "" : String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function isValidDateString(s) { return !isNaN(Date.parse(s)); }
function pickLatestDate(a,b){
  if (!a) return b || a;
  if (!b) return a;
  const da = Date.parse(a);
  const db = Date.parse(b);
  if (!isNaN(da) && !isNaN(db)) {
    return new Date(Math.max(da, db)).toISOString().split('T')[0];
  }
  return b.length >= a.length ? b : a;
}

// process parsed CSV rows (merge by summary)
function processSaleCsvRows(rows) {
  if (!rows || rows.length === 0) return;

  let added = 0, merged = 0;
  rows.forEach(row => {
    if (!row || row.length === 0) return;

    const summary = (row[10] || "").toString().trim();
    const company = (row[11] || "").toString().trim();
    const valueRaw = (row[12] || "").toString().trim();
    const date = (row[13] || "").toString().trim();

    if (!summary) return;

    const value = parseFloat(valueRaw.replace(/,/g, ''));
    if (isNaN(value)) return;

    const existing = mySaleData.find(s => String(s.summary) === summary);
    if (existing) {
      existing.value = Number(existing.value || 0) + value;
      existing.company = company || existing.company;
      existing.date = pickLatestDate(existing.date, date);
      merged++;
    } else {
      mySaleData.push({ summary, company, value: Number(value), date });
      added++;
    }
  });

  localStorage.setItem("mySaleData", JSON.stringify(mySaleData));
  renderMySaleTable();
  console.log(`MySale: added ${added}, merged ${merged}`);
}

// CSV change handler
function handleSaleCsvFileChange(e) {
  const file = e?.target?.files?.[0];
  if (!file) return;
  Papa.parse(file, {
    skipEmptyLines: true,
    complete: function(results) {
      if (results && results.data) {
        processSaleCsvRows(results.data);
      } else {
        console.warn("No rows parsed from CSV");
      }
    },
    error: function(err) {
      console.error("CSV parse error:", err);
      alert("CSV parse error: " + err.message);
    }
  });
  e.target.value = "";
}

// reset function with password check
function resetMySale() {
  const password = prompt("🔑 Enter password to reset My Sale data:");
  if (password !== "985973") {
    alert("❌ Wrong password! Reset cancelled.");
    return;
  }
  if (!confirm("⚠️ Are you sure you want to reset My Sale data? This will remove all saved sales.")) {
    return;
  }
  mySaleData = [];
  localStorage.removeItem("mySaleData");
  renderMySaleTable();
  console.log("MySale: reset");
  alert("✅ My Sale data has been reset successfully!");
}

// Attach listeners after DOM ready
document.addEventListener("DOMContentLoaded", () => {
  const saleInput = document.getElementById("saleCsvFile");
  if (saleInput) {
    saleInput.removeEventListener("change", handleSaleCsvFileChange);
    saleInput.addEventListener("change", handleSaleCsvFileChange);
  }

  const nav = document.getElementById("navMySale") || document.getElementById("navMysale");
  if (nav) {
    nav.removeEventListener("click", showMySalePage);
    nav.addEventListener("click", showMySalePage);
  }

  // ✅ add refresh button listener
  const refreshBtn = document.getElementById("refreshMySale");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      renderMySaleTable();
      alert("🔄 My Sale data refreshed!");
    });
  }

  renderMySaleTable();
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
  .then(() => alert("✅ Data uploaded to Firebase!"))
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
      syncUserDataFromFirebase(() => {
        console.log('✅ Data synced after CSV upload');
      });
    });
  } else {
    alert('Please upload a valid CSV file.');
    event.target.value = '';
  }
});

// ✅ Manual Sync Button Handler
document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("syncBtn");
    if (btn) {
        btn.addEventListener("click", () => {
            btn.innerText = "⏳ Syncing...";
            btn.disabled = true;
            syncUserDataFromFirebase(() => {
                alert("✅ Data synced successfully!");
                btn.innerText = "🔄 Sync Data";
                btn.disabled = false;
                renderInvoiceTable();
                renderMySaleTable?.();
            });
        });
    }
});

// 🧩 Fix — Normalize & Map Columns before processing
const normalizedRows = rows.map(r => {
  const obj = {};
  for (let key in r) {
    const nk = key.trim().toLowerCase();
    obj[nk] = r[key];
  }

  return {
    City: obj['city'] || '',
    CustomerCode: obj['customercode'] || obj['code'] || '',
    Customer: obj['customer'] || obj['customername'] || '',
    Item1: obj['item1'] || obj['item'] || '',
    Target1: Number(obj['target1'] || obj['target'] || 0),
    Achieve1: Number(
      obj['achieve1'] ??
      obj['achieve'] ??
      obj['achieved'] ??
      obj['achievedvalue'] ??
      obj['achv'] ??
      0
    ),
    User1: obj['user1'] || '',
    User2: obj['user2'] || '',
    DealQty: Number(obj['dealqty'] || 0),
    DealBonus: Number(obj['dealbonus'] || 0),
    SummaryNumber: obj['summarynumber'] || '',
    CompanyName: obj['companyname'] || '',
    Value: Number(
      (obj['value'] ??
       obj['val'] ??
       obj['achievedvalue'] ??
       '0').toString().replace(/,/g, '')
    ),
    Date: obj['date'] || '',
    ItemRate: Number((obj['itemrate'] || '0').toString().replace(/,/g, ''))
  };
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

    const path = `csvUploads/${loggedUser.toUpperCase()}/latest`;
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
      uploadedBy: loggedUser,
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

    const path = `csvUploads/${loggedUser.toUpperCase()}/latest`;
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

    // 2) Build a lookup from existing by key (CustomerCode|Item1)
    const lookup = {};
    existingRows.forEach(r => {
      const key = ((r.CustomerCode || "") + "|" + (r.Item1 || "")).trim().toUpperCase();
      lookup[key] = r;
    });

    // 3) Merge: for each new row update Achieve/Value/Deal fields but keep Target1 if new target is 0
    const merged = Object.assign({}, lookup); // key -> row
    data.forEach(newRow => {
      const code = (newRow.CustomerCode || "").trim().toUpperCase();
      const item = (newRow.Item1 || "").trim().toUpperCase();
      const key = (code + "|" + item).trim().toUpperCase();
      const cleanNew = {
        City: newRow.City || "",
        CustomerCode: (newRow.CustomerCode || "").trim().toUpperCase(),
        Customer: newRow.Customer || "",
        Item1: (newRow.Item1 || "").trim().toUpperCase(),
        Target1: parseSafeInt(newRow.Target1),
        Achieve1: parseSafeInt(newRow.Achieve1),
        User1: newRow.User1 || "",
        User2: newRow.User2 || "",
        DealQty: parseSafeInt(newRow.DealQty),
        DealBonus: parseSafeInt(newRow.DealBonus),
        SummaryNumber: newRow.SummaryNumber || "",
        CompanyName: newRow.CompanyName || "",
        Value: parseSafeFloat(newRow.Value),
        Date: newRow.Date || "",
        ItemRate: parseSafeFloat(newRow.ItemRate)
      };

      if (merged[key]) {
        // Keep existing target if new target is zero or missing
        const existing = merged[key];
        merged[key] = {
          ...existing,
          // fields to keep from existing if new is empty/zero
          Target1: cleanNew.Target1 > 0 ? cleanNew.Target1 : (parseSafeInt(existing.Target1) || 0),
          // update Achieve/Value and deal fields to new values (even if 0)
          Achieve1: cleanNew.Achieve1,
          Value: cleanNew.Value,
          DealQty: cleanNew.DealQty,
          DealBonus: cleanNew.DealBonus,
          Date: cleanNew.Date || existing.Date || "",
          ItemRate: cleanNew.ItemRate || existing.ItemRate || 0,
          // keep general meta fields from either
          City: cleanNew.City || existing.City,
          Customer: cleanNew.Customer || existing.Customer,
          User1: cleanNew.User1 || existing.User1,
          User2: cleanNew.User2 || existing.User2,
          SummaryNumber: cleanNew.SummaryNumber || existing.SummaryNumber,
          CompanyName: cleanNew.CompanyName || existing.CompanyName,
        };
      } else {
        // New entry
        merged[key] = cleanNew;
      }
    });

    // Convert merged lookup back to array
    const mergedArray = Object.values(merged);

    // 4) Upload into latest.json (overwrite safely)
    const payload = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: loggedUser,
      rows: mergedArray
    };

    const putRes = await fetch(url, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (putRes.ok) {
      console.log(`✅ saveCSVToFirebase: Uploaded ${mergedArray.length} rows to ${url}`);
      localStorage.setItem("excelData", JSON.stringify(mergedArray));
      // also keep lastCsvUploadRef for debugging
      localStorage.setItem("lastCsvUploadRef", url);
    } else {
      console.error("❌ saveCSVToFirebase: Upload failed:", putRes.status);
      // fallback: save locally
      localStorage.setItem("excelData", JSON.stringify(Object.values(merged)));
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

    // Recompute mySaleData
    let mySaleData = JSON.parse(localStorage.getItem("mySaleData") || "[]");
    excelData.forEach(row => {
      if (!row.SummaryNumber) return;
      const existing = mySaleData.find(s => s.summary === row.SummaryNumber);
      if (existing) {
        existing.value = (Number(existing.value) || 0) + Number(row.Value || 0);
        existing.company = row.CompanyName || existing.company;
        existing.date = row.Date || existing.date;
      } else {
        mySaleData.push({
          summary: row.SummaryNumber,
          company: row.CompanyName,
          value: row.Value,
          date: row.Date
        });
      }
    });
    localStorage.setItem("mySaleData", JSON.stringify(mySaleData));

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
      console.log(`🔐 ADMIN override: syncing data for user ${targetUser}`);
    } else {
      // default: sync for the logged-in user only
      targetUser = currentUser;
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
    syncUserDataFromFirebase(() => {
      if (btn) {
        alert('✅ Data synced for ' + selected.toUpperCase());
        btn.innerText = "🔄 Sync Data";
        btn.disabled = false;
        renderInvoiceTable();
        renderMySaleTable?.();
      }
    }, selected);
  } else {
    // normal sync (no override)
    syncUserDataFromFirebase(() => {
      if (btn) {
        alert('✅ Data synced successfully!');
        btn.innerText = "🔄 Sync Data";
        btn.disabled = false;
        renderInvoiceTable();
        renderMySaleTable?.();
      }
    });
  }
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

    // ---------- SAME RANK LOGIC AS ALLOCATION ----------
    const allCustomers = Object.entries(customerTargets).map(([code, data]) => ({
        code,
        name: data.name || "Unknown",
        itemsCount: Object.keys(data.items || {}).length
    }));

    const itemCountGroups = {};
    allCustomers.forEach(c => {
        if (!itemCountGroups[c.itemsCount]) itemCountGroups[c.itemsCount] = [];
        itemCountGroups[c.itemsCount].push(c);
    });

    const sortedGroups = Object.keys(itemCountGroups)
        .map(Number)
        .sort((a, b) => b - a)
        .map(count => itemCountGroups[count]);

    sortedGroups.forEach((group, index) => {
        let level, levelColor;

        if (index === 0) {
            level = "🥇 Golden"; levelColor = "#FFD700";
        } else if (index === 1) {
            level = "🥈 Silver"; levelColor = "#C0C0C0";
        } else if (index === 2) {
            level = "🥉 Bronze"; levelColor = "#CD7F32";
        } else {
            level = `Level ${index - 2}`;
            const shades = ["#E0FFFF", "#B0E0E6", "#ADD8E6", "#87CEEB", "#6495ED", "#4169E1", "#0000CD"];
            levelColor = shades[(index - 3) % shades.length];
        }

        group.forEach(c => {
            c.level = level;
            c.levelColor = levelColor;
        });
    });

    const ranked = sortedGroups.flat();
    const rankInfo = ranked.find(c => c.code === customerCode);

    const customerLevel = rankInfo?.level || "";
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
