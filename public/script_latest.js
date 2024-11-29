const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<your-client-id>",
        authority: "https://login.microsoftonline.com/<your-tenant-id>",
        redirectUri: "http://localhost:8000",
    },
});

let allContacts = []; // To store all tenant contacts

// Login and Logout
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");

        // Fetch and display all contacts after login
        await fetchAndDisplayAllContacts();
        await populateCompanyDropdown();
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}

function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch and display all contacts
async function fetchAndDisplayAllContacts() {
    try {
        const response = await callGraphApi("/contacts?$select=displayName,mail,companyName");
        allContacts = response.value; // Store all contacts globally
        populateTable(allContacts); // Display all contacts
    } catch (error) {
        console.error("Error fetching contacts:", error);
        alert("Failed to fetch tenant contacts.");
    }
}

// Populate Company Dropdown
async function populateCompanyDropdown() {
    const uniqueCompanies = [...new Set(allContacts.map(contact => contact.companyName).filter(Boolean))];
    const dropdown = document.getElementById("companyFilter");
    dropdown.innerHTML = `<option value="">Filter by Company</option>`; // Reset dropdown
    uniqueCompanies.forEach(company => {
        const option = document.createElement("option");
        option.value = company;
        option.textContent = company;
        dropdown.appendChild(option);
    });

    // Add event listener to filter contacts by company
    dropdown.addEventListener("change", () => {
        const selectedCompany = dropdown.value;
        if (selectedCompany) {
            const filteredContacts = allContacts.filter(contact => contact.companyName === selectedCompany);
            populateTable(filteredContacts);
        } else {
            populateTable(allContacts); // Reset to all contacts if no company is selected
        }
    });
}

// Contact Search
function filterContacts() {
    const searchText = document.getElementById("contactSearch").value.toLowerCase();
    const filteredContacts = allContacts.filter(contact =>
        (contact.displayName && contact.displayName.toLowerCase().includes(searchText)) ||
        (contact.mail && contact.mail.toLowerCase().includes(searchText))
    );
    populateTable(filteredContacts);
}

// Populate Table
function populateTable(data) {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");
    outputHeader.innerHTML = "<th>Name</th><th>Email</th><th>Company</th>";
    outputBody.innerHTML = data.map(contact => `
        <tr>
            <td>${contact.displayName || "N/A"}</td>
            <td>${contact.mail || "N/A"}</td>
            <td>${contact.companyName || "N/A"}</td>
        </tr>
    `).join("");
}

function resetTable() {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");
    outputHeader.innerHTML = "";
    outputBody.innerHTML = "";
}

// Send Report as Mail
async function sendReportAsMail() {
    const email = document.getElementById("adminEmail").value;
    if (!email) return alert("Please provide an admin email.");

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `<table border="1"><thead><tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr></thead><tbody>${emailContent}</tbody></table>`;

    const message = {
        message: {
            subject: "Contact Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: email } }]
        }
    };
    try {
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}

// Download Report as CSV
function downloadReportAsCSV() {
    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );
    if (!rows.length) return alert("No data to download.");

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Contact_Report.csv";
    a.click();
}

// Call Graph API
async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    const tokenResponse = await msalInstance.acquireTokenSilent({ scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"], account });
    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers: { Authorization: `Bearer ${tokenResponse.accessToken}`, "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : null
    });

    // Correctly handle empty responses
    if (response.ok) {
        const contentType = response.headers.get("content-type");
        if (contentType && contentType.includes("application/json")) {
            return await response.json(); // Parse JSON response
        }
        return {}; // Return empty object for empty responses like 204 No Content
    } else {
        const errorText = await response.text();
        console.error("Graph API error response:", errorText);
        throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
    }
}
