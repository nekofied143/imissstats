let charts = [];

document.getElementById("uploadBtn").onclick = () => document.getElementById("fileInput").click();
document.getElementById("fileInput").addEventListener("change", handleFile);

Chart.defaults.maintainAspectRatio = false;

// Handle file import and load content dynamically
function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    // Destroy old charts
    charts.forEach(c => c.destroy());
    charts = [];

    // Show loading spinner
    document.getElementById("loading").classList.add("show");

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        processData(json);
    };

    reader.readAsArrayBuffer(file);
}

// Store global data for filtering
let GLOBAL_DATA = [];

// Process data and render charts
function processData(rows) {
    GLOBAL_DATA = rows.slice(1); // store clean data

    // Extract year from the first valid date in the data
    const year = extractYearFromData(GLOBAL_DATA);

    // Dynamically update the header title with the extracted year
    updateHeaderTitle(year);

    const result = buildDataset(GLOBAL_DATA);
    renderAll(result);
    setupFilter();

    // Hide loading spinner and show the dashboard
    document.getElementById("loading").classList.remove("show");
    document.getElementById("dashboard-container").style.display = "block";  // Show the dashboard
}

// Extract year from the first valid date in the data
function extractYearFromData(data) {
    for (let i = 0; i < data.length; i++) {
        const dateRaw = data[i][0];
        const date = parseDate(dateRaw);
        if (date) {
            return date.getFullYear();
        }
    }
    return new Date().getFullYear(); // Default to current year if no date is found
}

// Update the header title with the year
function updateHeaderTitle(year) {
    const headerTitle = document.querySelector("h1");
    headerTitle.textContent = `IMISS STATISTICS DASHBOARD FOR ${year}`;
}

// Build dataset for charts
function buildDataset(data, monthFilter = "ALL") {
    const staffMap = {};
    const stationMap = {};
    const concernMap = {};
    const monthMap = {};
    let totalTickets = 0;

    data.forEach(r => {
        const dateRaw = r[0];
        const concern = r[2];
        const station = r[3];
        const staff = r[6];

        const date = parseDate(dateRaw);
        if (!date) return;

        const month = date.toLocaleString("en-US", { month: "long" });

        if (monthFilter !== "ALL" && month !== monthFilter) return;

        totalTickets++;

        monthMap[month] = (monthMap[month] || 0) + 1;
        if (staff) staffMap[staff] = (staffMap[staff] || 0) + 1;
        if (station) stationMap[station] = (stationMap[station] || 0) + 1;
        if (concern) concernMap[concern] = (concernMap[concern] || 0) + 1;
    });

    return {
        totalTickets,
        months: mapToArray(monthMap, true),
        staff: mapToArray(staffMap),
        stations: mapToArray(stationMap),
        concerns: mapToArray(concernMap)
    };
}

// Parse dates from Excel file
function parseDate(value) {
    if (!value) return null;
    if (typeof value === "number") {
        const d = new Date(Math.round((value - 25569) * 86400 * 1000));
        return isNaN(d) ? null : d;
    }
    if (value instanceof Date) return value;
    const d = new Date(value);
    if (!isNaN(d)) return d;
    const parts = value.split("/");
    if (parts.length === 3) {
        const d2 = new Date(parts[2], parts[0] - 1, parts[1]);
        return isNaN(d2) ? null : d2;
    }
    return null;
}

// Convert map to array and sort
function mapToArray(map, isMonth = false) {
    let arr = Object.keys(map).map(k => ({ name: k, value: map[k] }));
    if (isMonth) {
        const order = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ];
        return arr.sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
    }
    return arr.sort((a, b) => b.value - a.value);
}

// Render all data and charts
function renderAll(data) {
    // Update KPIs
    document.getElementById("kpiTotal").innerText = data.totalTickets;
    document.getElementById("kpiTopStaff").innerText = data.staff[0]?.name || "-";
    document.getElementById("kpiTopStation").innerText = data.stations[0]?.name || "-";
    document.getElementById("kpiTopConcern").innerText = data.concerns[0]?.name || "-";

    // Render Monthly Call Logs
    renderMonths(data.months);

    // Render charts
    renderStaffChart(data.staff);
    renderPie("stationChart", data.stations);
    renderPie("concernChart", data.concerns);
}

// Setup month filter options
function setupFilter() {
    const select = document.getElementById("monthFilter");
    const allData = buildDataset(GLOBAL_DATA);

    select.innerHTML = `<option value="ALL">All Months</option>`;
    allData.months.forEach(m => {
        const opt = document.createElement("option");
        opt.value = m.name;
        opt.textContent = m.name;
        select.appendChild(opt);
    });

    select.onchange = () => {
        const filtered = buildDataset(GLOBAL_DATA, select.value);
        renderAll(filtered);
    };
}

// Render monthly data cards
function renderMonths(months) {
    const container = document.getElementById("gauges");
    container.innerHTML = "";
    months.forEach(m => {
        container.innerHTML += `
            <div class="gauge-card">
                <h4>${m.name}</h4>
                <div class="month-value">${m.value}</div>
            </div>
        `;
    });
}

// Render bar chart for top 3 staff
function renderStaffChart(data) {
    const existing = charts.findIndex(c => c.canvas.id === "staffChart");
    if (existing !== -1) {
        charts[existing].destroy();
        charts.splice(existing, 1);
    }

    const top3 = [...data].sort((a, b) => b.value - a.value).slice(0, 3);
    const chart = new Chart(document.getElementById("staffChart"), {
        type: "bar",
        data: {
            labels: top3.map(d => d.name),
            datasets: [{
                data: top3.map(d => d.value),
                backgroundColor: ["#6fa3ff", "#4a7bd1", "#2e5aac"]
            }]
        },
        options: {
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, ticks: { color: "#fff" } },
                x: { ticks: { color: "#fff" } }
            }
        }
    });

    charts.push(chart);
}

// Render pie charts for stations and concerns
function renderPie(id, data) {
    const existing = charts.findIndex(c => c.canvas.id === id);
    if (existing !== -1) {
        charts[existing].destroy();
        charts.splice(existing, 1);
    }

    const chartData = getTop10(data);

    const chart = new Chart(document.getElementById(id), {
        type: "pie",
        data: {
            labels: chartData.map(d => d.name),
            datasets: [{
                data: chartData.map(d => d.value),
                backgroundColor: [
                    "#4a7bd1", "#e74c3c", "#f1c40f", "#2ecc71", "#9b59b6",
                    "#1abc9c", "#e67e22", "#3498db", "#9b59b6", "#2ecc71",
                    "#95a5a6"
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: { color: "#fff" }
                }
            }
        }
    });

    charts.push(chart);
}

// Get top 10 items from the data
function getTop10(data) {
    const sorted = [...data].sort((a, b) => b.value - a.value);
    const top10 = sorted.slice(0, 10);
    const othersSum = sorted.slice(10).reduce((sum, item) => sum + item.value, 0);
    if (othersSum > 0) {
        top10.push({ name: "Others", value: othersSum });
    }
    return top10;
}