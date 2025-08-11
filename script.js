// ================= Sidebar Active Link =================
const sections = document.querySelectorAll(".section");
const navLinks = document.querySelectorAll(".nav-links a");
window.addEventListener("scroll", () => {
  let current = "";
  sections.forEach((sec) => {
    const sectionTop = sec.offsetTop - 80;
    if (scrollY >= sectionTop) current = sec.getAttribute("id");
  });
  navLinks.forEach((link) => {
    link.classList.remove("active");
    if (link.getAttribute("href") === `#${current}`) {
      link.classList.add("active");
    }
  });
});

// ================= DOM References =================
const fileInput = document.getElementById("fileInput");
const searchInput = document.getElementById("searchInput");
const branchFilter = document.getElementById("branchFilter");
const yearFilter = document.getElementById("yearFilter");
const interestFilter = document.getElementById("interestFilter");
const exportBtn = document.getElementById("exportBtn");
const prepareEmailsBtn = document.getElementById("prepareEmailsBtn");
const generateReportBtn = document.getElementById("generateReportBtn"); // NEW
const tableBody = document.querySelector("#studentTable tbody");
const selectAll = document.getElementById("selectAll");
const mailtoLinks = document.getElementById("mailtoLinks");
const topInterestsList = document.getElementById("topInterestsList");
const keyInsights = document.getElementById("keyInsights");

let allData = [],
  filteredData = [];
let sortConfig = { key: null, asc: true };
let branchChartInstance, yearChartInstance;

// ================= Data Normalization =================
function normalizeRow(row) {
  const o = {};
  for (const k in row) {
    const key = k.trim().toLowerCase();
    const v = String(row[k] || "").trim();
    if (/(name|student)/.test(key)) o.name = v;
    else if (/(branch|dept|department)/.test(key)) o.branch = v;
    else if (/(year|batch)/.test(key)) o.year = v;
    else if (/email/.test(key)) o.email = v;
    else if (/(interest|hobby|skills)/.test(key)) o.interests = v;
  }
  return {
    name: o.name || "",
    branch: o.branch || "",
    year: o.year || "",
    email: o.email || "",
    interests: o.interests || "",
  };
}

// ================= Stats Calculation =================
function calculateStats() {
  const total = allData.length,
    filtered = filteredData.length;
  const branchCounts = {},
    yearCounts = {},
    interestCounts = {};

  filteredData.forEach((s) => {
    if (s.branch) branchCounts[s.branch] = (branchCounts[s.branch] || 0) + 1;
    if (s.year) yearCounts[s.year] = (yearCounts[s.year] || 0) + 1;
    if (s.interests) {
      s.interests
        .split(/[,;]/)
        .map((i) => i.trim())
        .filter(Boolean)
        .forEach((i) => {
          interestCounts[i] = (interestCounts[i] || 0) + 1;
        });
    }
  });
  return { total, filtered, branchCounts, yearCounts, interestCounts };
}
function updateOverallStats(stats) {
  statsCount.innerHTML = `
    <div>Total Students: <strong>${stats.total}</strong></div>
    <div>Displayed Students: <strong>${stats.filtered}</strong></div>
  `;
}
// ================= Chart Update =================
function updateCharts(stats) {
  if (branchChartInstance) branchChartInstance.destroy();
  branchChartInstance = new Chart(document.getElementById("branchChart"), {
    type: "pie",
    data: {
      labels: Object.keys(stats.branchCounts),
      datasets: [
        {
          data: Object.values(stats.branchCounts),
          backgroundColor: [
            "#42a5f5",
            "#66bb6a",
            "#ffca28",
            "#ef5350",
            "#ab47bc",
          ],
        },
      ],
    },
    options: { plugins: { legend: { position: "bottom" } } },
  });

  if (yearChartInstance) yearChartInstance.destroy();
  yearChartInstance = new Chart(document.getElementById("yearChart"), {
    type: "bar",
    data: {
      labels: Object.keys(stats.yearCounts),
      datasets: [
        {
          label: "Students",
          data: Object.values(stats.yearCounts),
          backgroundColor: "#42a5f5",
        },
      ],
    },
    options: { plugins: { legend: { display: false } } },
  });
}

// ================= Smart Insights Panel =================
function updateInsights(stats) {
  topInterestsList.innerHTML = "";
  const sortedInterests = Object.entries(stats.interestCounts).sort(
    (a, b) => b[1] - a[1]
  );
  sortedInterests.slice(0, 5).forEach(([i, c]) => {
    const li = document.createElement("li");
    li.className = "interest-badge";
    li.textContent = `${i} (${c})`;
    topInterestsList.appendChild(li);
  });

  let insight = "";
  if (Object.keys(stats.branchCounts).length) {
    const [b, bc] = Object.entries(stats.branchCounts).sort(
      (a, b) => b[1] - a[1]
    )[0];
    insight += `📊 Top Branch: ${b} (${((bc / stats.filtered) * 100).toFixed(
      1
    )}%).\n`;
  }
  if (sortedInterests.length) {
    insight += `🌟 Top Interest: ${sortedInterests[0][0]} (${sortedInterests[0][1]} students).\n`;
    insight += `🎯 Most Unique Interest: ${
      sortedInterests[sortedInterests.length - 1][0]
    } (${sortedInterests[sortedInterests.length - 1][1]} student). \n`;
  }
  const years = Object.keys(stats.yearCounts);
  if (years.length) {
    insight += `📅 Years Represented: ${years.join(", ")}.\n`;
  }
  keyInsights.textContent = insight;
}

// ================= Filtering =================
function applyFilters() {
  const sq = searchInput.value.toLowerCase(),
    bq = branchFilter.value,
    yq = yearFilter.value,
    iq = interestFilter.value.toLowerCase();
  filteredData = allData.filter((r) => {
    if (bq && r.branch !== bq) return false;
    if (yq && r.year !== yq) return false;
    if (iq && !r.interests.toLowerCase().includes(iq)) return false;
    if (sq)
      return `${r.name} ${r.branch} ${r.year} ${r.email} ${r.interests}`
        .toLowerCase()
        .includes(sq);
    return true;
  });
  applySort();
  renderTable();
  const stats = calculateStats();
  updateOverallStats(stats);
  updateCharts(stats);
  updateInsights(stats);
}

// ================= Sorting =================
function applySort() {
  if (!sortConfig.key) return;
  filteredData.sort((a, b) => {
    const va = (a[sortConfig.key] || "").toLowerCase(),
      vb = (b[sortConfig.key] || "").toLowerCase();
    return (va > vb ? 1 : -1) * (sortConfig.asc ? 1 : -1);
  });
}

// ================= Table Rendering =================
function renderTable() {
  tableBody.innerHTML = "";
  filteredData.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input type="checkbox" class="rowCheck" data-idx="${idx}"></td>
                    <td>${r.name}</td><td>${r.branch}</td><td>${r.year}</td><td>${r.email}</td><td>${r.interests}</td>`;
    tableBody.appendChild(tr);
  });
  selectAll.checked = false;
  mailtoLinks.innerHTML = "";
}

// ================= CSV Export =================
function exportCSV() {
  if (!filteredData.length) return alert("No data to export!");
  const headers = ["Name", "Branch", "Year", "Email", "Interests"];
  const rows = filteredData.map((r) => [
    r.name,
    r.branch,
    r.year,
    r.email,
    r.interests,
  ]);
  const csv = [headers, ...rows]
    .map((row) =>
      row.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "students.csv";
  a.click();
}

// ================= PDF Report Export =================
function generateReport() {
  if (!filteredData.length) return alert("No data to generate report!");

  const reportContent = document.getElementById("statistics");

  html2canvas(reportContent).then((canvas) => {
    const imgData = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf; // ✅ Correct way to access jsPDF in UMD build
    const pdf = new jsPDF("p", "mm", "a4");

    // Title & Timestamp
    pdf.setFontSize(16);
    pdf.text("Student Statistics Report", 10, 10);
    pdf.setFontSize(10);
    pdf.text(`Generated on: ${new Date().toLocaleString()}`, 10, 18);

    // Image of the Statistics Section
    const imgProps = pdf.getImageProperties(imgData);
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    pdf.addImage(imgData, "PNG", 10, 25, pdfWidth - 20, pdfHeight);

    pdf.save("student_report.pdf");
  });
}

// ================= Email Preparation =================
function prepareEmails() {
  const selected = Array.from(document.querySelectorAll(".rowCheck"))
    .filter((cb) => cb.checked)
    .map((cb) => filteredData[cb.dataset.idx]);
  if (!selected.length) return alert("Select at least one student!");
  mailtoLinks.innerHTML = "<h3>Click to open email client:</h3>";
  selected.forEach((s) => {
    if (!s.email) {
      mailtoLinks.innerHTML += `<div>No email for ${s.name}</div>`;
      return;
    }
    const subject = `Hello ${s.name}`;
    const body = `Hi ${s.name},\n\nWe noticed your interests: ${s.interests}.\nWe have opportunities for you!\n\nRegards,\nCSI Team`;
    mailtoLinks.innerHTML += `<a href="mailto:${encodeURIComponent(
      s.email
    )}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(
      body
    )}" target="_blank">${s.name} &lt;${s.email}&gt;</a>`;
  });
}

// ================= Filter Dropdowns =================
function populateFilters() {
  const branches = [
    ...new Set(allData.map((r) => r.branch).filter(Boolean)),
  ].sort();
  const years = [...new Set(allData.map((r) => r.year).filter(Boolean))].sort();
  branchFilter.innerHTML =
    '<option value="">All Branches</option>' +
    branches.map((b) => `<option value="${b}">${b}</option>`).join("");
  yearFilter.innerHTML =
    '<option value="">All Years</option>' +
    years.map((y) => `<option value="${y}">${y}</option>`).join("");
}

// ================= Event Listeners =================
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const wb = XLSX.read(evt.target.result, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    allData = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(normalizeRow);
    populateFilters();
    applyFilters();
  };
  reader.readAsArrayBuffer(file);
});
searchInput.addEventListener("input", applyFilters);
branchFilter.addEventListener("change", applyFilters);
yearFilter.addEventListener("change", applyFilters);
interestFilter.addEventListener("input", applyFilters);
exportBtn.addEventListener("click", exportCSV);
prepareEmailsBtn.addEventListener("click", prepareEmails);
generateReportBtn.addEventListener("click", generateReport);
selectAll.addEventListener("change", () =>
  document
    .querySelectorAll(".rowCheck")
    .forEach((cb) => (cb.checked = selectAll.checked))
);
document.querySelectorAll("#studentTable th[data-key]").forEach((th) =>
  th.addEventListener("click", () => {
    const key = th.dataset.key;
    if (sortConfig.key === key) sortConfig.asc = !sortConfig.asc;
    else {
      sortConfig.key = key;
      sortConfig.asc = true;
    }
    applySort();
    renderTable();
  })
);
