let excelData = [];
let chart;

// 🔹 GitHub Pages URL for Excel
const EXCEL_URL = "https://raw.githubusercontent.com/yourusername/shipment-dashboard/main/shipment.xlsx";

// 🔹 Load Excel automatically on page load
window.onload = () => {
  fetch(EXCEL_URL)
    .then(res => res.arrayBuffer())
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      excelData = XLSX.utils.sheet_to_json(sheet);

      displayTable(excelData);
      updateChart(excelData);
    })
    .catch(err => console.error("Error loading Excel:", err));
};

// 🔹 Display data in table
function displayTable(data) {
  const tableHead = document.getElementById("tableHead");
  const tableBody = document.getElementById("tableBody");
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  if (data.length === 0) return;

  Object.keys(data[0]).forEach(key => {
    tableHead.innerHTML += `<th>${key}</th>`;
  });

  data.forEach(row => {
    let tr = "<tr>";
    Object.values(row).forEach(val => {
      tr += `<td>${val}</td>`;
    });
    tr += "</tr>";
    tableBody.innerHTML += tr;
  });
}

// 🔹 Apply filters
function applyFilter() {
  const product = document.getElementById('filterProduct').value.toLowerCase();
  const year = document.getElementById('filterYear').value;
  const month = document.getElementById('filterMonth').value;

  let filtered = excelData.filter(row => {
    let match = true;
    if (product && !String(row.Product).toLowerCase().includes(product)) match = false;
    if (year && new Date(row.Date).getFullYear() != year) match = false;
    if (month && new Date(row.Date).getMonth() + 1 != month) match = false;
    return match;
  });

  displayTable(filtered);
  updateChart(filtered);

  let totalQty = filtered.reduce((sum, row) => sum + Number(row.Quantity || 0), 0);
  alert(`📦 Total Quantity Shipped: ${totalQty}`);
}

// 🔹 Update chart
function updateChart(data) {
  if (!data || data.length === 0) return;

  let counts = {};
  data.forEach(row => {
    let date = new Date(row.Date);
    if (isNaN(date)) return;
    let key = `${date.getFullYear()}-${date.getMonth() + 1}`;
    counts[key] = (counts[key] || 0) + Number(row.Quantity || 0);
  });

  const labels = Object.keys(counts);
  const values = Object.values(counts);

  if (chart) chart.destroy();
  chart = new Chart(document.getElementById('salesChart'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Quantity Shipped',
        data: values
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: true } },
      scales: { y: { beginAtZero: true } }
    }
  });
}
