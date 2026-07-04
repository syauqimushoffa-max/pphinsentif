function formatIDR(number) {
  return new Intl.NumberFormat("id-ID", {
    style: "decimal",
    maximumFractionDigits: 0,
  }).format(number);
}

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const ws_data = [
    ["Customer", "Bulan", "Total Invoice"],
    ["Toko Berkah", "Januari", 3841418100],
    ["Toko Berkah", "Februari", 3761107900],
    ["Toko Berkah", "Maret", 4460308800],
    ["Sinar Jaya", "Januari", 1500000000],
    ["Sinar Jaya", "Februari", 1200000000],
    ["Sinar Jaya", "Maret", 1800000000],
  ];
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, "Template_Kalkulator_Insentif.xlsx");
}

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById("uploadStatus").innerText =
    "File Berhasil Dimuat: " + file.name;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Mengubah ke JSON array of objects
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    prosesData(jsonData);
  };
  reader.readAsArrayBuffer(file);
}

function hitungPPh21Progresif(dppPPh21) {
  let pph = 0;
  let sisa = dppPPh21;

  if (sisa > 0) {
    let porsi = Math.min(sisa, 60000000);
    pph += porsi * 0.05;
    sisa -= porsi;
  }
  if (sisa > 0) {
    let porsi = Math.min(sisa, 190000000);
    pph += porsi * 0.15;
    sisa -= porsi;
  }
  if (sisa > 0) {
    let porsi = Math.min(sisa, 250000000);
    pph += porsi * 0.25;
    sisa -= porsi;
  }
  if (sisa > 0) {
    let porsi = Math.min(sisa, 4500000000);
    pph += porsi * 0.3;
    sisa -= porsi;
  }
  if (sisa > 0) {
    pph += sisa * 0.35;
  }

  return pph;
}

let allData = [];
let currentMode = "detail";

function computeRow(row) {
  const customer = row["Customer"] || "-";
  const bulan = row["Bulan"] || "-";
  const totalInvoice = parseFloat(row["Total Invoice"]) || 0;

  const dppInvoice = Math.round(totalInvoice * 0.900900900922559264 );
  const pctCara1 = 0.04;
  const insentifReal1 = totalInvoice * pctCara1;

  const pctCara2 = dppInvoice !== 0 ? insentifReal1 / dppInvoice : 0;
  const insentifReal2 = dppInvoice * pctCara2;

  const pctTax = 0.01;
  const insentifTax = dppInvoice * pctTax;

  const dppPPh21 = 0.5 * insentifTax;
  const pph21 = Math.round(hitungPPh21Progresif(dppPPh21));

  const totalTransferTax = insentifTax - pph21;
  const totalTransferKtp = insentifReal2 - insentifTax;

  return {
    customer,
    bulan,
    totalInvoice,
    dppInvoice,
    pctCara1,
    insentifReal1,
    pctCara2,
    insentifReal2,
    pctTax,
    insentifTax,
    pph21,
    totalTransferTax,
    totalTransferKtp,
  };
}

function prosesData(rows) {
  if (rows.length === 0) {
    alert("Data Excel kosong atau format tidak sesuai.");
    return;
  }

  allData = rows.map(computeRow);

  document.getElementById("resultCard").style.display = "block";
  document.getElementById("filterCustomer").value = "";
  renderTable();
}

function getFilteredData() {
  const filterVal = document
    .getElementById("filterCustomer")
    .value.trim()
    .toLowerCase();
  if (!filterVal) return allData;
  return allData.filter((d) => d.customer.toLowerCase().includes(filterVal));
}

function setMode(mode) {
  currentMode = mode;
  document
    .getElementById("btnModeDetail")
    .classList.toggle("active", mode === "detail");
  document
    .getElementById("btnModeSummary")
    .classList.toggle("active", mode === "summary");
  document.getElementById("detailView").style.display =
    mode === "detail" ? "block" : "none";
  document.getElementById("summaryView").style.display =
    mode === "summary" ? "block" : "none";
  renderTable();
}

function renderTable() {
  if (currentMode === "detail") {
    renderDetail();
  } else {
    renderSummary();
  }
}

function renderDetail() {
  const tbody = document.getElementById("detailTableBody");
  tbody.innerHTML = "";
  const data = getFilteredData();

  let totalInv_All = 0,
    dppInv_All = 0,
    insReal1_All = 0,
    insReal2_All = 0,
    insTax_All = 0,
    pph21_All = 0,
    tfTax_All = 0,
    tfKtp_All = 0;

  data.forEach((d) => {
    totalInv_All += d.totalInvoice;
    dppInv_All += d.dppInvoice;
    insReal1_All += d.insentifReal1;
    insReal2_All += d.insentifReal2;
    insTax_All += d.insentifTax;
    pph21_All += d.pph21;
    tfTax_All += d.totalTransferTax;
    tfKtp_All += d.totalTransferKtp;

    const tr = document.createElement("tr");
    tr.innerHTML = `
                <td style="text-align:left;">${d.customer}</td>
                <td style="text-align:left;">${d.bulan}</td>
                <td>${formatIDR(d.totalInvoice)}</td>
                <td>${formatIDR(d.dppInvoice)}</td>
                <td style="text-align:center;">4.0%</td>
                <td>${formatIDR(d.insentifReal1)}</td>
                <td style="text-align:center;">${(d.pctCara2 * 100).toFixed(2)}%</td>
                <td>${formatIDR(d.insentifReal2)}</td>
                <td style="text-align:center;">1.0%</td>
                <td>${formatIDR(d.insentifTax)}</td>
                <td>${formatIDR(d.pph21)}</td>
                <td>${formatIDR(d.totalTransferTax)}</td>
                <td>${formatIDR(d.totalTransferKtp)}</td>
            `;
    tbody.appendChild(tr);
  });

  const totalTr = document.createElement("tr");
  totalTr.className = "total-row";
  totalTr.innerHTML = `
            <td colspan="2" style="text-align:center;">TOTAL KESELURUHAN</td>
            <td>${formatIDR(totalInv_All)}</td>
            <td>${formatIDR(dppInv_All)}</td>
            <td>-</td>
            <td class="highlight-green">${formatIDR(insReal1_All)}</td>
            <td>-</td>
            <td class="highlight-green">${formatIDR(insReal2_All)}</td>
            <td>-</td>
            <td>${formatIDR(insTax_All)}</td>
            <td class="highlight-yellow">${formatIDR(pph21_All)}</td>
            <td class="highlight-blue">${formatIDR(tfTax_All)}</td>
            <td class="highlight-orange">${formatIDR(tfKtp_All)}</td>
        `;
  tbody.appendChild(totalTr);

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="13" style="text-align:center; color:#888;">Tidak ada data yang cocok dengan filter customer.</td></tr>`;
  }
}

function groupByCustomer(data) {
  const map = new Map();
  data.forEach((d) => {
    if (!map.has(d.customer)) {
      map.set(d.customer, {
        customer: d.customer,
        bulanCount: 0,
        totalInvoice: 0,
        dppInvoice: 0,
        insentifReal1: 0,
        insentifReal2: 0,
        insentifTax: 0,
        pph21: 0,
        totalTransferTax: 0,
        totalTransferKtp: 0,
      });
    }
    const g = map.get(d.customer);
    g.bulanCount += 1;
    g.totalInvoice += d.totalInvoice;
    g.dppInvoice += d.dppInvoice;
    g.insentifReal1 += d.insentifReal1;
    g.insentifReal2 += d.insentifReal2;
    g.insentifTax += d.insentifTax;
    g.pph21 += d.pph21;
    g.totalTransferTax += d.totalTransferTax;
    g.totalTransferKtp += d.totalTransferKtp;
  });
  return Array.from(map.values());
}

function renderSummary() {
  const tbody = document.getElementById("summaryTableBody");
  tbody.innerHTML = "";
  const data = getFilteredData();
  const groups = groupByCustomer(data);

  let totalInv_All = 0,
    dppInv_All = 0,
    insReal1_All = 0,
    insReal2_All = 0,
    insTax_All = 0,
    pph21_All = 0,
    tfTax_All = 0,
    tfKtp_All = 0;

  groups.forEach((g) => {
    totalInv_All += g.totalInvoice;
    dppInv_All += g.dppInvoice;
    insReal1_All += g.insentifReal1;
    insReal2_All += g.insentifReal2;
    insTax_All += g.insentifTax;
    pph21_All += g.pph21;
    tfTax_All += g.totalTransferTax;
    tfKtp_All += g.totalTransferKtp;

    const tr = document.createElement("tr");
    tr.innerHTML = `
                <td style="text-align:left;">${g.customer}</td>
                <td style="text-align:center;">${g.bulanCount}</td>
                <td>${formatIDR(g.totalInvoice)}</td>
                <td>${formatIDR(g.dppInvoice)}</td>
                <td>${formatIDR(g.insentifReal1)}</td>
                <td>${formatIDR(g.insentifReal2)}</td>
                <td>${formatIDR(g.insentifTax)}</td>
                <td>${formatIDR(g.pph21)}</td>
                <td>${formatIDR(g.totalTransferTax)}</td>
                <td>${formatIDR(g.totalTransferKtp)}</td>
            `;
    tbody.appendChild(tr);
  });

  const totalTr = document.createElement("tr");
  totalTr.className = "total-row";
  totalTr.innerHTML = `
            <td style="text-align:center;">TOTAL KESELURUHAN</td>
            <td>-</td>
            <td>${formatIDR(totalInv_All)}</td>
            <td>${formatIDR(dppInv_All)}</td>
            <td class="highlight-green">${formatIDR(insReal1_All)}</td>
            <td class="highlight-green">${formatIDR(insReal2_All)}</td>
            <td>${formatIDR(insTax_All)}</td>
            <td class="highlight-yellow">${formatIDR(pph21_All)}</td>
            <td class="highlight-blue">${formatIDR(tfTax_All)}</td>
            <td class="highlight-orange">${formatIDR(tfKtp_All)}</td>
        `;
  tbody.appendChild(totalTr);

  if (groups.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; color:#888;">Tidak ada data yang cocok dengan filter customer.</td></tr>`;
  }
}

function exportToExcel() {
  if (allData.length === 0) {
    alert(
      "Belum ada data untuk diexport. Silakan unggah file Excel terlebih dahulu.",
    );
    return;
  }
  const tableId = currentMode === "detail" ? "detailTable" : "summaryTable";
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(document.getElementById(tableId));
  XLSX.utils.book_append_sheet(
    wb,
    ws,
    currentMode === "detail" ? "Detail" : "Summary",
  );

  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  XLSX.writeFile(
    wb,
    `Laporan_Insentif_${currentMode === "detail" ? "Detail" : "Summary"}_${dateStr}.xlsx`,
  );
}

function printReport() {
  if (allData.length === 0) {
    alert(
      "Belum ada data untuk dicetak. Silakan unggah file Excel terlebih dahulu.",
    );
    return;
  }
  const now = new Date();
  const dateStr = now.toLocaleDateString("id-ID", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  });
  document.getElementById("printDate").innerText = dateStr;
  document.getElementById("printMode").innerText =
    currentMode === "detail" ? "Detail" : "Summary (per Customer)";
  const filterVal = document.getElementById("filterCustomer").value.trim();
  document.getElementById("printFilter").innerText = filterVal
    ? filterVal
    : "Semua";

  window.print();
}
