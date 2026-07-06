function formatIDR(number) {
  return new Intl.NumberFormat("id-ID", {
    style: "decimal",
    maximumFractionDigits: 0,
  }).format(number);
}

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const ws_data = [
    ["Customer", "Bulan", "Total Invoice", "Tonase"],
    ["Toko Berkah", "Januari", 3841418100, 4.5],
    ["Toko Berkah", "Februari", 3761107900, 3.2],
    ["Toko Berkah", "Maret", 4460308800, 6.8],
    ["Sinar Jaya", "Januari", 1500000000, 2.5],
    ["Sinar Jaya", "Februari", 1200000000, 8.1],
    ["Sinar Jaya", "Maret", 1800000000, 10.5],
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

// Mencari nama kolom di baris Excel secara fleksibel (tidak peduli huruf besar/kecil atau spasi ekstra)
function getFieldValue(row, targetName) {
  const target = targetName.trim().toLowerCase();
  const key = Object.keys(row).find((k) => k.trim().toLowerCase() === target);
  return key !== undefined ? row[key] : undefined;
}

// Parsing angka yang mendukung format Indonesia (koma sebagai desimal, titik sebagai ribuan)
// dan membuang satuan/teks tambahan seperti "ton" atau "Rp"
function parseNumberID(value) {
  if (typeof value === "number") return value;
  if (value === undefined || value === null) return 0;

  let str = String(value).trim();
  str = str.replace(/[^0-9,.\-]/g, ""); // buang semua kecuali angka, koma, titik, minus

  if (str.includes(",") && str.includes(".")) {
    // Format Indonesia: titik = ribuan, koma = desimal -> "1.234,5"
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    // Hanya koma -> anggap sebagai desimal -> "4,5"
    str = str.replace(",", ".");
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

// Menentukan persentase Cara Ke-1 berdasarkan total tonase pembelian customer
function tonaseToPct(tonase) {
  const t = parseFloat(tonase) || 0;
  if (t >= 10000) return 0.03;
  if (t >= 8000) return 0.025;
  if (t >= 6000) return 0.02;
  if (t >= 4000) return 0.015;
  if (t >= 2000) return 0.01;
  return 0;
}

// Ambang batas tonase minimum agar berhak atas Insentif Tax (dan turunannya:
// PPh 21, Total Transfer Tax, Total Transfer via Rek KTP). Satu konstanta ini
// dipakai di seluruh perhitungan supaya tidak ada selisih logika antar field.
const TONASE_ELIGIBLE_MIN = 2000;

let allData = [];
let currentMode = "detail";

function computeRow(row) {
  const customer = getFieldValue(row, "Customer") || "-";
  const bulan = getFieldValue(row, "Bulan") || "-";
  const totalInvoice = parseNumberID(getFieldValue(row, "Total Invoice"));
  const tonase = parseNumberID(getFieldValue(row, "Tonase"));

  const dppInvoice = Math.round(totalInvoice * 0.900900900922559264);
  const pctCara1 = tonaseToPct(tonase);
  const insentifReal1 = totalInvoice * pctCara1;

  const pctCara2 = dppInvoice !== 0 ? insentifReal1 / dppInvoice : 0;
  const insentifReal2 = dppInvoice * pctCara2;

  const pctTax = 0.01;

  // Customer berhak atas insentif tax hanya jika tonase bulan tsb sudah
  // mencapai TONASE_ELIGIBLE_MIN. Jika belum, Insentif Tax bulan itu = 0
  // (otomatis tidak ikut menambah akumulasi PPh 21 milik customer tsb).
  const eligibleTax = tonase >= TONASE_ELIGIBLE_MIN;

  const pctTaxDisplay = eligibleTax ? pctTax : 0;
  const insentifTax = eligibleTax ? dppInvoice * pctTax : 0;

  return {
    customer,
    bulan,
    totalInvoice,
    tonase,
    dppInvoice,
    pctCara1,
    insentifReal1,
    pctCara2,
    insentifReal2,
    pctTax: pctTaxDisplay,
    insentifTax,
    // pph21, totalTransferTax, dan totalTransferKtp dihitung secara kumulatif
    // per customer oleh computeCustomerTax() setelah semua baris diproses,
    // jadi nilainya sementara diisi 0 dulu di sini.
    pph21: 0,
    totalTransferTax: 0,
    totalTransferKtp: 0,
    _groupRowCount: 1,
    _isGroupFirst: true,
  };
}

// Mengurutkan data supaya baris-baris milik customer yang sama berdekatan
// (diperlukan agar sel PPh 21 / Total Transfer bisa digabung / rowspan di
// tampilan Detail). Urutan customer mengikuti kemunculan pertama di Excel,
// dan urutan bulan di dalam satu customer tetap sesuai urutan asli.
function sortByCustomerGroup(data) {
  const order = [];
  const seen = new Set();
  data.forEach((d) => {
    if (!seen.has(d.customer)) {
      seen.add(d.customer);
      order.push(d.customer);
    }
  });
  const orderIndex = new Map(order.map((c, i) => [c, i]));
  return [...data].sort(
    (a, b) => orderIndex.get(a.customer) - orderIndex.get(b.customer),
  );
}

// Menghitung PPh 21 secara KUMULATIF per customer:
// 1. Insentif Tax dari seluruh bulan milik satu customer dijumlahkan dulu.
// 2. DPP PPh 21 = 50% x total Insentif Tax kumulatif tsb.
// 3. Tarif progresif dikenakan ke DPP tsb (bertingkat, sesuai lapisan tarif).
// Hasil pph21 / totalTransferTax / totalTransferKtp akan SAMA untuk semua
// baris/bulan milik customer yang sama, dan ditandai (_isGroupFirst,
// _groupRowCount) supaya bisa digabung (merge) saat dirender.
function computeCustomerTax(data) {
  const groups = new Map();
  data.forEach((d, idx) => {
    if (!groups.has(d.customer)) groups.set(d.customer, []);
    groups.get(d.customer).push(idx);
  });

  groups.forEach((indices) => {
    const totalInsentifTax = indices.reduce(
      (sum, i) => sum + data[i].insentifTax,
      0,
    );
    const dppPPh21Group = 0.5 * totalInsentifTax;
    const pph21Group = Math.round(hitungPPh21Progresif(dppPPh21Group));
    const totalTransferTaxGroup = totalInsentifTax - pph21Group;
    const totalTransferKtpGroup = totalInsentifTax;

    indices.forEach((i, pos) => {
      data[i].pph21 = pph21Group;
      data[i].totalTransferTax = totalTransferTaxGroup;
      data[i].totalTransferKtp = totalTransferKtpGroup;
      data[i]._groupRowCount = indices.length;
      data[i]._isGroupFirst = pos === 0;
    });
  });
}

function prosesData(rows) {
  if (rows.length === 0) {
    alert("Data Excel kosong atau format tidak sesuai.");
    return;
  }

  const tonaseMissing = getFieldValue(rows[0], "Tonase") === undefined;
  if (tonaseMissing) {
    alert(
      'Kolom "Tonase" tidak ditemukan di file Excel Anda. Pastikan ada kolom dengan header persis "Tonase", jika tidak semua persentase Cara Ke-1 akan dihitung 0%.',
    );
  }

  let computed = rows.map(computeRow);
  computed = sortByCustomerGroup(computed);
  computeCustomerTax(computed);
  allData = computed;

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
    tfKtp_All = 0,
    tonase_All = 0;

  data.forEach((d) => {
    totalInv_All += d.totalInvoice;
    dppInv_All += d.dppInvoice;
    insReal1_All += d.insentifReal1;
    insReal2_All += d.insentifReal2;
    insTax_All += d.insentifTax;
    tonase_All += d.tonase;

    // pph21 / totalTransferTax / totalTransferKtp adalah nilai KUMULATIF per
    // customer (sama untuk setiap baris/bulan milik customer tsb), jadi hanya
    // dihitung sekali per customer (saat baris pertama grup) supaya tidak
    // dobel-hitung di total keseluruhan.
    if (d._isGroupFirst) {
      pph21_All += d.pph21;
      tfTax_All += d.totalTransferTax;
      tfKtp_All += d.totalTransferKtp;
    }

    // Sel PPh 21 / Total Transfer Tax / Total Transfer via Rek KTP digabung
    // (rowspan) sesuai jumlah bulan milik customer tsb. Hanya baris pertama
    // dari grup customer yang merender ketiga sel ini.
    let mergedCells = "";
    if (d._isGroupFirst) {
      mergedCells = `
                <td rowspan="${d._groupRowCount}">${formatIDR(d.pph21)}</td>
                <td rowspan="${d._groupRowCount}">${formatIDR(d.totalTransferTax)}</td>
                <td rowspan="${d._groupRowCount}">${formatIDR(d.totalTransferKtp)}</td>
            `;
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
                <td style="text-align:left;">${d.customer}</td>
                <td style="text-align:left;">${d.bulan}</td>
                <td>${formatIDR(d.totalInvoice)}</td>
                <td style="text-align:center;">${d.tonase.toLocaleString("id-ID")} ton</td>
                <td>${formatIDR(d.dppInvoice)}</td>
                <td style="text-align:center;">${(d.pctCara1 * 100).toFixed(2)}%</td>
                <td>${formatIDR(d.insentifReal1)}</td>
                <td style="text-align:center;">${(d.pctCara2 * 100).toFixed(2)}%</td>
                <td>${formatIDR(d.insentifReal2)}</td>
                <td style="text-align:center;">${(d.pctTax * 100).toFixed(2)}%</td>
                <td>${formatIDR(d.insentifTax)}</td>
                ${mergedCells}
            `;
    tbody.appendChild(tr);
  });

  const totalTr = document.createElement("tr");
  totalTr.className = "total-row";
  totalTr.innerHTML = `
            <td colspan="2" style="text-align:center;">TOTAL KESELURUHAN</td>
            <td>${formatIDR(totalInv_All)}</td>
            <td style="text-align:center;">${tonase_All.toLocaleString("id-ID")} ton</td>
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
    tbody.innerHTML = `<tr><td colspan="14" style="text-align:center; color:#888;">Tidak ada data yang cocok dengan filter customer.</td></tr>`;
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
        tonase: 0,
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
    g.tonase += d.tonase;
    g.dppInvoice += d.dppInvoice;
    g.insentifReal1 += d.insentifReal1;
    g.insentifReal2 += d.insentifReal2;
    g.insentifTax += d.insentifTax;
    // pph21 / totalTransferTax / totalTransferKtp sudah berupa nilai kumulatif
    // per customer (dihitung sekali di computeCustomerTax), jadi cukup
    // di-assign (bukan dijumlahkan per baris) supaya tidak dobel-hitung.
    g.pph21 = d.pph21;
    g.totalTransferTax = d.totalTransferTax;
    g.totalTransferKtp = d.totalTransferKtp;
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
    tfKtp_All = 0,
    tonase_All = 0;

  groups.forEach((g) => {
    totalInv_All += g.totalInvoice;
    dppInv_All += g.dppInvoice;
    insReal1_All += g.insentifReal1;
    insReal2_All += g.insentifReal2;
    insTax_All += g.insentifTax;
    pph21_All += g.pph21;
    tfTax_All += g.totalTransferTax;
    tfKtp_All += g.totalTransferKtp;
    tonase_All += g.tonase;

    const tr = document.createElement("tr");
    tr.innerHTML = `
                <td style="text-align:left;">${g.customer}</td>
                <td style="text-align:center;">${g.bulanCount}</td>
                <td>${formatIDR(g.totalInvoice)}</td>
                <td style="text-align:center;">${g.tonase.toLocaleString("id-ID")} ton</td>
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
            <td style="text-align:center;">${tonase_All.toLocaleString("id-ID")} ton</td>
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
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; color:#888;">Tidak ada data yang cocok dengan filter customer.</td></tr>`;
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