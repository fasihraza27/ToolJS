let excelData = [];
let processedData = [];
let t0Businesses = new Map();
let processType = "t1";
let txtExcelData = [];
let txtFiles = [];
let matchedTXTData = [];
let unmatchedData = [];

// CSV file name display
document
  .getElementById("fileInput")
  .addEventListener("change", function (event) {
    const files = event.target.files;
    const csvFileName = document.getElementById("csvFileName");
    if (files.length === 0) {
      csvFileName.textContent = "";
      return;
    }
    if (files.length === 1) {
      csvFileName.textContent = `Uploaded: ${files[0].name}`;
    } else {
      csvFileName.textContent = `Uploaded: ${files.length} CSV files`;
    }
  });

document
  .getElementById("folderInput")
  .addEventListener("change", function (event) {
    const files = event.target.files;

    if (files.length > 0) {
      document.getElementById("txtFileName").textContent =
        files.length === 1
          ? `Uploaded: ${files[0].name}`
          : `Uploaded: ${files.length} TXT files`;
    } else {
      document.getElementById("txtFileName").textContent = "";
    }
  });

document.getElementById("fileInput").addEventListener("change", readCSV);

function readCSV(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);

    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: true,
    });

    const sheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[sheetName];

    excelData = XLSX.utils.sheet_to_json(worksheet);

    document.getElementById("result").innerHTML =
      `${excelData.length} records loaded successfully`;

    // Reset previous processing
    processedData = [];

    document.getElementById("processStatus").style.display = "none";

    document.getElementById("matchingStatus").innerHTML =
      "Waiting for TXT file...";
  };

  reader.readAsArrayBuffer(file);
}

function calculateFee(amount) {
  amount = Number(amount);

  if (amount <= 5000) return 12.5;
  if (amount <= 10000) return 15.5;
  if (amount <= 100000) return 39;
  if (amount <= 250000) return 62.5;
  if (amount <= 1000000) return 125;
  if (amount <= 2500000) return 250;
  if (amount <= 5000000) return 375;

  return 500;
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function formatDate(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      return "";
    }

    return (
      value.getFullYear() +
      "-" +
      pad(value.getMonth() + 1) +
      "-" +
      pad(value.getDate()) +
      " " +
      pad(value.getHours()) +
      ":" +
      pad(value.getMinutes()) +
      ":" +
      pad(value.getSeconds())
    );
  }

  value = String(value).trim();

  if (value.includes("T")) {
    value = value.replace("T", " ");
  }

  // MM/DD/YYYY

  if (value.includes("/")) {
    const parts = value.split(/\s+/);

    const datePart = parts[0];

    let timePart = parts[1] || "00:00:00";

    let ampm = "";

    if (parts[2]) {
      ampm = parts[2].toUpperCase();
    }

    let [month, day, year] = datePart.split("/").map(Number);

    if (year < 100) {
      year += 2000;
    }

    const timeParts = timePart.split(":").map(Number);

    let hour = timeParts[0] || 0;

    let minute = timeParts[1] || 0;

    let second = timeParts[2] || 0;

    if (ampm === "PM" && hour < 12) {
      hour += 12;
    }

    if (ampm === "AM" && hour === 12) {
      hour = 0;
    }

    const date = new Date(year, month - 1, day, hour, minute, second);

    if (isNaN(date.getTime())) {
      return "";
    }

    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      " " +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes()) +
      ":" +
      pad(date.getSeconds())
    );
  }

  // YYYY-MM-DD HH:mm:ss

  let match = value.match(
    /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/,
  );

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6]);

    const date = new Date(year, month - 1, day, hour, minute, second);

    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      " " +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes()) +
      ":" +
      pad(date.getSeconds())
    );
  }

  // YYYY-MM-DD HH:mm

  match = value.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const hour = Number(match[4]);
    const minute = Number(match[5]);

    const date = new Date(year, month - 1, day, hour, minute, 0);

    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      " " +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes()) +
      ":00"
    );
  }

  // Excel serial date

  if (!isNaN(value)) {
    const excelDate = Number(value);

    let date;

    if (typeof XLSX !== "undefined" && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const parsed = XLSX.SSF.parse_date_code(excelDate);

      if (parsed) {
        date = new Date(
          parsed.y,
          parsed.m - 1,
          parsed.d,
          parsed.H || 0,
          parsed.M || 0,
          Math.floor(parsed.S || 0),
        );
      }
    }

    if (!date) {
      const excelEpoch = new Date(1899, 11, 30);

      date = new Date(excelEpoch.getTime() + excelDate * 86400000);
    }

    if (isNaN(date.getTime())) {
      return "";
    }

    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      " " +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes()) +
      ":" +
      pad(date.getSeconds())
    );
  }

  return value;
}

function formatDateOnly(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      return "";
    }

    return (
      value.getFullYear() +
      "-" +
      pad(value.getMonth() + 1) +
      "-" +
      pad(value.getDate())
    );
  }

  value = String(value).trim();

  if (value.includes("/")) {
    const datePart = value.split(/\s+/)[0];

    let [month, day, year] = datePart.split("/").map(Number);

    if (year < 100) {
      year += 2000;
    }

    return year + "-" + pad(month) + "-" + pad(day);
  }

  if (!isNaN(value)) {
    const excelDate = Number(value);

    let date;

    if (typeof XLSX !== "undefined" && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const parsed = XLSX.SSF.parse_date_code(excelDate);

      if (parsed) {
        date = new Date(parsed.y, parsed.m - 1, parsed.d);
      }
    }

    if (!date) {
      const excelEpoch = new Date(1899, 11, 30);

      date = new Date(excelEpoch.getTime() + excelDate * 86400000);
    }

    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate())
    );
  }

  return value;
}

function Transactiondatetime(value) {
  if (value === null || value === undefined || value === "") {
    return NaN;
  }

  if (value instanceof Date) {
    return value.getTime();
  }

  value = String(value).trim();

  if (value.includes("T")) {
    value = value.replace("T", " ");

    if (value.split(":").length === 2) {
      value += ":00";
    }
  }

  // YYYY-MM-DD HH:mm:ss

  let match = value.match(
    /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/,
  );

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6]);

    return new Date(year, month - 1, day, hour, minute, second).getTime();
  }

  // YYYY-MM-DD HH:mm

  match = value.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const hour = Number(match[4]);
    const minute = Number(match[5]);

    return new Date(year, month - 1, day, hour, minute, 0).getTime();
  }

  // MM/DD/YYYY

  if (value.includes("/")) {
    const parts = value.split(/\s+/);

    const datePart = parts[0];

    const timePart = parts[1] || "00:00:00";

    let ampm = "";

    if (parts[2]) {
      ampm = parts[2].toUpperCase();
    }

    let [month, day, year] = datePart.split("/").map(Number);

    if (year < 100) {
      year += 2000;
    }

    const timeParts = timePart.split(":").map(Number);

    let hour = timeParts[0] || 0;

    const minute = timeParts[1] || 0;

    const second = timeParts[2] || 0;

    if (ampm === "PM" && hour < 12) {
      hour += 12;
    }

    if (ampm === "AM" && hour === 12) {
      hour = 0;
    }

    const date = new Date(year, month - 1, day, hour, minute, second);

    if (isNaN(date.getTime())) {
      return NaN;
    }

    return date.getTime();
  }

  // Excel serial date

  if (!isNaN(value)) {
    const excelDate = Number(value);

    if (typeof XLSX !== "undefined" && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const parsed = XLSX.SSF.parse_date_code(excelDate);

      if (parsed) {
        const date = new Date(
          parsed.y,
          parsed.m - 1,
          parsed.d,
          parsed.H || 0,
          parsed.M || 0,
          Math.floor(parsed.S || 0),
        );

        return date.getTime();
      }
    }

    const excelEpoch = new Date(1899, 11, 30);

    const date = new Date(excelEpoch.getTime() + excelDate * 86400000);

    return date.getTime();
  }

  return NaN;
}

function selectTime() {
  const fromInput = document.getElementById("fromDateTime");

  const toInput = document.getElementById("toDateTime");

  const fromValue = fromInput.value.trim();

  const toValue = toInput.value.trim();

  if (!fromValue && !toValue) {
    return excelData;
  }

  let fromTime = -Infinity;

  let toTime = Infinity;

  if (fromValue) {
    fromTime = Transactiondatetime(fromValue);

    if (isNaN(fromTime)) {
      alert("Invalid From date/time.");

      return null;
    }
  }

  if (toValue) {
    toTime = Transactiondatetime(toValue);

    if (isNaN(toTime)) {
      alert("Invalid To date/time.");

      return null;
    }

    // Include entire selected minute

    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(toValue)) {
      toTime += 60 * 1000 - 1;
    }
  }

  if (fromTime > toTime) {
    alert("From date/time cannot be greater than To date/time.");

    return null;
  }

  return excelData.filter((row) => {
    const value = row["Transaction Date Time"];

    if (value === null || value === undefined || value === "") {
      return false;
    }

    const transactionTime = Transactiondatetime(value);

    if (isNaN(transactionTime)) {
      console.log("Invalid Transaction Date Time:", value);

      return false;
    }

    return transactionTime >= fromTime && transactionTime <= toTime;
  });
}

async function processT0() {
  processType = "t0";

  if (excelData.length === 0) {
    alert("Please upload a CSV file first.");

    return;
  }

  const filteredData = selectTime();

  if (filteredData === null) {
    return;
  }

  if (filteredData.length === 0) {
    alert("No records found in the selected Transaction Date Time range.");

    return;
  }

  const loaded = await loadT0Businesses();

  if (!loaded) {
    return;
  }

  processedData = filteredData
    .filter((row) => {
      return t0Businesses.has(String(row["Foree ID"]));
    })

    .map((row) => {
      let newRow = {
        ...row,
      };

      if (newRow["Bill Creation Date"]) {
        newRow["Bill Creation Date"] = formatDateOnly(
          newRow["Bill Creation Date"],
        );
      }

      if (newRow["Bill Due Date"]) {
        newRow["Bill Due Date"] = formatDateOnly(newRow["Bill Due Date"]);
      }

      if (newRow["Transaction Date Time"]) {
        newRow["Transaction Date Time"] = formatDate(
          newRow["Transaction Date Time"],
        );
      }

      if (newRow["Initiator Settlement Date"]) {
        newRow["Initiator Settlement Date"] = formatDateOnly(
          newRow["Initiator Settlement Date"],
        );
      }

      const business = t0Businesses.get(String(newRow["Foree ID"]));

      if (business) {
        newRow["Business Name"] = business.business_name;

        const amount = Number(newRow["Applicable Amount"]) || 0;

        const fee = calculateFee0(amount, business["1_bill_charges"]);

        newRow["Initiator Fee"] = fee;

        newRow["Tax On Initiator Fee"] = fee * 0.15;
      }

      return newRow;
    });

  showProcessResult();
}

function processT1() {
  processType = "t1";

  if (excelData.length === 0) {
    alert("Please upload a CSV file first.");

    return;
  }

  const filteredData = selectTime();

  if (filteredData === null) {
    return;
  }

  if (filteredData.length === 0) {
    alert("No records found in the selected Transaction Date Time range.");

    return;
  }

  processedData = filteredData.map((row) => {
    let newRow = {
      ...row,
    };

    if (newRow["Bill Creation Date"]) {
      newRow["Bill Creation Date"] = formatDateOnly(
        newRow["Bill Creation Date"],
      );
    }

    if (newRow["Bill Due Date"]) {
      newRow["Bill Due Date"] = formatDateOnly(newRow["Bill Due Date"]);
    }

    if (newRow["Transaction Date Time"]) {
      newRow["Transaction Date Time"] = formatDate(
        newRow["Transaction Date Time"],
      );
    }

    if (newRow["Initiator Settlement Date"]) {
      newRow["Initiator Settlement Date"] = formatDateOnly(
        newRow["Initiator Settlement Date"],
      );
    }

    const amount = Number(newRow["Applicable Amount"]) || 0;

    const fee = calculateFee(amount);

    newRow["Initiator Fee"] = fee;

    newRow["Tax On Initiator Fee"] = (fee * 0.15).toFixed(2);

    return newRow;
  });

  showProcessResult();
}

async function loadT0Businesses() {
  try {
    const response = await fetch(
      "https://api-ridercollection.foreebusiness.com/businessdata/get",
      {
        method: "POST",

        headers: {
          "Content-Type": "application/json",
        },

        body: JSON.stringify({
          settlement_period: "0",
        }),
      },
    );

    const result = await response.json();

    if (!result.status) {
      alert("Unable to fetch T+0 businesses.");

      return false;
    }

    t0Businesses.clear();

    result.data.forEach((item) => {
      const foreeId = String(item.business_id).slice(-4);

      t0Businesses.set(foreeId, item);
    });

    return true;
  } catch (error) {
    console.error("T+0 API Error:", error);

    alert("Unable to load T+0 businesses.");

    return false;
  }
}

// ============================================================
// T+0 FEE
// ============================================================

function calculateFee0(amount, charges) {
  amount = Number(amount);

  const percentage = parseFloat(charges);

  if (isNaN(percentage)) {
    return 0;
  }

  return amount * (percentage / 100);
}

// ============================================================
// STEP 5 — SHOW PROCESS RESULT
// ============================================================

function showProcessResult() {
  const result = document.getElementById("processStatus");

  result.style.display = "block";

  if (processType === "t0") {
    result.innerHTML =
      `<strong>T+0 processing completed.</strong><br>` +
      `${processedData.length} T+0 rows processed successfully.`;
  } else {
    result.innerHTML =
      `<strong>T+1 processing completed.</strong><br>` +
      `${processedData.length} T+1 rows processed successfully.`;
  }

  // Tell user TXT can now be uploaded

  document.getElementById("matchingStatus").innerHTML =
    "CSV processing completed. Waiting for TXT file...";
}

// ============================================================
// STEP 6 — TXT UPLOAD
// ============================================================

document.getElementById("folderInput").addEventListener("change", readTXT);

// ============================================================
// STEP 7 — READ TXT
// ============================================================

function readTXT(event) {
  const files = Array.from(event.target.files).filter((file) =>
    file.name.toLowerCase().endsWith(".txt"),
  );

  if (files.length === 0) {
    document.getElementById("status").style.display = "block";

    document.getElementById("status").innerHTML = "No TXT files selected.";

    return;
  }

  if (!processedData || processedData.length === 0) {
    alert("Please process the CSV file before uploading TXT files.");

    event.target.value = "";

    return;
  }

  txtFiles = files;

  // Start next step

  convertTXTToExcel();
}

// ============================================================
// STEP 8 — CONVERT TXT TO EXCEL DATA
// ============================================================

async function convertTXTToExcel() {
  const status = document.getElementById("status");

  try {
    status.style.display = "block";

    status.innerHTML = "Reading TXT files...";

    txtExcelData = [];

    txtExcelData = [
      [
        "1Link prefix",
        "Foree ID",
        "Map Reference Number",
        "Payer ID",
        "Payer Branch",
        "Amount",
        "Date",
        "Time",
        "Payment Channel",
        "Bank",
        "Stan",
        "Auth ID",
      ],
    ];

    let processedFiles = 0;

    let failedFiles = [];

    for (const file of txtFiles) {
      try {
        status.innerHTML =
          `Processing ${processedFiles + 1} of ${txtFiles.length}<br>` +
          `${file.name}`;

        const content = await file.text();

        const lines = content.split(/\r?\n/);

        for (const line of lines) {
          if (!line || line.trim() === "") {
            continue;
          }

          try {
            const columns = parseFixedWidth(line);

            if (columns && columns.length > 0) {
              txtExcelData.push(columns);
            }
          } catch (parseError) {
            console.warn("Line skipped:", file.name, parseError);
          }
        }

        processedFiles++;
      } catch (fileError) {
        console.error("File could not be read:", file.name, fileError);

        failedFiles.push(file.name);
      }
    }

    // ========================================================
    // CHECK DATA
    // ========================================================

    if (txtExcelData.length <= 1) {
      throw new Error("No transaction data was found.");
    }

    // ========================================================
    // NEXT STEP
    // ========================================================

    matchData();

    // ========================================================
    // STATUS
    // ========================================================

    if (failedFiles.length > 0) {
      status.innerHTML =
        `<strong>TXT conversion completed.</strong><br>` +
        `TXT files found: ${txtFiles.length}<br>` +
        `Files processed: ${processedFiles}<br>` +
        `Files failed: ${failedFiles.length}<br>` +
        `Transactions: ${txtExcelData.length - 1}`;
    } else {
      status.innerHTML =
        `<strong>TXT conversion completed.</strong><br>` +
        `Files processed: ${processedFiles}<br>` +
        `Transactions: ${txtExcelData.length - 1}`;
    }
  } catch (error) {
    console.error("Conversion Error:", error);

    status.style.display = "block";

    status.innerHTML = `<strong>Error:</strong> ${error.message}`;
  }
}

// ============================================================
// FIXED WIDTH TXT PARSER
// ============================================================

function parseFixedWidth(line) {
  // Make sure line is at least 145 characters

  line = line.padEnd(145, " ");

  const linkPrefix = line.substring(0, 8).trim();

  const foreeId = line.substring(13, 27).trim();

  const mapReferenceNumber = line.substring(33, 70).trim();

  const payerId = line.substring(70, 82).trim();

  const payerBranch = line.substring(82, 90).trim();

  const amount = Number(line.substring(90, 104)) / 100;

  const date = line.substring(104, 112).trim();

  const time = line.substring(112, 118).trim();

  const paymentChannel = line.substring(126, 127).trim();

  const bank = line.substring(127, 130).trim();

  const stan = line.substring(133, 139).trim();

  const authId = line.substring(139, 145).trim();

  return [
    linkPrefix,
    foreeId,
    mapReferenceNumber,
    payerId,
    payerBranch,
    amount,
    date,
    time,
    paymentChannel,
    bank,
    stan,
    authId,
  ];
}

// ============================================================
// NORMALIZE MATCH VALUE
// ============================================================

function normalizeMatchValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  return String(value).trim().replace(/^0+/, "") || "0";
}

// ============================================================
// NORMALIZE AMOUNT
// ============================================================

function normalizeAmount(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const amount = Number(String(value).replace(/,/g, "").trim());

  if (isNaN(amount)) {
    return null;
  }

  return Number(amount.toFixed(2));
}

// ============================================================
// STEP 9 — MATCH DATA
// ============================================================

function matchData() {
  const matchingStatus = document.getElementById("matchingStatus");

  if (!processedData || processedData.length === 0) {
    matchingStatus.innerHTML = "No processed CSV data available.";
    return;
  }

  if (!txtExcelData || txtExcelData.length <= 1) {
    matchingStatus.innerHTML = "No TXT transaction data available.";
    return;
  }

  // ============================================================
  // RESET
  // ============================================================

  matchedTXTData = [];
  unmatchedData = [];

  // ============================================================
  // TXT HEADERS
  // ============================================================

  const txtHeaders = [
    "1Link prefix",
    "Foree ID",
    "Map Reference Number",
    "Payer ID",
    "Payer Branch",
    "Amount",
    "Date",
    "Time",
    "Payment Channel",
    "Bank",
    "Stan",
    "Auth ID",
    "Match OR Unmatch",
  ];

  // ============================================================
  // CSV HEADERS
  // ============================================================

  const csvHeaders =
    processedData.length > 0 ? Object.keys(processedData[0]) : [];

  // ============================================================
  // CREATE CSV LOOKUP
  // ============================================================

  const processedLookup = new Map();

  processedData.forEach((row) => {
    const paymentRefId = normalizeMatchValue(row["Payment Ref ID"]);

    const paidByCustomer = normalizeAmount(row["Paid By Customer"]);

    if (paymentRefId === "" || paidByCustomer === null) {
      return;
    }

    const key = paymentRefId + "|" + paidByCustomer;

    if (!processedLookup.has(key)) {
      processedLookup.set(key, []);
    }

    processedLookup.get(key).push(row);
  });

  // ============================================================
  // TRACK MATCHED CSV ROWS
  // ============================================================

  const matchedCSV = new Set();

  // ============================================================
  // HEADERS
  // ============================================================

  matchedTXTData = [txtHeaders];

  const unmatchedHeaders = [
    ...csvHeaders.map((header) => "CSV - " + header),

    ...txtHeaders.map((header) => "TXT - " + header),
  ];

  unmatchedData = [unmatchedHeaders];

  // ============================================================
  // COUNTERS
  // ============================================================

  let matched = 0;
  let unmatchedTXT = 0;
  let unmatchedCSV = 0;

  // ============================================================
  // CHECK EVERY TXT TRANSACTION
  // ============================================================

  for (let i = 1; i < txtExcelData.length; i++) {
    const txtRow = txtExcelData[i];

    // Make sure TXT row has 12 values
    const cleanTXTRow = [...txtRow.slice(0, 12)];

    while (cleanTXTRow.length < 12) {
      cleanTXTRow.push("");
    }

    // ==========================================================
    // GET STAN
    // ==========================================================

    const stan = normalizeMatchValue(cleanTXTRow[10]);

    // ==========================================================
    // GET AMOUNT
    // ==========================================================

    const amount = normalizeAmount(cleanTXTRow[5]);

    // ==========================================================
    // CREATE MATCH KEY
    // ==========================================================

    const key = stan + "|" + (amount === null ? "" : amount);

    // ==========================================================
    // MATCH FOUND
    // ==========================================================

    if (processedLookup.has(key) && processedLookup.get(key).length > 0) {
      const csvRow = processedLookup.get(key).shift();

      matchedCSV.add(csvRow);

      // Add status directly to TXT row
      const matchedRow = [...cleanTXTRow, "Matched"];

      matchedTXTData.push(matchedRow);

      // Also update original TXT data
      txtExcelData[i] = [...cleanTXTRow, "Matched"];

      matched++;

      continue;
    }

    // ==========================================================
    // TXT UNMATCHED
    // ==========================================================

    const unmatchedTXTRow = [...cleanTXTRow, "Unmatched"];

    // Update original TXT data
    txtExcelData[i] = unmatchedTXTRow;

    // CSV side empty
    const emptyCSV = csvHeaders.map(() => "");

    unmatchedData.push([...emptyCSV, ...cleanTXTRow, "Unmatched"]);

    unmatchedTXT++;
  }

  // ============================================================
  // CHECK CSV ROWS THAT WERE NEVER MATCHED
  // ============================================================

  processedData.forEach((csvRow) => {
    if (matchedCSV.has(csvRow)) {
      return;
    }

    // ========================================================
    // CSV UNMATCHED
    // ========================================================

    const csvValues = csvHeaders.map((header) => csvRow[header] ?? "");

    // Empty TXT side
    const emptyTXT = txtHeaders.slice(0, 12).map(() => "");

unmatchedData.push([...csvValues, ...emptyTXT, "Unmatched"]);

    unmatchedCSV++;
  });

  // ============================================================
  // MAKE SURE ALL TXT ROWS HAVE STATUS
  // ============================================================

  txtExcelData[0] = txtHeaders;

  // ============================================================
  // SHOW RESULT
  // ============================================================

  matchingStatus.innerHTML =
    `<strong>${processType.toUpperCase()} matching completed.</strong><br>` +
    `Processed CSV rows: ${processedData.length}<br>` +
    `TXT transactions: ${txtExcelData.length - 1}<br>` +
    `Matched: ${matched}<br>` +
    `Unmatched TXT: ${unmatchedTXT}<br>` +
    `Unmatched CSV: ${unmatchedCSV}<br>` +
    `Total unmatched: ${unmatchedTXT + unmatchedCSV}`;
}

// ============================================================
// STEP 10 — DOWNLOAD PROCESSED CSV
// ============================================================

function downloadProcessedCSV() {
  if (!processedData || processedData.length === 0) {
    alert("Please process the CSV file first.");

    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(processedData);

  const csv = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob([csv], {
    type: "text/csv;charset=utf-8;",
  });

  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");

  link.href = url;

  if (processType === "t0") {
    link.download = "processed_settlement_report_t0.csv";
  } else {
    link.download = "processed_settlement_report_t1.csv";
  }

  document.body.appendChild(link);

  link.click();

  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

// ============================================================
// STEP 11 — DOWNLOAD COMPLETE TXT EXCEL REPORT
// ============================================================

function downloadTXTExcel() {
  console.log("DOWNLOAD TXT REPORT");

  if (!txtExcelData || txtExcelData.length <= 1) {
    alert("Please upload TXT files first.");
    return;
  }

  try {
    const workbook = XLSX.utils.book_new();

    // ==========================================================
    // TXT HEADERS
    // ==========================================================

    const txtHeaders = [
      "1Link prefix",
      "Foree ID",
      "Map Reference Number",
      "Payer ID",
      "Payer Branch",
      "Amount",
      "Date",
      "Time",
      "Payment Channel",
      "Bank",
      "Stan",
      "Auth ID",
      "Match OR Unmatch",
    ];

    // ==========================================================
    // CSV HEADERS
    // ==========================================================

    const csvHeaders =
      processedData.length > 0 ? Object.keys(processedData[0]) : [];

    const csvColumnCount = csvHeaders.length;

    // ==========================================================
    // GET STAN + AUTH ID FROM MAP REFERENCE NUMBER
    // ==========================================================

    function getStanAndAuth(reference) {
      const value = String(reference ?? "").trim();

      if (!value) {
        return {
          stan: "",
          authId: "",
        };
      }

      // Get the last 12 digits
      const last12 = value.slice(-12);

      // First 6 of the last 12 = STAN
      const stan = last12.slice(0, 6);

      // Last 6 of the last 12 = Auth ID
      const authId = last12.slice(6, 12);

      return {
        stan: stan,
        authId: authId,
      };
    }

    // ==========================================================
    // SHEET 1 — ALL TRANSACTIONS
    // ==========================================================

    const allTransactions = [];

    allTransactions.push(txtHeaders);

    // ==========================================================
    // ADD ALL MATCHED TXT TRANSACTIONS
    // ==========================================================

    for (let i = 1; i < matchedTXTData.length; i++) {
      const row = matchedTXTData[i];

      if (!Array.isArray(row) || row.length === 0) {
        continue;
      }

      const txtRow = [...row.slice(0, 12)];

      // Map Reference Number
      const reference = txtRow[2];

      const result = getStanAndAuth(reference);

      // First 6 of last 12 = STAN
      txtRow[10] = result.stan;

      // Last 6 of last 12 = Auth ID
      txtRow[11] = result.authId;

      allTransactions.push([...txtRow, "Matched"]);
    }

    // ==========================================================
    // ADD ALL UNMATCHED TRANSACTIONS
    // ==========================================================

    for (let i = 1; i < unmatchedData.length; i++) {
      const row = unmatchedData[i];

      if (!Array.isArray(row)) {
        continue;
      }

      // TXT section starts after CSV columns
      const txtStart = csvColumnCount;

      // Get 12 TXT fields
      const txtPart = row.slice(txtStart, txtStart + 12);

      // Check if TXT data exists
      const hasTXTData = txtPart.some(
        (value) =>
          value !== null && value !== undefined && String(value).trim() !== "",
      );

      // ========================================================
      // TXT-ONLY UNMATCHED
      // ========================================================

      if (hasTXTData) {
        const txtRow = [...txtPart];

        // Map Reference Number
        const reference = txtRow[2];

        const result = getStanAndAuth(reference);

        // First 6 of last 12 = STAN
        txtRow[10] = result.stan;

        // Last 6 of last 12 = Auth ID
        txtRow[11] = result.authId;

        allTransactions.push([...txtRow, "Unmatched"]);

        continue;
      }

      // ========================================================
      // CSV-ONLY UNMATCHED
      // ========================================================

      const csvValues = row.slice(0, csvColumnCount);

      const csvRow = {};

      csvHeaders.forEach((header, index) => {
        csvRow[header] = csvValues[index] ?? "";
      });

      // Payment Ref ID is used as STAN in matching
      const stan = csvRow["Payment Ref ID"] ?? "";

      // Paid By Customer is used as Amount
      const amount = csvRow["Paid By Customer"] ?? "";
      
      const bank = csvRow["Instrument Institution"] ?? "";
      
      const authId = csvRow["Payment Ref ID"] ?? "";


      // Add CSV-only transaction
      allTransactions.push([
        "",
        "",
        "",
        "",
        "",
        amount,
        "",
        "",
        "",
        bank,
        stan,
        authId,
        "Unmatched",
      ]);
    }

    // ==========================================================
    // CREATE ALL TRANSACTIONS SHEET
    // ==========================================================

    const allWorksheet = XLSX.utils.aoa_to_sheet(allTransactions);

    XLSX.utils.book_append_sheet(workbook, allWorksheet, "All Transactions");

    // ==========================================================
    // SHEET 2 — MATCHED
    // ==========================================================

    const matchedTransactions = matchedTXTData.map((row) =>
      Array.isArray(row) ? row : [row],
    );

    const matchedWorksheet = XLSX.utils.aoa_to_sheet(
      matchedTransactions.length > 0 ? matchedTransactions : [txtHeaders],
    );

    XLSX.utils.book_append_sheet(workbook, matchedWorksheet, "Matched");

    // ==========================================================
    // SHEET 3 — UNMATCHED
    // ==========================================================

    const unmatchedTransactions = unmatchedData.map((row) =>
      Array.isArray(row) ? row : [row],
    );

    const unmatchedWorksheet = XLSX.utils.aoa_to_sheet(
      unmatchedTransactions.length > 0
        ? unmatchedTransactions
        : [["No unmatched transactions"]],
    );

    XLSX.utils.book_append_sheet(workbook, unmatchedWorksheet, "Unmatched");

    // ==========================================================
    // COLUMN WIDTHS — ALL TRANSACTIONS
    // ==========================================================

    allWorksheet["!cols"] = [
      { wch: 15 },
      { wch: 15 },
      { wch: 25 },
      { wch: 20 },
      { wch: 20 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 20 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 18 },
    ];

    // ==========================================================
    // COLUMN WIDTHS — MATCHED
    // ==========================================================

    matchedWorksheet["!cols"] = [
      { wch: 15 },
      { wch: 15 },
      { wch: 25 },
      { wch: 20 },
      { wch: 20 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 20 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 18 },
    ];

    // ==========================================================
    // COLUMN WIDTHS — UNMATCHED
    // ==========================================================

    const unmatchedColumnCount =
      unmatchedTransactions.length > 0 ? unmatchedTransactions[0].length : 1;

    unmatchedWorksheet["!cols"] = Array.from(
      {
        length: unmatchedColumnCount,
      },
      () => ({
        wch: 25,
      }),
    );

    // ==========================================================
    // CREATE XLSX
    // ==========================================================

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    // ==========================================================
    // CREATE BLOB
    // ==========================================================

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    // ==========================================================
    // DOWNLOAD
    // ==========================================================

    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");

    link.href = url;
    link.download = "TXT_Report.xlsx";
    link.style.display = "none";

    document.body.appendChild(link);

    link.click();

    document.body.removeChild(link);

    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 1000);

    console.log("TXT_Report.xlsx downloaded");
  } catch (error) {
    console.error("TXT Excel Download Error:", error);

    alert("Unable to download TXT Report.\n\n" + error.message);
  }
}
