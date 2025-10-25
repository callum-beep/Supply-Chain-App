function populateFirstSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allSheet = ss.getSheetByName("All");

  const refSheets = {
    doc: ss.getSheetByName("RefDoc"),
    cat: ss.getSheetByName("RefCat"),
    pharm: ss.getSheetByName("RefPharm"),
    area: ss.getSheetByName("RefArea"),
    prod: ss.getSheetByName("RefProd")
  };

  const today = new Date();
  const currentMonth = today.getMonth() + 1;
  const validMonths2025 = today.getFullYear() === 2025 ? currentMonth - 1 : 12;

  const formatDoctorName = name => {
    if (!name) return "";
    name = name.toString().replace(/^Dr\.?\s*/i, "").trim();
    const parts = name.split(/\s+/);
    return parts
      .map((p, i) => i < 2 ? p.charAt(0).toUpperCase() + p.slice(1).toLowerCase() : p)
      .join(" ")
      .trim();
  };

  const cleanPharmacyName = name => name ? name.toString().toLowerCase().trim() : "";

  const buildMap = sheet => {
    const values = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
    const map = new Map();
    values.forEach(([key, value]) => {
      if (key && value) {
        map.set(key.toString().trim(), value);
      }
    });
    return map;
  };

  const docAssSheet = ss.getSheetByName("DocAss");
  const docAssMap = docAssSheet ? buildMap(docAssSheet) : new Map();

  for (const [key, sheet] of Object.entries(refSheets)) {
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Missing reference sheet: Ref${key.charAt(0).toUpperCase() + key.slice(1)}`);
      return;
    }
  }

  if (!allSheet) {
    allSheet = ss.insertSheet("All");
  } else {
    allSheet.clear();
  }

  const headers = [
    "Date", "Channel", "Pharmacy", "Area", "Product", "Units", "Sales",
    "Doctor", "Year", "Month", "Day", "Date Long",
    "Pharmacy Address", "First Appearance", "Category", "Week Number", "Assumed",
    "Tier", "Avg in 2025", "Avg Units in NSW 2025", "Avg Units in VIC 2025", "Avg Units in QLD 2025",
    "Avg Units in Other States 2025", "Avg Units in Unknown Areas 2025", "Check", "Doctor Region",
    "SerialNumber" // <-- ADDED at the end
  ];
  allSheet.appendRow(headers);

  const docMap = buildMap(refSheets.doc);
  const catMap = buildMap(refSheets.cat);
  const pharmMap = buildMap(refSheets.pharm);
  const areaMap = buildMap(refSheets.area);
  const prodMap = buildMap(refSheets.prod);

  const excludedDoctorKeywords = ["distributor", "distirbutor", "wholesaler", "dispensed", "montu", "burleigh heads", "bhc"];
  const seenDoctors = new Set();
  const seenPharmacies = new Set();
  const unitTracker = {};
  const rawRowData = [];
  const regionUnitsByDoctor = {};

  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!["EE", "CW", "Leafio All", "CDA All", "BLS", "Aeris", "Alternaleaf"].includes(name)) return;

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const channel = row[1];
      const pharmacy = row[2];
      let area = row[3];
      let product = row[4];
      const units = parseFloat(row[5]) || 0;
      const sales = parseFloat(row[6]) || 0;
      let doctor = row[7];
      const year = parseInt(row[8]);
      const month = parseInt(row[9]);
      const day = parseInt(row[10]);
      const rawAddress = row[12];

      if (!year || !month || !day || !pharmacy || !product || units === 0) continue;
      const jsDate = new Date(year, month - 1, day);
      if (isNaN(jsDate.getTime())) continue;

      const doctorLower = doctor?.toString().toLowerCase() || "";
      const doctorExcluded = excludedDoctorKeywords.some(k => doctorLower.includes(k)) || doctorLower === "cw placeholder";
      if ((area && area.toString().toLowerCase().includes("distributor")) || doctorExcluded) continue;

      const finalPharmacy = pharmMap.get(pharmacy?.toString().toLowerCase().trim()) || pharmacy;
      const pharmacyKey = cleanPharmacyName(finalPharmacy);

      let finalDoctor = "";
      let isAssumed = "No";

      if (doctor && doctor.toString().trim() !== "") {
        const formatted = formatDoctorName(doctor);
        finalDoctor = docMap.get(formatted) || formatted;
      } else {
        const docAssDoctor = docAssMap.get(pharmacyKey);
        if (docAssDoctor) {
          const formatted = formatDoctorName(docAssDoctor);
          finalDoctor = docMap.get(formatted) || formatted;
          isAssumed = "Yes (DocAss)";
        } else {
          const tracker = unitTracker[pharmacyKey];
          if (tracker) {
            let maxUnits = 0;
            let topDoctor = null;
            for (const [doc, count] of Object.entries(tracker)) {
              if (count > maxUnits) {
                maxUnits = count;
                topDoctor = doc;
              }
            }
            if (topDoctor) {
              const formatted = formatDoctorName(topDoctor);
              finalDoctor = docMap.get(formatted) || formatted;
              isAssumed = "Yes";
            }
          }

          if (!finalDoctor) {
            finalDoctor = `Unknown (${finalPharmacy})`;
            isAssumed = "Yes";
          }
        }
      }

      if (!unitTracker[pharmacyKey]) unitTracker[pharmacyKey] = {};
      unitTracker[pharmacyKey][finalDoctor] = (unitTracker[pharmacyKey][finalDoctor] || 0) + units;

      if (!area || area === "") {
        const cleanedAddress = rawAddress?.toString().toLowerCase().trim();
        area = areaMap.get(cleanedAddress) || "#N/A";
      }

      const areaClean = (area || "").toString().trim().toUpperCase();
      const region = ["NSW", "VIC", "QLD"].includes(areaClean) ? areaClean : (!areaClean || areaClean === "#N/A" ? "Unknown" : "Other");

      if (!regionUnitsByDoctor[finalDoctor]) {
        regionUnitsByDoctor[finalDoctor] = { total: 0, NSW: 0, VIC: 0, QLD: 0, Other: 0, Unknown: 0 };
      }
      if (year === 2025 && month <= validMonths2025) {
        regionUnitsByDoctor[finalDoctor][region] += units;
        regionUnitsByDoctor[finalDoctor].total += units;
      }

      const lookupKey = product?.toString().toLowerCase().replace(/\s+/g, " ").trim();
      const prodMatch = [...prodMap.keys()].find(k => k.toLowerCase().trim() === lookupKey);
      if (prodMatch) {
        product = prodMap.get(prodMatch);
      }

      const dateLong = Utilities.formatDate(jsDate, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      const pharmacyAddress = rawAddress ? `${rawAddress}, ${area}, Australia` : "";
      const category = catMap.get(finalDoctor) || "";

      const rowToAppend = [
        jsDate, channel, finalPharmacy, area, product, units, sales,
        finalDoctor, year, month, day, dateLong,
        pharmacyAddress, "", category, getWeekNumber(jsDate), isAssumed
      ];

      rawRowData.push({ row: rowToAppend, date: jsDate });
    }
  });

  rawRowData.sort((a, b) => a.date - b.date);
  const dataToAppend = [];

  const docRegionSheet = ss.getSheetByName("DocRegion");
  const docRegionMap = docRegionSheet ? buildMap(docRegionSheet) : new Map();

  rawRowData.forEach(entry => {
    const row = entry.row;
    const doctor = row[7];
    const pharmacy = row[2];
    const channel = row[1];

    const doctorKey = doctor;
    const pharmacyKey = pharmacy.toLowerCase().trim();

    let firstAppearance = "";
    if (channel !== "CDA") {
      if (!seenDoctors.has(doctorKey) && !seenPharmacies.has(pharmacyKey)) firstAppearance = "NEW PHARMACY AND DOCTOR";
      else if (!seenDoctors.has(doctorKey)) firstAppearance = "NEW DOCTOR";
      else if (!seenPharmacies.has(pharmacyKey)) firstAppearance = "NEW PHARMACY";
      if (firstAppearance && channel === "BLS") firstAppearance += " BLS";
    }

    seenDoctors.add(doctorKey);
    seenPharmacies.add(pharmacyKey);
    row[13] = firstAppearance;

    const stats = regionUnitsByDoctor[doctor] || {};
    const avg = validMonths2025 > 0 ? +(stats.total / validMonths2025).toFixed(1) : 0;
    const regionAvgs = ["NSW", "VIC", "QLD", "Other", "Unknown"].map(r => validMonths2025 > 0 ? +(stats[r] / validMonths2025).toFixed(2) : 0);
    let tier = "";
    if (avg > 499) tier = "Tier 1";
    else if (avg >= 150) tier = "Tier 2";
    else if (avg >= 30.1) tier = "Tier 3";
    else if (avg > 0) tier = "Tier 4";
    const checkSum = +(avg - regionAvgs.reduce((a, b) => a + b, 0)).toFixed(2);

    const mappedRegion = docRegionMap.get(formatDoctorName(doctorKey));
    const doctorRegion = mappedRegion && !doctor.toLowerCase().startsWith("unknown (") ? mappedRegion : (row[3] || "").trim();

    row.push(tier, avg, ...regionAvgs, checkSum, doctorRegion);

    // --- ADDED: SerialNumber computation and append at the very end ---
    // SerialNumber = ddMMyyyy + first 12 alphanumeric chars of pharmacy name (lowercased)
    const jsDate = row[0];
    const serialDate = Utilities.formatDate(jsDate, ss.getSpreadsheetTimeZone(), "ddMMyyyy");
    const serialPharm = pharmacy.toString().toLowerCase().replace(/[^a-z0-9]/g, "").substring(0, 12);
    const serialNumber = serialDate + serialPharm;
    row.push(serialNumber);
    // --- END ADDED ---

    dataToAppend.push(row);
  });

  if (dataToAppend.length > 0) {
    dataToAppend.sort((a, b) => b[0] - a[0]);
    allSheet.getRange(2, 1, dataToAppend.length, headers.length).setValues(dataToAppend);
    allSheet.getRange("A2:A" + allSheet.getLastRow()).setNumberFormat("dd/MM/yyyy");
  }
}

function getWeekNumber(date) {
  const startDate = new Date(date.getFullYear(), 0, 1);
  const days = Math.floor((date - startDate) / (24 * 60 * 60 * 1000));
  return Math.ceil((days + 1) / 7);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Emoji Tools")
    .addItem("📥 Run Populate Script", "populateFirstSheet")
    .addItem("🟡 Show Unknown Areas", "showUnknownAreas")
    .addItem("🟤 Show Unknown Doctors", "showUnknownDoctors")
    .addItem("🟠 Show Assumed Doctors", "showAssumedDoctors")
    .addToUi();
}

function showUnknownAreas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheet = ss.getSheetByName("All");
  const output = ss.getSheetByName("Unknown Areas") || ss.insertSheet("Unknown Areas");
  output.clear();

  const data = allSheet.getDataRange().getValues();
  const headers = data[0];
  const areaIndex = headers.indexOf("Area");

  const filtered = data.filter((row, i) => i > 0 && (!row[areaIndex] || row[areaIndex] === "#N/A"));
  if (filtered.length) {
    output.appendRow(headers);
    filtered.forEach(row => output.appendRow(row));
  }
}

function showUnknownDoctors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheet = ss.getSheetByName("All");
  const output = ss.getSheetByName("Unknown Doctors") || ss.insertSheet("Unknown Doctors");
  output.clear();

  const data = allSheet.getDataRange().getValues();
  const headers = data[0];
  const doctorIndex = headers.indexOf("Doctor");

  const filtered = data.filter((row, i) => i > 0 && row[doctorIndex]?.toString().startsWith("Unknown ("));
  if (filtered.length) {
    output.appendRow(headers);
    filtered.forEach(row => output.appendRow(row));
  }
}

function showAssumedDoctors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheet = ss.getSheetByName("All");
  const output = ss.getSheetByName("Assumed Doctors") || ss.insertSheet("Assumed Doctors");
  output.clear();

  const data = allSheet.getDataRange().getValues();
  const headers = data[0];
  const assumedIndex = headers.indexOf("Assumed");

  const filtered = data.filter((row, i) => i > 0 && row[assumedIndex] && row[assumedIndex].toString().includes("Yes"));
  if (filtered.length) {
    output.appendRow(headers);
    filtered.forEach(row => output.appendRow(row));
  }
}
