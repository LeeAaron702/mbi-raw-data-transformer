const PASSWORD = "mbitoner";
const AUTH_KEY = "mbi-authenticated";

const REQUIRED_COLUMNS = new Set(["ITEM", "CLASS", "DESCRIP", "ONHAND", "UNITMS", "VPARTNO", "PRICE2"]);
const FINAL_COLUMNS = [
    { output: "Item No.", source: "ITEM" },
    { output: "Manufacturer", source: "VPARTNO" },
    { output: "Class", source: "CLASS" },
    { output: "Product Description", source: "DESCRIP" },
    { output: "Qty", source: "QTY" },
    { output: "Un/Ms", source: "UNITMS" },
    { output: "Price", source: "PRICE2" },
];

const form = document.getElementById("uploadForm");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("submitBtn");
const gateForm = document.getElementById("gateForm");
const gateInput = document.getElementById("gatePassword");
const gateError = document.getElementById("gateError");
const gateShell = document.getElementById("passwordGate");
const appShell = document.getElementById("appContent");

initializePasswordGate();

form?.addEventListener("submit", async (event) => {
    event.preventDefault();

    const fileInput = form.querySelector("input[type='file']");
    const files = fileInput?.files;
    if (!files || files.length === 0) {
        setStatus("Please choose an .xls or .xlsx file.", "error");
        return;
    }

    const file = files[0];
    toggleBusy(true);

    try {
        const arrayBuffer = await file.arrayBuffer();
        const result = convertWorkbook(arrayBuffer, file.name);
        triggerDownload(result.blob, result.filename);
        setStatus("Finished file generated successfully.", "success");
        form.reset();
    } catch (error) {
        setStatus(error.message || "Unable to convert the file.", "error");
    } finally {
        toggleBusy(false);
    }
});

function convertWorkbook(arrayBuffer, filename) {
    // Validate file extension
    const ext = filename.slice(filename.lastIndexOf(".")).toLowerCase();
    if (![".xls", ".xlsx", ".xlsm"].includes(ext)) {
        throw new Error("Unsupported file type. Please upload an .xls or .xlsx file.");
    }

    // Read workbook
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
        throw new Error("Uploaded workbook is empty.");
    }

    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (rawData.length === 0) {
        throw new Error("Uploaded workbook has no data rows.");
    }

    // Normalize column names to uppercase
    const normalizedData = rawData.map((row) => {
        const normalized = {};
        for (const key of Object.keys(row)) {
            normalized[key.trim().toUpperCase()] = row[key];
        }
        return normalized;
    });

    // Check for required columns
    const sampleRow = normalizedData[0];
    const columns = new Set(Object.keys(sampleRow));
    const missing = getMissingColumns(columns);
    if (missing.length > 0) {
        throw new Error("Missing required columns: " + missing.sort().join(", "));
    }

    // Process rows: calculate QTY = ONHAND - ALOC/ALLOC, filter QTY > 0
    const processedRows = [];
    for (const row of normalizedData) {
        const onhand = parseNumber(row.ONHAND);
        const allocation = getAllocation(row);
        const qty = Math.max(0, onhand - allocation);

        if (qty > 0) {
            row.QTY = qty;
            processedRows.push(row);
        }
    }

    if (processedRows.length === 0) {
        throw new Error("No rows with available quantity after applying the rules.");
    }

    // Sort by VPARTNO, then ITEM
    processedRows.sort((a, b) => {
        const vpartnoA = String(a.VPARTNO || "").toLowerCase();
        const vpartnoB = String(b.VPARTNO || "").toLowerCase();
        if (vpartnoA !== vpartnoB) return vpartnoA.localeCompare(vpartnoB);
        const itemA = String(a.ITEM || "").toLowerCase();
        const itemB = String(b.ITEM || "").toLowerCase();
        return itemA.localeCompare(itemB);
    });

    // Build output data with final columns
    const outputData = processedRows.map((row) => {
        const out = {};
        for (const col of FINAL_COLUMNS) {
            if (col.source === "QTY") {
                out[col.output] = row.QTY;
            } else if (col.source === "PRICE2") {
                out[col.output] = parseNumber(row[col.source]);
            } else {
                out[col.output] = String(row[col.source] ?? "").trim();
            }
        }
        return out;
    });

    // Create output workbook
    const outputSheet = XLSX.utils.json_to_sheet(outputData, {
        header: FINAL_COLUMNS.map((c) => c.output),
    });

    // Apply formatting
    applyOutputFormatting(outputSheet, outputData);

    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Sheet1");

    // Generate file
    const wbout = XLSX.write(outputWorkbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const timestamp = formatTimestamp(new Date());
    return { blob, filename: `finished_${timestamp}.xlsx` };
}

function getMissingColumns(columns) {
    const missing = [];
    for (const req of REQUIRED_COLUMNS) {
        if (!columns.has(req)) {
            missing.push(req);
        }
    }
    // Check for ALOC or ALLOC
    if (!columns.has("ALOC") && !columns.has("ALLOC")) {
        missing.push("ALOC/ALLOC");
    }
    return missing;
}

function getAllocation(row) {
    if ("ALOC" in row) {
        return parseNumber(row.ALOC);
    }
    if ("ALLOC" in row) {
        return parseNumber(row.ALLOC);
    }
    return 0;
}

function parseNumber(value) {
    const num = parseFloat(value);
    return isNaN(num) ? 0 : num;
}

function applyOutputFormatting(sheet, data) {
    if (!sheet["!ref"]) return;

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const colIndices = {
        qty: FINAL_COLUMNS.findIndex((c) => c.output === "Qty"),
        unit: FINAL_COLUMNS.findIndex((c) => c.output === "Un/Ms"),
        price: FINAL_COLUMNS.findIndex((c) => c.output === "Price"),
        class: FINAL_COLUMNS.findIndex((c) => c.output === "Class"),
    };

    // Set column widths
    sheet["!cols"] = FINAL_COLUMNS.map((col, idx) => {
        if (idx === colIndices.class) {
            return { hidden: true };
        }
        if (col.output === "Product Description") {
            return { wch: 40 };
        }
        if (col.output === "Item No.") {
            return { wch: 15 };
        }
        if (col.output === "Manufacturer") {
            return { wch: 20 };
        }
        return { wch: 12 };
    });

    // Apply number format to Price column
    for (let row = 1; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: colIndices.price });
        const cell = sheet[cellAddress];
        if (cell) {
            cell.z = '"$"#,##0.00';
        }
    }
}

function formatTimestamp(date) {
    const pad = (n) => String(n).padStart(2, "0");
    return (
        date.getUTCFullYear() +
        pad(date.getUTCMonth() + 1) +
        pad(date.getUTCDate()) +
        "_" +
        pad(date.getUTCHours()) +
        pad(date.getUTCMinutes()) +
        pad(date.getUTCSeconds())
    );
}

function toggleBusy(isBusy) {
    submitBtn.disabled = isBusy;
    setStatus(isBusy ? "Processing file..." : "", isBusy ? "" : undefined);
}

function setStatus(message, variant) {
    statusEl.textContent = message;
    statusEl.classList.remove("error", "success");
    if (variant) {
        statusEl.classList.add(variant);
    }
}

function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
}

function initializePasswordGate() {
    if (!gateForm || !gateShell || !appShell) {
        return;
    }

    const isUnlocked = sessionStorage.getItem(AUTH_KEY) === "true";
    if (isUnlocked) {
        unlockApp();
    }

    gateForm?.addEventListener("submit", (event) => {
        event.preventDefault();
        const provided = gateInput?.value?.trim() ?? "";

        if (provided !== PASSWORD) {
            gateError.textContent = "Incorrect password.";
            gateInput?.focus();
            gateInput?.select();
            return;
        }

        sessionStorage.setItem(AUTH_KEY, "true");
        gateError.textContent = "";
        gateInput.value = "";
        unlockApp();
    });
}

function unlockApp() {
    gateShell?.setAttribute("hidden", "true");
    appShell?.removeAttribute("hidden");
    form?.querySelector("input[type='file']")?.focus();
}
