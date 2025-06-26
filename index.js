// ====================================================================
// CONSTANTS
// ====================================================================

// ====================================================================
// EMPLOYEE MANAGEMENT SYSTEM
// ====================================================================
function getEmployeeEmails() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const empSheet = ss.getSheetByName("EmployeeDB");
        if (!empSheet) return [];

        const lastRow = empSheet.getLastRow();
        if (lastRow < 2) return [];

        // Get all data from column E (Email column)
        // Column E = 5th column (indices start at 1)
        return empSheet.getRange(2, 5, lastRow-1, 1)
            .getValues()
            .flat()
            .filter(email => {
                // Only return valid emails
                return email && typeof email === 'string' && email.includes('@');
            })
            .sort(); // Sort alphabetically
    } catch (e) {
        console.error("Email Fetch Error:", e);
        return [];
    }
}


// Stock management
const STOCK_COLUMNS = {
    STOCK: 5,        // Column E
    TIMESTAMP: 9     // Column I
};

const STOCK_UI_CONSTANTS = {
    COLUMN_WIDTH: 180,
    MIN_ROW_HEIGHT: 21,
    FLASH_DURATION: 300,
    MAX_TIMESTAMPS: 3
};

const STOCK_COLORS = {
    SUCCESS: "#C8E6C9",    // Light green with better contrast
    FIXED: "#E3F2FD",     // Light blue with better contrast
    ERROR: "#FFCDD2"      // Light red with better contrast
};

// Device tracking
const DEVICE_SHEETS = {
    LAPTOPS: 'Laptops',
    TABLETS: 'Tablets',
    PHONES: 'Smartphones',
    DESKTOPS: 'Desktops'
};

// History logging
const HISTORY_SHEET_NAME = "Device History Log";

// Main sheet reference - CHANGE THIS TO YOUR ACTUAL SHEET NAME
const MAIN_SHEET_NAME = 'IT_Storage';

// ====================================================================
// STOCK MANAGEMENT FUNCTIONS (FULLY FIXED)
// ====================================================================
function formatHistoryLogColumns() {
    try {
        const historySheet = SpreadsheetApp.getActive().getSheetByName(HISTORY_SHEET_NAME);
        if (!historySheet) {
            showToast('❌ Device History Log not found', '⚠️ Error', 5);
            return;
        }

        // Set header text colors and bold
        historySheet.getRange("A1:K1")
            .setFontColor("#ffffff")
            .setBackground("#2c3e50")
            .setFontWeight("bold");

        // Set alternating colors for rows
        const lastRow = Math.max(historySheet.getLastRow(), 2);
        const range = historySheet.getRange(2, 1, lastRow-1, 11);
        range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

        // Set dates in standard format
        historySheet.getRange("H:I")
            .setNumberFormat("dd/MM/yyyy HH:mm");

        showToast('✅ Device History Log formatted successfully', '📑 Format Complete', 5);
    } catch (e) {
        showToast('❌ Formatting failed: ' + e.message, '⚠️ Error', 5);
    }
}
function updateStock(change) {
    try {
        if (typeof change !== 'number') {
            throw new Error('Change must be a number');
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        const cell = sheet.getActiveCell();
        const row = cell.getRow();
        const col = cell.getColumn();

        // Validate we're in the stock sheet (case-insensitive)
        if (sheet.getName().toLowerCase() !== MAIN_SHEET_NAME.toLowerCase()) {
            const msg = `❌ Change stock on the "${MAIN_SHEET_NAME}" sheet\n` +
                `You are on: "${sheet.getName()}"`;
            ss.toast(msg, "⚠️ Stock Error", 6);
            return;
        }

        // Validate cell position - must be in the stock column (E)
        if (col !== STOCK_COLUMNS.STOCK) {
            const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
            ss.toast(`❌ Select a STOCK value in Column E\nCurrent column: ${cols[col-1] || col}`, "⚠️ Stock Error", 5);
            return;
        }

        // Don't allow changes to header row
        if (row === 1) {
            ss.toast("❌ Can't update headers", "⚠️ Error", 3);
            return;
        }

        const stockCell = sheet.getRange(row, STOCK_COLUMNS.STOCK);
        const timestampCell = sheet.getRange(row, STOCK_COLUMNS.TIMESTAMP);
        const username = Session.getActiveUser().getEmail() || "Unknown User";

        // Handle non-string values in timestamp cell
        const currentContent = timestampCell.getValue();
        if (currentContent && typeof currentContent !== 'string') {
            timestampCell.clearContent();
        }

        // Get and validate current stock value
        const currentVal = stockCell.getValue();
        let currentNum = 0;

        if (typeof currentVal === 'number') {
            currentNum = currentVal;
        } else if (typeof currentVal === 'string') {
            currentNum = Number(currentVal.replace(',', '.')) || 0; // Handle comma decimals
        } else if (currentVal === null || currentVal === undefined || currentVal === '') {
            currentNum = 0;
        } else {
            throw new Error('Stock value must be a number or numeric string');
        }

        const newVal = currentNum + change;

        // Prevent negative stock
        if (newVal < 0) {
            ss.toast(`❌ Stock cannot be negative\n(${currentNum} + ${change} = ${newVal})`, "⚠️ Stock Error", 5);
            return;
        }

        // Batch updates for better performance
        const now = new Date();
        const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        const sign = change > 0 ? '+' : '-';
        const newEntry = `${formattedDate} ${sign}${Math.abs(change)} (${username})`;

        // Perform the update
        stockCell.setValue(newVal);

        // Visual feedback and metadata updates
        flashCell(stockCell, STOCK_COLORS.SUCCESS);
        updateTimestampCell(timestampCell, newEntry);
        adjustCellDimensions(timestampCell);

        // Update stock cell comment
        stockCell.setComment(
            `Last update: ${formattedDate}\n` +
            `By: ${username}\n` +
            `Change: ${sign}${Math.abs(change)}`
        );

        // Success message
        const action = change > 0 ? "Increased by" : "Decreased by";
        const color = change > 0 ? "#0d652d" : "#a50e0e";
        const icon = change > 0 ? "✅" : "🔽";
        ss.toast(`${icon} Stock ${action} ${Math.abs(change)}\nNew value: ${newVal}`, "✓ Stock Updated", 3);

    } catch (error) {
        Logger.log(`Stock Error: ${error.message}`);
        const ui = SpreadsheetApp.getUi();
        ui.alert(`❌ Stock Update Error: ${error.message}`);
    }
}

function incrementStock() { updateStock(1); }
function decrementStock() { updateStock(-1); }

function updateTimestampCell(cell, newEntry) {
    // Handle all data types safely
    let currentString = "";
    try {
        const content = cell.getValue();

        if (typeof content === 'string') {
            currentString = content;
        } else if (content instanceof Date) {
            currentString = Utilities.formatDate(content, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } else if (typeof content === 'number') {
            currentString = String(content);
        } else if (content && content.toString) {
            currentString = content.toString();
        }
    } catch (e) {
        Logger.log(`Timestamp conversion error: ${e}`);
        currentString = "";
    }

    let entries = currentString ?
        currentString.split('\n').filter(e => e.trim()) : [];

    entries.unshift(newEntry);
    entries = entries.slice(0, STOCK_UI_CONSTANTS.MAX_TIMESTAMPS);

    cell.setValue(entries.join('\n')).setWrap(true);
}

function adjustCellDimensions(cell) {
    const sheet = cell.getSheet();
    const row = cell.getRow();
    const lineCount = (cell.getValue() || "").split('\n').length;

    sheet.setRowHeight(
        row,
        Math.max(STOCK_UI_CONSTANTS.MIN_ROW_HEIGHT, lineCount * STOCK_UI_CONSTANTS.MIN_ROW_HEIGHT)
    );
}

function flashCell(cell, color) {
    const originalColor = cell.getBackground();

    cell.setBackground(color);
    SpreadsheetApp.flush();

    Utilities.sleep(STOCK_UI_CONSTANTS.FLASH_DURATION);
    cell.setBackground(originalColor);
}

// Fix existing timestamp issues automatically
function clearTimestampsAndComments() {
    try {
        const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
        if (!sheet) {
            SpreadsheetApp.getActive().toast("❌ IT_Storage sheet not found", "Clear Error", 5);
            return;
        }

        const lastRow = Math.max(sheet.getLastRow(), 2);

        // Clear ALL timestamp data (Column I)
        const timestampRange = sheet.getRange(2, STOCK_COLUMNS.TIMESTAMP, lastRow-1, 1);
        timestampRange.clearContent(); // Clears both values and comments

        // Clear all comments in stock column (Column E)
        const stockRange = sheet.getRange(2, STOCK_COLUMNS.STOCK, lastRow-1, 1);
        stockRange.clearNote(); // Remove all comment histories

        SpreadsheetApp.getUi().alert(`✅ Cleared ALL timestamp history for ${lastRow-1} items!`);
    } catch (e) {
        Logger.log(`Clear Error: ${e}`);
        SpreadsheetApp.getUi().alert(`❌ Clear Error: ${e.message}`);
    }
}


// ====================================================================
// DEVICE LOG HISTORY SYSTEM
// ====================================================================

function createHistorySheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create history sheet if it doesn't exist
    let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) {
        historySheet = ss.insertSheet(HISTORY_SHEET_NAME);

        // Create headers
        historySheet.appendRow([
            'Event ID', 'Timestamp', 'Event Type', 'Device Type', 'Serial Number',
            'Device Specs', 'Assigned To', 'Assigned Date',
            'Return Date', 'Duration (Days)', 'Responsible User'
        ]);

        // Format headers
        const header = historySheet.getRange("A1:K1");
        header.setFontWeight('bold')
            .setBackground('#2c3e50')
            .setFontColor('#ecf0f1');

        // Set column widths
        const widths = [150, 150, 100, 120, 150, 300, 200, 150, 150, 100, 200];
        for (let i = 0; i < widths.length; i++) {
            historySheet.setColumnWidth(i + 1, widths[i]);
        }

        // Freeze header row and auto-format
        historySheet.setFrozenRows(1);
        historySheet.getDataRange()
            .setVerticalAlignment("top")
            .setWrap(true);

        // Apply alternating row colors
        const range = historySheet.getRange(2, 1, 1000, 11);
        range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }

    return historySheet;
}

// Log device assignment/return events
function logDeviceEvent(deviceSheet, row, eventType) {
    try {
        const historySheet = createHistorySheet();
        const timestamp = new Date();
        const currentUser = Session.getActiveUser().getEmail() || "Unknown User";

        // Get device details
        const deviceSheetObj = SpreadsheetApp.getActive()
            .getSheetByName(deviceSheet);

        if (!deviceSheetObj) {
            Logger.log(`Device sheet ${deviceSheet} not found`);
            return;
        }

        // Validate row exists
        if (row > deviceSheetObj.getLastRow()) {
            throw new Error(`Row ${row} exceeds sheet size`);
        }

        const deviceDataRange = deviceSheetObj.getRange(row, 1, 1, 7);
        const deviceValues = deviceDataRange.getValues()[0];

        const [id, serial, specs, assignedTo, assignDate, returnDate] = deviceValues;

        // Calculate duration
        let duration = "";
        if (eventType === "Return" && assignDate && returnDate) {
            try {
                const asDate = new Date(assignDate);
                const rtDate = new Date(returnDate);
                // Prevent date inversion errors
                if (rtDate >= asDate) {
                    const diffTime = rtDate - asDate;
                    duration = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                } else {
                    duration = "Date error";
                }
            } catch (e) {
                duration = "N/A";
            }
        }

        // Format dates safely
        const formatDate = (d) => {
            if (!d) return "";
            if (d instanceof Date) {
                return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            }
            if (typeof d === 'string') return d;
            try {
                return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            } catch (e) {
                return "Invalid date";
            }
        };

        // Create log entry
        historySheet.insertRowBefore(2);
        historySheet.getRange(2, 1, 1, 11).setValues([[
            Utilities.getUuid().slice(0, 8),
            Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            eventType,
            deviceSheet,
            serial,
            specs,
            assignedTo,
            formatDate(assignDate),
            eventType === "Return" ? formatDate(returnDate) : "",
            duration,
            currentUser
        ]]);

        // Add color coding
        const rowRange = historySheet.getRange(2, 1, 1, 11);
        if (eventType === "Assign") {
            rowRange.setBackground("#e8f5e9");
            rowRange.setBorder(true, true, true, true, true, true, "#c1e1cb", SpreadsheetApp.BorderStyle.SOLID);
        } else {
            rowRange.setBackground("#fbe9e7");
            rowRange.setBorder(true, true, true, true, true, true, "#f9c9c3", SpreadsheetApp.BorderStyle.SOLID);
        }

        // Auto-resize new row
        historySheet.setRowHeight(2, STOCK_UI_CONSTANTS.MIN_ROW_HEIGHT * 2);

    } catch (error) {
        Logger.log(`History Error: ${error.message}`);
    }
}

function clearHistory() {
    const historySheet = SpreadsheetApp.getActive().getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) return;

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Clear History',
        'Are you sure you want to clear ALL history records? This cannot be undone.',
        ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
        historySheet.clearContents();
        historySheet.appendRow([
            'Event ID', 'Timestamp', 'Event Type', 'Device Type', 'Serial Number',
            'Device Specs', 'Assigned To', 'Assigned Date',
            'Return Date', 'Duration (Days)', 'Responsible User'
        ]);
        historySheet.getRange(1, 1, 1, 11)
            .setFontWeight('bold')
            .setBackground('#2c3e50')
            .setFontColor('#ecf0f1');

        ui.alert("✅ History log cleared successfully");
    }
}

// ====================================================================
// DEVICE ASSIGNMENT RETURN FUNCTIONS
// ====================================================================

function createDeviceSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Object.values(DEVICE_SHEETS).forEach(name => {
        try {
            if (!ss.getSheetByName(name)) {
                const sheet = ss.insertSheet(name);

                sheet.appendRow([
                    'ID', 'Serial Number', 'Specifications',
                    'Assigned To', 'Date Assigned',
                    'Date Returned', 'Status'
                ]);

                // Format headers
                const header = sheet.getRange("A1:G1");
                header.setFontWeight('bold')
                    .setBackground('#f0f0f0')
                    .setWrap(true);

                // Set column widths
                const widths = [120, 150, 300, 200, 150, 150, 100];
                for (let i = 0; i < widths.length; i++) {
                    sheet.setColumnWidth(i + 1, widths[i]);
                }

                // Format date columns
                sheet.getRange("E:F").setNumberFormat("dd/MM/yyyy HH:mm");
            }
        } catch (e) {
            Logger.log(`Error creating ${name} sheet: ${e}`);
        }
    });
}



function getEmployeeInfo(email) {
    const sheet = ss.getSheetByName("EmployeeDB");
    const data = sheet.getDataRange().getValues();

    // Find matching employee
    return data.find(row => row[4] === email) || [];
}

function assignDevice() {
    const template = HtmlService.createTemplateFromFile('mobileAssignDialog');
    template.employeeEmails = getEmployeeEmails();
    const html = template.evaluate().setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, '📱 Assign Device');
}

function assignDeviceRequest(sheetName, row, recipient) {
    function assignDeviceRequest(sheetName, row, recipient) {
        // Validate email exists in database
        const validEmails = getEmployeeEmails();
        if (!validEmails.includes(recipient)) {
            return {
                success: false,
                message: `❌ "${recipient}" is not a valid employee email. Please choose from EmployeeDB.`
            };
        }

        // ... rest of existing code ...
    }

    try {
        // Sanitize input
        recipient = recipient.toString().replace(/</g, "&lt;").replace(/>/g, "&gt;");

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) throw new Error('Device sheet not found');

        // Validate device existence
        if (row > sheet.getLastRow() || row < 2) {
            throw new Error('Selected device not found in sheet');
        }

        // Set assignment details
        sheet.getRange(row, 4).setValue(recipient);          // Assigned To
        sheet.getRange(row, 5).setValue(new Date());         // Date Assigned
        sheet.getRange(row, 6).setValue('');                 // Clear return date
        sheet.getRange(row, 7).setValue('Assigned');         // Status

        // Log assignment to history
        logDeviceEvent(sheetName, row, "Assign");

        return {
            success: true,
            message: `✅ ${getDeviceType(sheetName)} assigned to ${recipient}!`
        };
    } catch (error) {
        return {
            success: false,
            message: `❌ Error: ${error.message}`
        };
    }
}

function getDeviceType(sheetName) {
    switch(sheetName) {
        case DEVICE_SHEETS.LAPTOPS: return "Laptop";
        case DEVICE_SHEETS.TABLETS: return "Tablet";
        case DEVICE_SHEETS.PHONES: return "Smartphone";
        case DEVICE_SHEETS.DESKTOPS: return "Desktop PC";
        default: return "Device";
    }
}

function returnDevice() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('🔙 Return Device', 'Enter recipient name/email:', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
        const recipient = response.getResponseText().trim().toLowerCase();
        const devices = findDevicesByRecipient(recipient);

        if (devices.length === 0) {
            ui.alert('❌ No devices found for: ' + recipient);
            return;
        }

        // Return all devices and log history
        const errors = [];

        devices.forEach(device => {
            try {
                const sheet = SpreadsheetApp.getActive().getSheetByName(device.sheet);
                const assignDate = sheet.getRange(device.row, 5).getValue();
                const returnDate = new Date();

                // Skip if no assignment date
                if (!assignDate) {
                    errors.push(`${device.serial}: No assignment date found`);
                    return;
                }

                // Log date inversion issues without disrupting workflow
                if (returnDate < new Date(assignDate)) {
                    errors.push(`${device.serial}: Return date before assign date`);
                    return;
                }

                sheet.getRange(device.row, 6).setValue(returnDate);  // Return date
                sheet.getRange(device.row, 7).setValue('Available'); // Status

                // Log return to history
                logDeviceEvent(device.sheet, device.row, "Return");
            } catch (e) {
                errors.push(`${device.serial}: ${e.message}`);
            }
        });

        if (errors.length > 0) {
            ui.alert(`⚠️ Completed with ${errors.length} errors:\n\n${errors.join('\n')}`);
        } else {
            ui.alert(`✅ ${devices.length} devices returned for: ${recipient}`);
        }
    }
}

function findDevicesByRecipient(recipient) {
    const results = [];
    const searchTerm = recipient.toLowerCase();

    Object.values(DEVICE_SHEETS).forEach(name => {
        const sheet = SpreadsheetApp.getActive().getSheetByName(name);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();

        // Start from row 1 (skip header)
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const assignedTo = (row[3] || '').toString().toLowerCase();
            const status = row[6] || '';

            if (assignedTo.includes(searchTerm) && status === 'Assigned') {
                results.push({
                    sheet: name,
                    row: i + 1,
                    serial: row[1] || ""
                });
            }
        }
    });

    return results;
}

function getAvailableDevices() {
    const devices = {};

    Object.values(DEVICE_SHEETS).forEach(name => {
        devices[name] = [];
        const sheet = SpreadsheetApp.getActive().getSheetByName(name);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const status = (row[6] || '').toString();
            if (status === 'Available' && row[1]) {
                devices[name].push({
                    row: i + 1,
                    serial: row[1].toString(),
                    specs: row[2] ?
                        (typeof row[2] === 'string' ?
                            (row[2].length > 40 ?
                                row[2].substring(0, 40) + '...' :
                                row[2]) :
                            row[2].toString()) :
                        ''
                });
            }
        }
    });

    return devices;
}

// ====================================================================
// DIALOG SUPPORT FUNCTIONS
// ====================================================================

function getAvailableDevicesByType(type) {
    const allDevices = getAvailableDevices();
    return allDevices[type] || [];
}

function getDeviceTypeOptions() {
    return Object.entries(DEVICE_SHEETS)
        .map(([key, name]) => `<option value="${name}">${key}</option>`)
        .join('');
}

// ====================================================================
// SYSTEM INITIALIZATION & MENU
// ====================================================================

function onOpen(e) {
    // Only run when UI is available
    if (!e || e.authMode !== ScriptApp.AuthMode.NONE) {
        try {
            const ui = SpreadsheetApp.getUi();
            const menu = ui.createMenu('📊 Inventory Manager');

            // Add inside onOpen():
            menu.addSeparator()
                .addItem('👥 Employee Database', 'showEmployeeDB')
                .addToUi();


            // Stock operations
            menu.addItem('➕ Increase Stock', 'incrementStock');
            menu.addItem('➖ Decrease Stock', 'decrementStock');

            // Device management
            menu.addSeparator();
            menu.addItem('📱 Assign Device', 'assignDevice');
            menu.addItem('🔙 Return Device', 'returnDevice');

            // Create Device History menu
            const devHistoryMenu = ui.createMenu('📜 Device History')
                .addItem('View History', 'viewHistory')
                .addItem('Clear History', 'clearHistory')
                .addItem('Fix Formatting', 'formatHistoryLogColumns');

            // Create Audit Log menu
            const auditMenu = ui.createMenu('📝 Audit Log')
                .addItem('View Log', 'auditLog_show')
                .addItem('Initialize', 'auditLog_initialize')
                .addItem('Clear Log', 'auditLog_clear');

            // Add submenus
            menu.addSeparator();
            menu.addSubMenu(devHistoryMenu);
            menu.addSubMenu(auditMenu);

            // Maintenance
            menu.addSeparator();
            menu.addItem('⚙️ Fix Settings', 'fixAllSettings');
            menu.addItem('🧹 Clear Timestamps', 'clearTimestampsAndComments')

            menu.addSeparator()
                .addItem('🧹 Clear Timestamps', 'fixTimestampColumns')  // New option added HERE

            menu.addToUi();

        } catch (error) {
            Logger.log("Menu Error: " + error.message);
        }
    }
}
function showEmployeeDB() {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("EmployeeDB");
    if (sheet) ss.setActiveSheet(sheet);
}



function reinitializeSystem() {
    try {
        createDeviceSheets();
        createHistorySheet();
        SpreadsheetApp.getUi().alert("✅ System reinitialized successfully!");
    } catch (e) {
        SpreadsheetApp.getUi().alert(`❌ Reinit Error: ${e.message}`);
    }
}

function viewHistory() {
    const ss = SpreadsheetApp.getActive();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);

    if (historySheet) {
        ss.setActiveSheet(historySheet);
        const ui = SpreadsheetApp.getUi();
        ui.alert(
            "Device History Log",
            `You are now viewing the complete device assignment history.\n\n✅ Assignments shown in green\n🔴 Returns shown in orange`,
            ui.ButtonSet.OK
        );
    } else {
        SpreadsheetApp.getUi().alert("❌ History sheet not found. Use Reinitialize System");
    }
}

function onEdit(e) {
    const SHEET_NAME = e.range.getSheet().getName();
    const ROW = e.range.getRow();
    const COL = e.range.getColumn();

    // Main sheet maintenance
    if (SHEET_NAME === MAIN_SHEET_NAME && ROW > 1 && COL > 1) {
        const idCell = e.range.getSheet().getRange(ROW, 1);
        try {
            if (!idCell.getValue()) {
                idCell.setValue(Utilities.getUuid().substring(0, 8));
            }
        } catch (e) {
            Logger.log(`ID Gen Error: ${e}`);
        }
        return;
    }

    // Device status handling
    if (Object.values(DEVICE_SHEETS).includes(SHEET_NAME) && ROW > 1) {
        const statusCell = e.range.getSheet().getRange(ROW, 7);

        // Set status when entering serial/specs
        try {
            if ((COL >= 2 && COL <= 3) && !statusCell.getValue()) {
                statusCell.setValue('Available');
            }
        } catch (e) {
            Logger.log(`Status Set Error: ${e}`);
        }
    }
}

// Required for dialogs
function getDevicesForDialog() {
    return getAvailableDevices();
}

// ====================================================================
// EMPLOYEE DATABASE FUNCTIONS
// ====================================================================

function getAllActiveEmails() {
    try {
        const ss = SpreadsheetApp.getActive();
        const empSheet = ss.getSheetByName("EmployeeDB");
        if (!empSheet) return [];

        const data = empSheet.getDataRange().getValues();
        const emails = [];

        // Start from row 1 (skip header), col6 = email (F), col8 = status (H)
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[7] === "Active" && row[5]) { // Status=Active and email exists
                emails.push(row[5]);
            }
        }
        return emails.sort();
    } catch (e) {
        return [];
    }
}

// ====================================================================
// AUDIT LOG SYSTEM WITH FULL EDIT TRACKING
// ====================================================================

// ... (Previous code remains unchanged) ...

// ====================================================================
// AUDIT LOG SYSTEM WITH FULL EDIT TRACKING
// ====================================================================

const AUDIT_LOG_SHEET_NAME = 'Audit Log';
const AUDIT_COLORS = {
    STOCK_UP: '#81C784',    // Medium green for increases (better contrast)
    STOCK_DOWN: '#E57373',  // Medium red for decreases (better contrast)
    USER_EDIT: '#64B5F6',   // Medium blue for manual edits
    DEVICE_CHANGE: '#81C784'  // Same as stock up for consistency
};

// Create or get the audit log sheet
function auditLog_createSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);

    if (!auditSheet) {
        auditSheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);

        // Create headers and format
        // NEW: Added 'Item_name' as the last column
        const headers = ['Timestamp', 'Action', 'Sheet', 'Cell', 'User', 'Old Value', 'New Value', 'Details', 'Item_name'];
        auditSheet.appendRow(headers);

        const headerRange = auditSheet.getRange('A1:I1'); // CHANGED: From A1:H1 to A1:I1
        headerRange
            .setBackground('#2c3e50')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontSize(11);

        // Set column widths
        // NEW: Added width for Item_name column
        const widths = [150, 120, 100, 80, 150, 120, 120, 300, 150]; // MODIFIED: Added 150 for Item_name
        widths.forEach((w, i) => auditSheet.setColumnWidth(i+1, w));

        // Basic formatting
        auditSheet.setFrozenRows(1);
        auditSheet.getDataRange().setWrap(true);
        auditSheet.autoResizeRows(1, 1);
    }

    return auditSheet;
}

// Handle ALL manual edit events
function auditLog_onEdit(e) {
    try {
        const range = e.range;
        const sheet = range.getSheet();
        const sheetName = sheet.getName();

        // Skip if edit in audit log itself
        if (sheetName === AUDIT_LOG_SHEET_NAME) return;

        // Only track IT_Storage and device sheets
        const trackableSheets = ['IT_Storage', 'Laptops', 'Tablets', 'Smartphones', 'Desktops'];
        if (!trackableSheets.includes(sheetName)) return;

        // Skip the timestamp column in IT_Storage
        if (sheetName === 'IT_Storage' && range.getColumn() === 9) return;  // Skip column I (timestamps)

        const oldValue = e.oldValue || "";
        const newValue = e.value || "";

        // Skip if no actual change
        if (oldValue === newValue) return;

        // Determine action and color based on sheet and context
        let action = "Edit";
        let color = AUDIT_COLORS.USER_EDIT;

        // NEW: Get item name from IT Storage column D if available
        let itemName = '';
        if (sheetName === 'IT_Storage') {
            const itemRow = range.getRow();
            itemName = sheet.getRange(itemRow, 4).getValue(); // Column D = index 4
            action = "Inventory Edit";
            color = AUDIT_COLORS.USER_EDIT;
        } else {
            action = "Device Edit";
            color = AUDIT_COLORS.DEVICE_CHANGE;
        }

        // Log the edit
        auditLog_writeEntry({
            timestamp: new Date(),
            action: action,
            sheet: sheetName,
            location: range.getA1Notation(),
            user: e.user?.email || Session.getActiveUser().getEmail() || "Unknown",
            oldValue: oldValue,
            newValue: newValue,
            details: `Changed from "${oldValue}" to "${newValue}"`,
            bgColor: color,
            itemName: itemName // NEW: Added item name
        });

    } catch (error) {
        Logger.log('Audit Edit Error: ' + error.message);
    }
}

// Log programmatic stock changes
function auditLog_logStockChange(sheet, cell, oldValue, newValue, action = "Stock Update") {
    try {
        const color = action.includes("Increase") ? AUDIT_COLORS.STOCK_UP : AUDIT_COLORS.STOCK_DOWN;

        // NEW: Get item name from column D for stock changes
        let itemName = '';
        if (sheet.getName() === 'IT_Storage') {
            const itemRow = cell.getRow();
            itemName = sheet.getRange(itemRow, 4).getValue(); // Column D = index 4
        }

        auditLog_writeEntry({
            timestamp: new Date(),
            action: action,
            sheet: sheet.getName(),
            location: cell.getA1Notation(),
            user:  Session.getActiveUser().getEmail() || "Unknown User",  // Actual user email
            oldValue: oldValue,
            newValue: newValue,
            details: `Changed from ${oldValue} to ${newValue}`,
            bgColor: color,
            itemName: itemName // NEW: Added item name
        });
    } catch (e) {
        Logger.log('Audit Stock Error: ' + e.message);
    }
}

// Write entry to TOP of audit log
function auditLog_writeEntry(entry) {
    try {
        const auditSheet = auditLog_createSheet();

        // Insert new row directly below header
        if (auditSheet.getLastRow() > 0) {
            auditSheet.insertRowAfter(1);
        }

        // Write to row 2 (top position)
        // NEW: Added entry.itemName as the last column (index 9)
        const rowPos = 2;
        auditSheet.getRange(rowPos, 1, 1, 9).setValues([[
            Utilities.formatDate(entry.timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            entry.action,
            entry.sheet,
            entry.location,
            entry.user,
            entry.oldValue,
            entry.newValue,
            entry.details,
            entry.itemName || '' // NEW: Item Name column
        ]]);

        // Apply formatting
        const rowRange = auditSheet.getRange(rowPos, 1, 1, 9); // CHANGED: From 8 to 9 columns
        rowRange
            .setBackground(entry.bgColor)
            .setFontSize(10)
            .setVerticalAlignment('top')
            .setWrap(true);

        // Auto-adjust row height
        const lineCount = (entry.details || '').split('\n').length + 1;
        auditSheet.setRowHeight(rowPos, Math.max(24, lineCount * 15));

        // Auto-resize columns
        auditSheet.autoResizeColumns(1, 9); // CHANGED: From 8 to 9 columns

    } catch (error) {
        Logger.log('Audit Write Error: ' + error.message);
    }
}

// Clear audit log content (keep headers)
function auditLog_clear() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Clear Audit Log',
        'Delete ALL audit records? This cannot be undone!',
        ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
        const auditSheet = SpreadsheetApp.getActive().getSheetByName(AUDIT_LOG_SHEET_NAME);
        if (auditSheet && auditSheet.getLastRow() > 1) {
            const lastRow = auditSheet.getLastRow();
            auditSheet.deleteRows(2, lastRow - 1);
            showToast('✅ Audit log cleared', '📋 Audit Log', 3);
        } else {
            showToast('Audit log is empty', '📋 Audit Log', 3);
        }
    }
}

// ... (Rest of the code remains unchanged) ...


// ====================================================================
// ENHANCED STOCK MANAGEMENT WITH AUDIT LOGGING
// ====================================================================

function updateStock(change) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();

        // Validate we're in IT_Storage
        if (sheet.getName() !== 'IT_Storage') {
            showToast('❌ Only change stock in "IT_Storage" tab', '⚠️ Wrong Tab', 5);
            return;
        }

        const cell = sheet.getActiveCell();
        const row = cell.getRow();
        const col = cell.getColumn();

        // Skip header row
        if (row === 1) return;

        // Only apply to stock column (E)
        if (col !== 5) {
            showToast('❌ Select a value in column E (Stock)', '⚠️ Invalid Cell', 5);
            return;
        }

        const stockCell = sheet.getRange(row, 5);
        const timestampCell = sheet.getRange(row, 9);

        const currentValue = stockCell.getValue() || 0;
        const newValue = currentValue + change;

        // Prevent negative stock
        if (newValue < 0) {
            showToast(`❌ Stock cannot be negative: ${newValue}`, '⚠️ Invalid Operation', 5);
            return;
        }

        // === AUDIT LOGGING ===
        const action = change > 0 ? "Stock Increase" : "Stock Decrease";
        auditLog_logStockChange(sheet, stockCell, currentValue, newValue, action);

        // Perform the update
        stockCell.setValue(newValue);

        // Update timestamp
        const now = new Date();
        const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        const user = Session.getActiveUser().getEmail() || "User";
        const sign = change > 0 ? "+" : "";
        const newEntry = `${formattedDate} ${sign}${change} (${user})`;

        // Add to timestamp column
        const prevTimestamp = timestampCell.getValue() || "";
        const timestamps = [newEntry].concat(prevTimestamp.split('\n').filter(t => t));
        timestampCell.setValue(timestamps.slice(0, 3).join('\n')).setWrap(true);

        // Visual feedback
        const color = change > 0 ? "#d0f0c0" : "#f4c2c2";
        const prevColor = stockCell.getBackground();
        stockCell.setBackground(color);
        Utilities.sleep(300);
        stockCell.setBackground(prevColor);

        // Update comment
        const commentText = `Stock ${sign}${change} at ${formattedDate}`;
        const currentComment = stockCell.getComment() || "";
        stockCell.setComment(commentText + (currentComment ? "\n" + currentComment : ""));

        // Success message
        const actionText = change > 0 ? "increased" : "decreased";
        showToast(`✅ Stock ${actionText} by ${Math.abs(change)}. New value: ${newValue}`, '✓ Success', 3);

    } catch (error) {
        showToast(`❌ Stock update failed: ${error.message}`, '⚠️ Error', 5);
        Logger.log(`Stock Error: ${error.message}`);
    }
}

function incrementStock() { updateStock(1); }
function decrementStock() { updateStock(-1); }

// ====================================================================
// UNIVERSAL HELPER FUNCTIONS
// ====================================================================

// Show notification that works in all environments
function showToast(message, title = "Notification", timeout = 5) {
    try {
        SpreadsheetApp.getActive().toast(message, title, timeout);
    } catch (e) {
        // Fallback to dialog if needed
        const ui = SpreadsheetApp.getUi();
        ui.showModelessDialog(
            HtmlService.createHtmlOutput(`
        <div style="padding:15px;font-family:Arial;line-height:1.5">
          <h3 style="margin:0 0 10px">${title}</h3>
          <p>${message}</p>
        </div>
      `).setWidth(400).setHeight(150),
            title
        );

        // Auto-close after timeout
        Utilities.sleep(timeout * 1000);
        ui.hideModelessDialog();
    }
}

// ====================================================================
// INTEGRATED MENU SYSTEM
// ====================================================================

// Add this function to retrieve employee data
function getEmployeeData() {
    try {
        const sheet = SpreadsheetApp.getActive().getSheetByName("EmployeeDB");
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        // Skip header row
        const rows = data.slice(1);

        // Return employee data as objects
        return rows.map(row => ({
            name: row[0] || '',           // Column A: Name
            position: row[1] || '',       // Column B: Position
            department: row[2] || '',     // Column C: Department
            location: row[3] || '',       // Column D: Location
            email: row[4] || '',          // Column E: Email
            mobile: row[5] || ''          // Column F: Mobile
        })).filter(e => e.email);       // Only include entries with email
    } catch (e) {
        Logger.log("Employee data error: " + e);
        return [];
    }
}

// Update your assignDevice function
function assignDevice() {
    const employeeEmails = getEmployeeEmails(); // Reuse existing function
    const template = HtmlService.createTemplateFromFile('mobileAssignDialog');
    template.employeeEmails = employeeEmails;

    const html = template.evaluate()
        .setWidth(500)
        .setHeight(employeeEmails.length > 5 ? 450 : 400);

    SpreadsheetApp.getUi().showModalDialog(html, '📱 Assign Device');
}

// Keep your existing getEmployeeEmails function
function getEmployeeEmails() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const empSheet = ss.getSheetByName("EmployeeDB");
        if (!empSheet) return [];

        const lastRow = empSheet.getLastRow();
        if (lastRow < 2) return [];

        // Column E for email
        return empSheet.getRange(2, 5, lastRow-1, 1)
            .getValues()
            .flat()
            .filter(email => typeof email === 'string' && email.includes('@'))
            .sort();
    } catch (e) {
        console.error("Email Fetch Error:", e);
        return [];
    }
}


function fixAllSettings() {
    try {
        // Set up audit log
        auditLog_initialize();

        // Format main sheet
        const mainSheet = SpreadsheetApp.getActive().getSheetByName('IT_Storage');
        if (mainSheet) {
            const timestampCol = 9; // Column I
            mainSheet.setColumnWidth(timestampCol, 250);
            mainSheet.getRange('A1:Z1').setFontWeight('bold');
        }

        showToast('✅ All settings fixed!', '⚙️ System Check', 5);

    } catch (e) {
        showToast('❌ Settings fix failed: ' + e.message, '⚠️ Error', 5);
    }
}
