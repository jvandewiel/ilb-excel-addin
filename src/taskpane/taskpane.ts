/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { dataLoader, Partner, RemunerationData } from "../services/dataLoaderService";

// Global variables
let partners: Partner[] = [];
let selectedPartnerId: string = "";

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  // Wire up event handlers
  document.getElementById("load-partners").onclick = loadPartners;
  document.getElementById("partner-select").onchange = onPartnerSelect;
  document.getElementById("load-remuneration").onclick = loadRemuneration;
  
  // Load partners on startup
  loadPartners();
});

/**
 * Load partners and populate the dropdown
 */
export async function loadPartners() {
  console.log('Starting to load partners...');
  updateStatus("Loading partners from JSON data...", "loading");
  
  try {
    // Load the data first
    await dataLoader.loadData();
    
    // Get partners from the loaded data
    partners = dataLoader.getPartners();
    console.log('Partners loaded:', partners);
    
    const select = document.getElementById("partner-select") as HTMLSelectElement;
    if (select) {
      // Clear existing options
      select.innerHTML = '<option value="">Select a partner...</option>';
      
      // Add partner options
      partners.forEach(partner => {
        console.log('Adding partner to dropdown:', partner);
        const option = document.createElement('option');
        option.value = partner.id;
        option.textContent = partner.businessName;
        select.appendChild(option);
      });
      
      // Enable the dropdown
      select.disabled = false;
    }
    
    updateStatus(`Loaded ${partners.length} partners`, "success");
  } catch (error) {
    console.error("Error loading partners:", error);
    updateStatus(`Error loading partners: ${error.message}`, "error");
  }
}

/**
 * Handle partner selection
 */
export function onPartnerSelect() {
  const select = document.getElementById("partner-select") as HTMLSelectElement;
  if (select) {
    selectedPartnerId = select.value;
    
    const loadButton = document.getElementById("load-remuneration") as HTMLElement;
    
    if (loadButton) {
      if (selectedPartnerId) {
        loadButton.classList.remove('disabled');
      } else {
        loadButton.classList.add('disabled');
      }
    }
    
    if (selectedPartnerId) {
      const selectedPartner = partners.find(p => p.id === selectedPartnerId);
      updateStatus(`Selected partner: ${selectedPartner?.businessName || selectedPartnerId}`, "success");
    } else {
      updateStatus("Select a partner to continue", "loading");
    }
  }
}

/**
 * Load remuneration data for the selected partner
 */
export async function loadRemuneration() {
  const loadButton = document.getElementById("load-remuneration") as HTMLElement;
  if (loadButton && loadButton.classList.contains('disabled')) {
    return; // Button is disabled, do nothing
  }
  
  if (!selectedPartnerId) {
    updateStatus("Please select a partner first", "error");
    return;
  }
  
  const selectedPartner = partners.find(p => p.id === selectedPartnerId);
  updateStatus(`Loading remuneration and allowances for ${selectedPartner?.businessName}...`, "loading");
  
  try {
    // Get remuneration data for the partner
    const remunerationData: RemunerationData[] = dataLoader.getPartnerRemuneration(selectedPartnerId);
    
    if (!remunerationData || remunerationData.length === 0) {
      updateStatus("No remuneration available for this partner", "error");
      return;
    }

    // Load data into Excel ILB sheet (includes allowances)
    await Excel.run(async (context) => {
      await loadDataToILBSheet(context, selectedPartnerId, selectedPartner?.businessName || 'Unknown');
      updateStatus(`Remuneration and allowances loaded successfully for ${selectedPartner?.businessName} (${remunerationData.length} records)`, "success");
    });
    
  } catch (error) {
    console.error("Error loading remuneration:", error);
    updateStatus(`Error loading remuneration: ${error.message}`, "error");
  }
}

/**
 * Load data into the ILB sheet with proper structure
 */
async function loadDataToILBSheet(context: Excel.RequestContext, partnerId: string, partnerName: string) {
  console.log(`Preparing to load data into ILB sheet for partner ID: ${partnerId}`);
  
  // Get or create ILB worksheet using a more reliable approach
  let worksheet: Excel.Worksheet;
  
  // Load all worksheets to check if ILB exists
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();
  
  // Check if ILB worksheet already exists
  const existingWorksheet = worksheets.items.find(ws => ws.name === "ILB");
  
  if (existingWorksheet) {
    worksheet = existingWorksheet;
    //console.log(`Found existing worksheet: ${worksheet.name}`);
  } else {
    console.log("ILB worksheet not found, creating a new one.");
    worksheet = worksheets.add("ILB");
    //console.log(`Created new worksheet: ${worksheet.name}`);
    await context.sync();
  }

  console.log("Clearing existing data in ILB sheet...");

  // Clear existing named ranges first
  await clearExistingNamedRanges(context);

  // Clear the entire sheet
  const usedRange = worksheet.getUsedRange();
  usedRange.load("address");
  await context.sync();
  
  if (usedRange.address !== "A1:A1" || usedRange.values[0][0] !== "") {
    usedRange.clear();
    await context.sync();
  }

  console.log(`Loading data for partner: ${partnerName} into ILB sheet`);
  
  // Generate metadata (rows 1-10)
  const metadata = dataLoader.generateMetadata(partnerId); 
  console.log('Metadata generated:', metadata);

  const metadataRange = worksheet.getRange("A1:C10"); // Dynamically set range for metadata
  metadataRange.values = metadata;
  
  // Format metadata section
  metadataRange.format.font.bold = true;
  metadataRange.getColumn(0).format.fill.color = "#E3F2FD";
  setAllBorders(metadataRange, false);

  // Set named ranges using the workbook's named items collection
  const workingHoursCell = worksheet.getRange("B7");
  const namedItem = context.workbook.names.add("WorkingHoursPerWeek", workingHoursCell);

  await context.sync();

  let currentDataRow = 11; // A12 starts at row index 9 (0-based)

  
  // Generate holiday allowances data (starting from A10)  
  const holidayAllowancesData = dataLoader.generateHolidayAllowances(partnerId);
  console.log('Holiday Allowances Data:', holidayAllowancesData);
  if (holidayAllowancesData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Holiday Allowances", holidayAllowancesData, currentDataRow);
  }

  // Generate leave data
  const leaveData = dataLoader.generateLeaveData(partnerId);
  if (leaveData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Leave", leaveData, currentDataRow);
  }

  // Generate individual choice budget data
  const individualChoiceBudgetData = dataLoader.generateIndividualChoiceBudget(partnerId);
  if (individualChoiceBudgetData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Individual Choice Budget", individualChoiceBudgetData, currentDataRow);
  }

  // Generate pension data
  const pensionData = dataLoader.generatePensionData(partnerId);
  if (pensionData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Pension", pensionData, currentDataRow);
  }

    // Generate sick pay data
  const sickPayData = dataLoader.generateSickPayData(partnerId);
  if (sickPayData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Sick Pay", sickPayData, currentDataRow);
  }
  
  // Generate allowances data
  const allowancesData = dataLoader.generateAllowances(partnerId); 
  if (allowancesData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Allowances", allowancesData, currentDataRow);
  }

  // Generate reimbursements data
  const reimbursementsData = dataLoader.generateReimbursementsData(partnerId);
  if (reimbursementsData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Reimbursements", reimbursementsData, currentDataRow);
  }

  // Generate payments data (anniversary payments)
  const paymentsData = dataLoader.generatePaymentsData(partnerId);
  if (paymentsData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Payments", paymentsData, currentDataRow);
  }
  
  // Auto-fit columns
  worksheet.getUsedRange().format.autofitColumns();
  
  // Set specific column widths and word wrap
  const columnB = worksheet.getRange("B:B");
  columnB.format.columnWidth = 100;
  const cellB8 = worksheet.getCell(7, 1); // B8 (0-based indexing: row 7, col 1)
  cellB8.format.wrapText = false;
  const columnCF = worksheet.getRange("C:F");
  columnCF.format.columnWidth = 100;
  
  await context.sync();
  console.log(`Data loaded successfully into ILB sheet for ${partnerName}`);
}

/**
 * Clears all existing named ranges in the workbook
 * @param context - The Excel request context
 */
async function clearExistingNamedRanges(context: Excel.RequestContext) {
  console.log("Clearing existing named ranges...");
  
  const namedItems = context.workbook.names;
  namedItems.load("items");
  await context.sync();
  
  // Remove all existing named ranges
  const itemsToDelete = namedItems.items.slice(); // Create a copy to avoid modification during iteration
  for (const item of itemsToDelete) {
    console.log(`Deleting named range: ${item.name}`);
    item.delete();
  }
  
  if (itemsToDelete.length > 0) {
    await context.sync();
    console.log(`Cleared ${itemsToDelete.length} existing named ranges`);
  } else {
    console.log("No existing named ranges to clear");
  }
}

/**
 * Sets all borders (edge and inside) for a given range
 * @param range - The Excel range to apply borders to
 * @param includeInside - Whether to include inside borders (default: true)
 */
function setAllBorders(range: Excel.Range, includeInside: boolean = true) {  
  range.format.borders.getItem("EdgeTop").style = "Continuous";
  range.format.borders.getItem("EdgeBottom").style = "Continuous";
  range.format.borders.getItem("EdgeLeft").style = "Continuous";
  range.format.borders.getItem("EdgeRight").style = "Continuous";
  
  if (includeInside) {
    range.format.borders.getItem("InsideVertical").style = "Continuous";
    range.format.borders.getItem("InsideHorizontal").style = "Continuous";
  }
}

/**
 * Generic function to add an ILB component table to Excel with consistent formatting
 * @param worksheet - The Excel worksheet
 * @param title - The component title (e.g., "Allowances", "Holiday Allowances")
 * @param data - The data array with headers as first row
 * @param startRow - The row index to start placing the component (0-based)
 * @returns The next available row index after this component
 */
async function addComponent(
  worksheet: Excel.Worksheet, 
  title: string, 
  data: any[][], 
  startRow: number
): Promise<number> {
  if (!data || data.length === 0) {
    return startRow;
  }

  let currentRow = startRow;
   
  // Add component title - only 1 column
  const titleRange = worksheet.getRangeByIndexes(currentRow, 0, 1, 1);
  titleRange.values = [[title]];
  titleRange.format.font.bold = true;
  titleRange.format.fill.color = "#ADD8E6"; // Light blue
  setAllBorders(titleRange, false);
  currentRow++;
  
  // Add data (including headers)
  const dataEndCol = Math.max(0, Math.max(...data.map(row => row.length)) - 1);
  const dataRange = worksheet.getRangeByIndexes(currentRow, 0, data.length, dataEndCol + 1);
  
  dataRange.values = data;
  
  // Format all data cells with borders
  setAllBorders(dataRange);
  
  // Format header row (first row of data) as bold
  if (data.length > 0) {
    const headerRange = worksheet.getRangeByIndexes(currentRow, 0, 1, dataEndCol + 1);
    headerRange.format.font.bold = true;
  }
  
  // Special formatting for description and conditions columns (last two columns if they exist)
  if (dataEndCol >= 6) { // Only if we have at least 7 columns (including description and conditions)
    // Format conditions column (2nd to last)
    const conditionsCol = dataEndCol - 1;
    const conditionsRange = worksheet.getRangeByIndexes(currentRow, conditionsCol, data.length, 1);
    conditionsRange.format.wrapText = true;
    conditionsRange.format.columnWidth = 400;
    conditionsRange.format.verticalAlignment = "Top";
    
    // Format description column (last column) 
    const descriptionCol = dataEndCol;
    const descriptionRange = worksheet.getRangeByIndexes(currentRow, descriptionCol, data.length, 1);
    descriptionRange.format.wrapText = true;
    descriptionRange.format.columnWidth = 400;
    descriptionRange.format.verticalAlignment = "Top";
  }
  
  // Auto-fit other columns
  for (let col = 0; col <= Math.min(dataEndCol - 2, dataEndCol); col++) {
    const colRange = worksheet.getRangeByIndexes(currentRow, col, data.length, 1);
    colRange.format.autofitColumns();
    colRange.format.verticalAlignment = "Top";
  }
  
  return currentRow + data.length + 1; // Return next available row with spacing
}

/**
 * Updates the status message
 */
function updateStatus(message: string, type: "loading" | "success" | "error") {
  const statusDiv = document.getElementById("status");
  if (statusDiv) {
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
  }
}

