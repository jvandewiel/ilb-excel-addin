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
  document.getElementById("generate-ep-sheet").onclick = generateEP;
  
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

  // Enable EP generation button 
  const generateEPButton = document.getElementById("generate-ep-sheet") as HTMLElement;
  if (generateEPButton) {
    generateEPButton.classList.remove('disabled');
  }
}

 /**
  * Trigger Generate EP sheet for the selected partner
  */
export async function generateEP() {
  if (!selectedPartnerId) {
    updateStatus("Please select a partner first", "error");
    return;
  }

  const selectedPartner = partners.find(p => p.id === selectedPartnerId);
  updateStatus(`Generating equal pay sheet for ${selectedPartner?.businessName}...`, "loading");

  try {
    await Excel.run(async (context) => {
      // Call function to generate EP sheet
      await generateEPSheet(context, selectedPartnerId, selectedPartner?.businessName || 'Unknown');
      updateStatus(`Equal pay sheet generated successfully for ${selectedPartner?.businessName}`, "success");
    });
  } catch (error) {
    console.error("Error generating equal pay sheet:", error);
    updateStatus(`Error generating equal pay sheet: ${error.message}`, "error");
  }
}
  
async function generateEPSheet(context: Excel.RequestContext, partnerId: string, partnerName: string) {
  console.log(`Preparing to generate EP sheet for partner ID: ${partnerId}`);
  
  // Get or create EP worksheet and clear if exists
  let worksheet: Excel.Worksheet;
  try {
    worksheet = context.workbook.worksheets.getItem("EP");
    await context.sync();
    console.log("EP worksheet found.");
    worksheet.getUsedRange().clear();
    
  } catch (error) {
    console.log("EP worksheet not found, creating a new one.");
    worksheet = context.workbook.worksheets.add("EP");
    await context.sync();
  }

  worksheet.activate();
  await context.sync();

  console.log(`Generating EP data for partner: ${partnerName}`);
  
  // Load partner data and reference data, display side-by-side in EP sheet
  const epModel = dataLoader.generateEPModel(partnerId);
  
  if (epModel.error) {
    console.error("Error generating EP model:", epModel.error);
    const errorRange = worksheet.getRange("A1:B1");
    errorRange.values = [["Error", epModel.error]];
    errorRange.format.font.color = "red";
    await context.sync();
    return;
  }

  console.log("EP model generated:", epModel);

  // Generate metadata comparison at the top (rows 1-12)
  const metadataComparison = [
    ["Equal Pay Analysis", "", "", "Partner Data", "", "", "Reference Data", "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["Partner Name", "", "", epModel.metadata.partner.businessName, "", "", epModel.metadata.reference.businessName, "", ""],
    ["Commerce Number", "", "", epModel.metadata.partner.commerceNumber, "", "", epModel.metadata.reference.commerceNumber, "", ""],
    ["Start Date", "", "", epModel.metadata.partner.startDate || "N/A", "", "", epModel.metadata.reference.startDate || "N/A", "", ""],
    ["Publish Date", "", "", epModel.metadata.partner.publishDate || "N/A", "", "", epModel.metadata.reference.publishDate || "N/A", "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["Working Hours", "", "", "", "", "", "", "", ""],
    ["Full Time Hours", "", "", `${epModel.workingHours.partner.fullTimeHours} ${epModel.workingHours.partner.fullTimeHoursPer}`, "", "", `${epModel.workingHours.reference.fullTimeHours} ${epModel.workingHours.reference.fullTimeHoursPer}`, "", ""],
    ["Comments", "", "", epModel.workingHours.partner.comments, "", "", epModel.workingHours.reference.comments, "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["Remuneration Components:", "", "", "", "", "", "", "", ""],
  ];

  console.log(metadataComparison);

  const metadataRange = worksheet.getRange("A1:I12");
  metadataRange.values = metadataComparison;
  
  // Format
  worksheet.getRange("A1:I1").format.font.bold = true;
  await context.sync();

  console.log("Metadata comparison added to EP sheet.");

  let currentDataRow = 13; // Start after metadata

  // Component mapping for display order
  const componentOrder = [
    //{ key: 'allowances', title: 'Allowances' },
    { key: 'holidayAllowances', title: 'Holiday Allowances' },
    { key: 'leave', title: 'Leave' },
    { key: 'pensionData', title: 'Pension' },
    { key: 'sickPay', title: 'Sick Pay' },
    { key: 'individualChoiceBudget', title: 'Individual Choice Budget' },
    { key: 'reimbursementsData', title: 'Reimbursements' },
    { key: 'paymentsData', title: 'Payments' },
    { key: 'wageIncrements', title: 'Wage Increments' },
    { key: 'travelExpensesData', title: 'Travel Expenses' }
  ];

  // Process each component
  for (const component of componentOrder) {
    const componentData = epModel[component.key];
    
    if (componentData && (componentData.partner.hasData || componentData.reference.hasData)) {
      currentDataRow = await addEPComponent(
        worksheet,
        component.title,
        componentData.partner,
        componentData.reference,
        currentDataRow
      );
    }
  }
  
  // Set column widths and formatting
  
  worksheet.getRange("A:A").format.columnWidth = 120;
  worksheet.getRange("B:C").format.columnWidth = 60;
  
  worksheet.getRange("D:D").format.columnWidth = 120;

  worksheet.getRange("E:E").format.columnWidth = 120;
  worksheet.getRange("F:G").format.columnWidth = 60;
  
  
  await context.sync();
  console.log(`EP sheet generated successfully for ${partnerName}`);
}

/**
 * Load data into the ILB sheet with proper structure
 */
async function loadDataToILBSheet(context: Excel.RequestContext, partnerId: string, partnerName: string) {
  console.log(`Preparing to load data into ILB sheet for partner ID: ${partnerId}`);
  
  let worksheet: Excel.Worksheet;
  try {
    worksheet = context.workbook.worksheets.getItem("ILB");
    await context.sync();
    console.log("ILB worksheet found, clearing");
    worksheet.getUsedRange().clear();    
  } catch (error) {
    console.log("ILB worksheet not found, creating a new one.");
    worksheet = context.workbook.worksheets.add("ILB");
    // Delete Sheet1 if existing
    context.workbook.worksheets.getItem("Sheet1").delete();
    await context.sync();
  }

  worksheet.activate();
  await context.sync();
  
  // Clear existing named ranges first
  await clearExistingNamedRanges(context);

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
  await context.sync();

  let currentDataRow = 11; // A12 starts at row index 9 (0-based)

  // Load components that are pensionable
  const pensionableComponents = dataLoader.getPensionableComponents(partnerId);
  console.log('Pensionable Components:', pensionableComponents);

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

  // Generate travel expenses data
  const travelExpensesData = dataLoader.generateTravelExpensesData(partnerId);
  if (travelExpensesData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Travel Expenses", travelExpensesData, currentDataRow);
  }

  // Generate payments data (anniversary payments)
  const paymentsData = dataLoader.generatePaymentsData(partnerId);
  if (paymentsData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Payments", paymentsData, currentDataRow);
  }

  // Generate wage increments data
  const wageIncrementsData = dataLoader.generateWageIncrementsData(partnerId);
  if (wageIncrementsData.length > 0) {
    currentDataRow = await addComponent(worksheet, "Wage Increments", wageIncrementsData, currentDataRow);
  }
  
  // Set column widths and word wrap
  worksheet.getRange("A:I").format.columnWidth = 100;
  worksheet.getRange("A:I").format.wrapText = false;
    
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
    
  return currentRow + data.length + 1; // Return next available row with spacing
}

/**
 * Add an EP component with side-by-side comparison layout
 * @param worksheet - The Excel worksheet
 * @param title - The component title
 * @param partnerData - Partner component data
 * @param referenceData - Reference component data
 * @param startRow - The row index to start placing the component (0-based)
 * @returns The next available row index after this component
 */
async function addEPComponent(
  worksheet: Excel.Worksheet,
  title: string,
  partnerData: { hasData: boolean; rows: any[][] },
  referenceData: { hasData: boolean; rows: any[][] },
  startRow: number
): Promise<number> {
  let currentRow = startRow;
  
  console.log(`Adding EP component: ${title} at row ${currentRow}`);

  console.log('Partner Data:', partnerData);
  console.log('Reference Data:', referenceData);

  // Add component title spanning all columns
  // Add component title - only 1 column
  const titleRange = worksheet.getRangeByIndexes(currentRow, 0, 1, 9);
  titleRange.values = [[title, "", "", "", "", "", "", "", ""]];
  titleRange.format.font.bold = true;
  titleRange.format.fill.color = "#ADD8E6"; // Light blue
  setAllBorders(titleRange, false);
  currentRow++;

  
  // Add section headers for Partner and Reference
  // Should make this into a function to add a data block dynamically and return used range for formatting
  // with some headers etc
  // table: title, data, range by index
  const sectionHeaderRange = worksheet.getRangeByIndexes(currentRow, 0, 1, 9);
  sectionHeaderRange.values = [["Partner Data", "", "", "", "Reference Data", "", "", "", ""]];
  sectionHeaderRange.format.font.bold = true;
  sectionHeaderRange.format.fill.color = "#E3F2FD";
  
  // Merge cells for section headers
  //worksheet.getRangeByIndexes(currentRow, 3, 1, 3).merge(); // Partner Data
  //worksheet.getRangeByIndexes(currentRow, 6, 1, 3).merge(); // Reference Data
  
  //setAllBorders(sectionHeaderRange, true);
  currentRow++;
  
  // Determine the maximum number of rows needed
  const partnerRows = partnerData.hasData ? partnerData.rows : [["No data available", "", "", "", "", "", "", "", ""]];
  const referenceRows = referenceData.hasData ? referenceData.rows : [["No data available", "", "", "", "", "", "", "", ""]];
  const maxRows = Math.max(partnerRows.length, referenceRows.length);
  
  // Create side-by-side data layout
  const combinedData: any[][] = [];

  console.log(`Combining data for EP component: ${title}, max rows: ${maxRows}`, combinedData);
  
  for (let i = 0; i < maxRows; i++) {
    const partnerRow = i < partnerRows.length ? partnerRows[i] : ["", "", "", "", "", "", "", "", ""];
    const referenceRow = i < referenceRows.length ? referenceRows[i] : ["", "", "", "", "", "", "", "", ""];
    
    // Take first 3 columns from partner, first 3 columns from reference
    // Layout: [Partner Col1, Partner Col2, Partner Col3, spacer, Reference Col1, Reference Col2, Reference Col3]
    const combinedRow = [
      partnerRow[0] || "", // Partner component name
      partnerRow[1] || "", // Partner amount
      partnerRow[2] || "", // Partner amount type
      "", // Spacer column
      referenceRow[0] || "", // Reference component name
      referenceRow[1] || "", // Reference amount
      referenceRow[2] || "", // Reference amount type
      "", // Extra spacer
      ""  // Extra spacer
    ];
    
    combinedData.push(combinedRow);
  }

  console.log(`Combined data for EP component: ${title}`, combinedData);
  
  // Add the combined data to the worksheet
  if (combinedData.length > 0) {
    const dataRange = worksheet.getRangeByIndexes(currentRow, 0, combinedData.length, combinedData[0].length);
    console.log(`Adding data range for EP component: ${title} at row ${currentRow}`, dataRange);
    dataRange.values = combinedData;
    
    // Format data borders
    //setAllBorders(dataRange, true);
    
    // Format header row (first row of data) as bold if it contains headers
    /*
    if (combinedData.length > 0 && (
      partnerRows[0]?.[0] === "Allowance Name" ||
      partnerRows[0]?.[0] === "Component Name" ||
      partnerRows[0]?.[0] === "Leave Type" ||
      partnerRows[0]?.[0] === "Payment Type" ||
      partnerRows[0]?.[0] === "Reimbursement Name"
    )) {
      const headerRange = worksheet.getRangeByIndexes(currentRow, 0, 1, 9);
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#F0F8FF";
    }
    */

    // Add alternating row colors for better readability
    /*
    for (let i = 1; i < combinedData.length; i += 2) {
      const alternateRange = worksheet.getRangeByIndexes(currentRow + i, 0, 1, 9);
      alternateRange.format.fill.color = "#F8F9FA";
    }
    */
    currentRow += combinedData.length;
  }
  
  // Add spacing after component
  currentRow++;
  
  console.log(`Finished adding EP component: ${title}, next available row is ${currentRow}`);
  
  return currentRow;
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

