/*
 * Data Loader Service for Inlenersbeloning Excel Plugin
 * Loads remuneration data from local JSON file
 */

export interface Partner {
  id: string;
  businessName: string;
}

export interface Company {
  id: string;
  businessName: string;
  commerceNumber: string;
  type: string;
  legalForm: string;
  sbiCodes: string[];
  employees: number | null;
  phoneNumber: string;
  addresses: Address[];
}

export interface Address {
  type: string;
  street: string;
  houseNumber: number;
  houseNumberAddition: string | null;
  postalCode: string;
  city: string;
  country: string;
}

export interface FunctionGroup {
  id: string;
  code: string;
  codeID: number;
  tableType: string;
  tableTypeID: number;
  hasHourRateCalculationMethods: boolean;
  hasCustomHourRateCalculationMethods: boolean;
  steps: FunctionStep[];
}

export interface FunctionStep {
  index: number;
  stepTitle: string;
  otherStepTitle: string | null;
  salary: number;
  hourWage: number;
  functionGroupStepID: number;
  ageStepID: number | null;
}

export interface WorkingHours {
  fullTimeHours: number;
  fullTimeHoursPer: string;
  fullTimeHoursPerID: number;
  workingHoursComments: string | null;
  wHRRegulationID: number;
  singleRegulationTypeID: number;
  whrApplytoPayTypesID: number;
}

export interface RemunerationData {
  id: string;
  title: string;
  intermediaryID: string | null;
  startDate: string;
  publishDate: string;
  lenderVersion: number;
  totalVersions: number;
  company: Company;
  workingHours: WorkingHours;
  functionGroups: FunctionGroup[];
  allowances?: AllowancesContainer;
  holidayAllowances?: HolidayAllowance[];
  pensionData?: PensionData;
  leave?: LeaveData;
  sickPay?: SickPayData;
  individualChoiceBudget?: IndividualChoiceBudget;
  reimbursementsData?: ReimbursementsData;
  paymentsData?: PaymentsData;
  // Add other fields as needed
  hasProprietaryRemuneration?: boolean;
  hasFunctionExists?: boolean;
  hasComparableFunction?: boolean;
  hasFunctionRemuneration?: boolean;
  hasRemunerationExists?: boolean;
}

export interface AllowancesContainer {
  id: string;
  hasAllowances: boolean;
  allowances: Allowance[];
}

export interface Allowance {
  id: string;
  typeTitle: string;
  typeCode: string;
  typeOther: string | null;
  name: string;
  conditions: string | null;
  phaseOutSchema: string | null;
  allowanceComponents: AllowanceComponent[];
}

export interface AllowanceComponent {
  id: string;
  isDescribedAsAmount: boolean;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  basedOnOther: string | null;
  intervalAmount: number;
  interval: string;
  hasTimelyConditions?: boolean;
  description?: string;
  conditions?: string;
  // Add other fields as needed for completeness
}

export interface HolidayAllowance {
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  interval: string;
  basedOn: string;
  intervalAmount: number;
  isPartOfIndividualChoiceBudget?: boolean;
  description?: string | null;
  appliesToAllFunctions?: boolean;
  dependencyFunctions?: any[];
}

export interface PensionData {
  id: string;
  hasPension: boolean;
  hasAdditionalPensionArrangements: boolean;
  items: PensionItem[];
}

export interface PensionItem {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  interval: string;
}

export interface LeaveData {
  id: string;
  hasWorkingHoursReduction: boolean;
  hasVacationDays: boolean;
  hasSpecialLeave: boolean;
  hasAdditionalWAZO: boolean;
  hasMandatoryLeave: boolean;
  hasHoliday: boolean;
  holidayCount: number;
  workingHoursReductions: LeaveComponent[];
  vacationDays: LeaveComponent[];
  specialLeave: LeaveComponent[];
  additionalWAZO: LeaveComponent[];
  mandatoryLeave: LeaveComponent[];
}

export interface LeaveComponent {
  id: string;
  targetGroup?: {
    value: string | null;
    otherValue: string | null;
  };
  amount: number;
  amountType: {
    value: string | null;
    otherValue: string | null;
  };
  intervalAmount: number;
  interval: string;
  description?: string;
  isPartOfIndividualChoiceBudget?: boolean;
  conditions?: string;
}

export interface SickPayData {
  id: string;
  sickPayCondition: string | null;
  hasWaitingDays: boolean;
  amount: number;
  sickPayAmounts: SickPayAmount[];
}

export interface SickPayAmount {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  sickPayYear: {
    value: string;
  };
  basedOn: string;
  intervalAmount: number;
  interval: string;
}

export interface IndividualChoiceBudget {
  id: string;
  hasIndividualChoiceBudget: boolean;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  intervalAmount: number;
  interval: string;
  items: IndividualChoiceBudgetItem[];
}

export interface IndividualChoiceBudgetItem {
  id: string;
  name: string;
  typeOther: string | null;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  intervalAmount: number;
  interval: string;
  description: string | null;
  isPartOfIndividualChoiceBudget?: boolean;
  appliesToAllFunctions: boolean;
  dependencyFunctions: any[];
}

export interface ReimbursementsData {
  id: string;
  hasReimbursements: boolean;
  reimbursements: Reimbursement[];
}

export interface Reimbursement {
  id: string;
  name: string;
  typeCode: string;
  isDescribedAsAmount: boolean;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  intervalAmount: number;
  interval: string;
  description?: string;
  conditions?: string;
}

export interface PaymentsData {
  id: string;
  hasAdditionalSocialSecurityArrangements: boolean;
  hasAnniversaryPayments: boolean;
  hasFixedPayments: boolean;
  hasOneTimePayments: boolean;
  hasVariablePayments: boolean;
  additionalSecurityArrangements: any[];
  anniversaryPayments: AnniversaryPayment[];
  fixedPayments: any[];
  oneTimePayments: any[];
  variablePayments: any[];
}

export interface AnniversaryPayment {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  amountOfYears: number;
  basedOn: string;
  basedOnOther?: string | null;
  intervalAmount: number;
  interval: string;
  date?: string | null;
  dateOrCondition: string;
  dateOrConditionOther?: string | null;
  description: string;
  appliesToAllFunctions: boolean;
  isProportional: boolean;
  proportionalBasedOn: {
    value: string | null;
  };
  proportionalDescription?: string | null;
  dependencyFunctions: any[];
}

/**
 * Data loader service to load remuneration data from JSON file
 */
export class DataLoaderService {
  private remunerationData: RemunerationData[] = [];
  private isLoaded = false;

  /**
   * Load the remuneration data from JSON file
   */
  async loadData(): Promise<void> {
    if (this.isLoaded) {
      return;
    }

    try {
      console.log("Loading remuneration data from JSON file...");

      const response = await fetch("./assets/renumerarions-export.json");

      if (!response.ok) {
        throw new Error(`Failed to load data: ${response.status} ${response.statusText}`);
      }

      this.remunerationData = await response.json();
      this.isLoaded = true;

      console.log(`Loaded ${this.remunerationData.length} remuneration records`);
    } catch (error) {
      console.error("Error loading remuneration data:", error);
      throw new Error(`Failed to load remuneration data: ${error.message}`);
    }
  }

  /**
   * Get all unique partners from the data
   */
  getPartners(): Partner[] {
    if (!this.isLoaded) {
      console.warn("Data not loaded yet. Call loadData() first.");
      return [];
    }

    const partnersMap = new Map<string, Partner>();

    this.remunerationData.forEach((item) => {
      if (item.company && item.company.id && item.company.businessName) {
        partnersMap.set(item.company.id, {
          id: item.company.id,
          businessName: item.company.businessName,
        });
      }
    });

    const partners = Array.from(partnersMap.values());
    console.log(`Found ${partners.length} unique partners`);

    return partners.sort((a, b) => a.businessName.localeCompare(b.businessName));
  }

  /**
   * Get remuneration data for a specific partner
   */
  getPartnerRemuneration(partnerId: string): RemunerationData[] {
    if (!this.isLoaded) {
      console.warn("Data not loaded yet. Call loadData() first.");
      return [];
    }

    const partnerData = this.remunerationData.filter(
      (item) => item.company && item.company.id === partnerId
    );

    console.log(`Found ${partnerData.length} remuneration records for partner: ${partnerId}`);

    return partnerData;
  }

  /**
   * Generate metadata for Excel sheet (rows 1-10)
   */
  generateMetadata(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);

    if (!partnerData || partnerData.length === 0) {
      return [
        ["Metadata", "", ""],
        ["Partner Name", "No data available", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
      ];
    }

    const firstRecord = partnerData[0];
    const company = firstRecord.company;
    const workingHours = firstRecord.workingHours;

    return [
      ["ILB Remuneration Data", "", ""],
      ["Partner Name", company?.businessName || "Unknown", ""],
      ["Commerce Number", company?.commerceNumber || "N/A", ""],
      ["Version", `${firstRecord.lenderVersion}/${firstRecord.totalVersions}`, ""],
      ["Start Date", firstRecord.startDate || "N/A", ""],
      ["Publish Date", firstRecord.publishDate || "N/A", ""],
      [
        "Full Time Hours",
        workingHours?.fullTimeHours || "N/A",
        workingHours?.fullTimeHoursPer || "",
      ],
      ["Working Hours Comments", workingHours?.workingHoursComments || "N/A", ""],
      ["", "", ""],
      ["Remuneration Components:", "", ""],
    ];
  }

  /**
   * Generate holiday allowances data for Excel format
   * Takes holidayAllowances objects from remuneration data
   * Returns formatted data with holiday allowance details
   */
  generateHolidayAllowances(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No holiday allowances data available", "", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Allowance Name",
      "Amount",
      "Amount Type",
      "Interval",
      "Based On",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.holidayAllowances && record.holidayAllowances.length > 0) {
        record.holidayAllowances.forEach((holidayAllowance) => {
          const amountDisplay = holidayAllowance.amount
            ? `${holidayAllowance.amount}${holidayAllowance.amountType?.value === "Percentage" ? "%" : ""}`
            : "N/A";

          result.push([
            "Holiday Allowance",
            amountDisplay,
            holidayAllowance.amountType?.value || "N/A",
            holidayAllowance.interval || "N/A",
            holidayAllowance.basedOn || "N/A",
            (holidayAllowance as any).isPartOfIndividualChoiceBudget ? "true" : "false",
            "N/A",
            (holidayAllowance as any).description || "N/A",
          ]);
        });
      }
    });

    // If no holiday allowances found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate allowances data for Excel format
   * Takes allowances.allowances objects from remuneration data
   * Returns formatted data with allowance name first, then component details in following columns
   */
  generateAllowances(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No allowances data available", "", "", "", "", "", "", ""];

    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add main header with allowance name first, then component fields
    result.push([
      "Allowance Name",
      "Amount",
      "Amount Type",
      "Based On",
      "Interval",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.allowances && record.allowances.hasAllowances && record.allowances.allowances) {
        // Process each allowance
        record.allowances.allowances.forEach((allowance) => {
          const allowanceName = allowance.name || allowance.typeTitle || "Unnamed Allowance";

          // Process allowance components (lowest level)
          if (allowance.allowanceComponents && allowance.allowanceComponents.length > 0) {
            allowance.allowanceComponents.forEach((component) => {
              const amountDisplay = component.amount
                ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
                : "N/A";

              result.push([
                allowanceName,
                amountDisplay,
                component.amountType?.value || "N/A",
                component.basedOn || "N/A",
                component.interval || "N/A",
                "false", // Allowances are typically not part of ICB
                component.conditions || "N/A",
                component.description || "N/A",
              ]);
            });
          } else {
            // No components, show allowance with no component data
            result.push([
              allowanceName,
              "N/A",
              "N/A",
              "N/A",
              "N/A",
              "false",
              allowance.conditions || "N/A",
              "No components available",
            ]);
          }
        });
      }
    });

    // If no allowances found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    console.log(result);

    return result;
  }

  /**
   * Generate pension data for Excel format
   * Similar to holiday allowances - simple table format
   */
  generatePensionData(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No pension data available", "", "", "", "", "", "", ""];

    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Component Name",
      "Amount",
      "Amount Type",
      "Based On",
      "Interval",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.pensionData && record.pensionData.hasPension && record.pensionData.items) {
        record.pensionData.items.forEach((item) => {
          const amountDisplay = item.amount
            ? `${item.amount}${item.amountType?.value === "Percentage" ? "%" : ""}`
            : "N/A";

          result.push([
            "Pension",
            amountDisplay,
            item.amountType?.value || "N/A",
            item.basedOn || "N/A",
            item.interval || "N/A",
            "false", // Pension typically not part of ICB
            "N/A",
            (item as any).description || "N/A",
          ]);
        });
      }
    });

    // If no pension data found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate leave data for Excel format
   * Each subcomponent treated as a new row like allowances
   */
  generateLeaveData(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No leave data available", "", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Leave Type",
      "Amount",
      "Amount Type",
      "Interval",
      "Based On",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.leave) {
        const leave = record.leave;

        // Process working hours reductions
        if (leave.hasWorkingHoursReduction && leave.workingHoursReductions) {
          leave.workingHoursReductions.forEach((component) => {
            const amountDisplay = component.amount
              ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            result.push([
              "Working Hours Reduction",
              //component.targetGroup?.value || "N/A",
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "",
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "false",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        }

        // Process vacation days
        if (leave.hasVacationDays && leave.vacationDays) {
          leave.vacationDays.forEach((component) => {
            const amountDisplay = component.amount
              ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            result.push([
              "Vacation Days",
              //component.targetGroup?.value || "N/A",
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "",
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "false",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        }

        // Process special leave
        if (leave.hasSpecialLeave && leave.specialLeave) {
          leave.specialLeave.forEach((component) => {
            const amountDisplay = component.amount
              ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            result.push([
              "Special Leave",
              //component.targetGroup?.value || "N/A",
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "",
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "false",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        }

        // Process additional WAZO
        if (leave.hasAdditionalWAZO && leave.additionalWAZO) {
          leave.additionalWAZO.forEach((component) => {
            const amountDisplay = component.amount
              ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            result.push([
              "Additional WAZO",
              //component.targetGroup?.value || "N/A",
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "",
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "false",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        }

        // Process mandatory leave
        if (leave.hasMandatoryLeave && leave.mandatoryLeave) {
          leave.mandatoryLeave.forEach((component) => {
            const amountDisplay = component.amount
              ? `${component.amount}${component.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            result.push([
              "Mandatory Leave",
              //component.targetGroup?.value || "N/A",
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "",
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "false",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        }
      }
    });

    // If no leave data found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate sick pay data for Excel format
   * Each sick pay amount treated as a new row like allowances
   */
  generateSickPayData(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No sick pay data available", "", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Component Name",
      "Amount",
      "Amount Type",
      "Based On", 
      "Interval",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.sickPay && record.sickPay.sickPayAmounts && Array.isArray(record.sickPay.sickPayAmounts) && record.sickPay.sickPayAmounts.length > 0) {
        record.sickPay.sickPayAmounts.forEach((sickPayAmount) => {
          const amountDisplay = sickPayAmount.amount
            ? `${sickPayAmount.amount}${sickPayAmount.amountType?.value === "Percentage" ? "%" : ""}`
            : "N/A";

          result.push([
            `Sick Pay (${sickPayAmount.sickPayYear?.value || "Unknown Year"})`,
            amountDisplay,
            sickPayAmount.amountType?.value || "N/A",
            sickPayAmount.basedOn || "N/A",
            sickPayAmount.interval || "N/A",
            "false", // Sick pay typically not part of ICB
            "N/A",
            "N/A",
          ]);
        });
      }
    });

    // If no sick pay data found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate individual choice budget data for Excel format
   * Includes items and ICB column
   */
  generateIndividualChoiceBudget(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No individual choice budget data available", "", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Component Name",
      "Amount",
      "Amount Type",
      "Based On",
      "Interval",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (
        record.individualChoiceBudget &&
        record.individualChoiceBudget.hasIndividualChoiceBudget
      ) {
        const budget = record.individualChoiceBudget;

        // Add main budget entry
        const amountDisplay = budget.amount
          ? `${budget.amount}${budget.amountType?.value === "Percentage" ? "%" : ""}`
          : "N/A";

        result.push([
          "Individual Choice Budget",
          amountDisplay,
          budget.amountType?.value || "N/A",
          budget.basedOn || "N/A",
          budget.interval || "N/A",
          "true", // Main budget is always part of ICB
          "N/A",
          "Main budget allocation",
        ]);

        // Add items if they exist
        if (budget.items && budget.items.length > 0) {
          budget.items.forEach((item) => {
            const itemAmountDisplay = item.amount
              ? `${item.amount}${item.amountType?.value === "Percentage" ? "%" : ""}`
              : "N/A";

            const componentName =
              item.name === "Other" && item.typeOther
                ? item.typeOther
                : item.name || "Unnamed Component";

            result.push([
              componentName,
              itemAmountDisplay,
              item.amountType?.value || "N/A",
              item.basedOn || "N/A",
              item.interval || "N/A",
              item.isPartOfIndividualChoiceBudget ? "true" : "false",
              "N/A", // No conditions field in the data structure
              item.description || "N/A",
            ]);
          });
        }
      }
    });

    // If no individual choice budget found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate reimbursements data for Excel format
   * Each reimbursement treated as a new row like allowances
   */
  generateReimbursementsData(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = ["No reimbursements data available","", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push([
      "Reimbursement Name",
      "Amount",
      "Amount Type",
      "Based On",
      "Interval",
      "ICB",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (
        record.reimbursementsData &&
        record.reimbursementsData.hasReimbursements &&
        record.reimbursementsData.reimbursements
      ) {
        record.reimbursementsData.reimbursements.forEach((reimbursement) => {
          const amountDisplay = reimbursement.amount
            ? `${reimbursement.amount}${reimbursement.amountType?.value === "Percentage" ? "%" : ""}`
            : "N/A";

          result.push([
            reimbursement.name || "Unnamed Reimbursement",
            amountDisplay,
            reimbursement.amountType?.value || "N/A",
            reimbursement.basedOn || "N/A",
            reimbursement.interval || "N/A",
            "false", // Reimbursements typically not part of ICB
            reimbursement.conditions || "N/A",
            reimbursement.description || "N/A",
          ]);
        });
      }
    });

    // If no reimbursements found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }

  /**
   * Generate payments data for Excel format
   * Each anniversary payment treated as a new row like allowances
   */
  generatePaymentsData(partnerId: string): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);

    if (!partnerData || partnerData.length === 0) {
      return [["No payments data available", "", "", "", "", "", "", ""]];
    }

    const result: any[][] = [];
    const noDataFound = ["No payments data found", "", "", "", "", "", "", ""];

    // Add header
    result.push(["Payment Type", "Amount", "Amount Type", "Based On", "Interval", "ICB", "Conditions", "Description"]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.paymentsData && record.paymentsData.hasAnniversaryPayments && record.paymentsData.anniversaryPayments) {
        record.paymentsData.anniversaryPayments.forEach((payment) => {
          const amountDisplay = payment.amount
            ? `${payment.amount}${payment.amountType?.value === "Percentage" ? "%" : ""}`
            : "N/A";

          const paymentName = `Anniversary Payment (${payment.amountOfYears} years)`;
          const basedOnText = payment.basedOn === "Other" && payment.basedOnOther 
            ? payment.basedOnOther 
            : payment.basedOn || "N/A";
          
          const conditionsText = payment.dateOrCondition === "Other" && payment.dateOrConditionOther
            ? payment.dateOrConditionOther
            : payment.dateOrCondition || "N/A";

          result.push([
            paymentName,
            amountDisplay,
            payment.amountType?.value || "N/A",
            basedOnText,
            payment.interval || "N/A",
            "false", // Anniversary payments typically not part of ICB
            conditionsText,
            payment.description || "N/A",
          ]);
        });
      }
    });

    // If no payments found, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
  }
}

// Create a singleton instance
export const dataLoader = new DataLoaderService();
