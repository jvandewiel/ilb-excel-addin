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
  wageIncrements?: WageIncrementsData;
  travelExpensesData?: TravelExpensesData;
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

export interface PensionableEarning {
  title: string;
  code: string;
  component?: string;
}

export interface PensionItem {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  intervalAmount: number;
  interval: string;
  hasFranchise?: boolean;
  franchise?: string;
  description?: string | null;
  appliesToAllFunctions?: boolean;
  dependencyFunctions?: any[];
  pensionableEarnings?: PensionableEarning[];
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
  fixedPayments: FixedPayment[];
  oneTimePayments: OneTimePayment[];
  variablePayments: VariablePayment[];
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

export interface FixedPayment {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  basedOnOther?: string | null;
  intervalAmount: number;
  interval: string;
  date: string;
  description: string;
  isPartOfIndividualChoiceBudget: boolean;
  appliesToAllFunctions: boolean;
  isProportional: boolean;
  proportionalBasedOn: {
    value: string | null;
  };
  proportionalDescription?: string | null;
  dependencyFunctions: any[];
}

export interface OneTimePayment {
  id: string;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  basedOn: string;
  basedOnOther?: string | null;
  intervalAmount: number;
  interval: string;
  date?: string | null;
  description: string;
  isPartOfIndividualChoiceBudget: boolean;
  appliesToAllFunctions: boolean;
  isProportional: boolean;
  proportionalBasedOn: {
    value: string | null;
  };
  proportionalDescription?: string | null;
  dependencyFunctions: any[];
}

export interface VariablePayment {
  id: string;
  hasMinMaxAmount: boolean;
  amount: number;
  minAmount?: number | null;
  maxAmount?: number | null;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  name: string;
  typeOther?: string | null;
  basedOn: string;
  basedOnOther?: string | null;
  intervalAmount: number;
  interval: string;
  date: string;
  hasApplicableDate: boolean;
  applicableDateTo?: string | null;
  applicableDateFrom?: string | null;
  description: string;
  appliesToAllFunctions: boolean;
  isProportional: boolean;
  proportionalBasedOn: {
    value: string | null;
  };
  proportionalDescription?: string | null;
  dependencyFunctions: any[];
}

export interface WageIncrementsData {
  hasWageIncrements: boolean;
  hasRetroactiveWageIncrements: boolean;
  hasYouthScales: boolean;
  youthApplication?: string | null;
  youthApplicationID?: number;
  otherYouthApplication?: string | null;
  youthScalesDisclaimer?: string | null;
  isEligibleForIncrement: boolean;
  incrementData: IncrementData[];
  retroactiveIncrementData: IncrementData[];
  periodicalData: PeriodicalData[];
}

export interface IncrementData {
  incrementType: string;
  incrementTypeID: number;
  percentage: number;
  date: string;
  explanation?: string | null;
  hasWageIncrementsApplied: boolean;
  appliesToAllFunctions: boolean;
  linkedFunctions: any[];
}

export interface PeriodicalData {
  eligibleType: string;
  eligibleTypeID: number;
  experience: string;
  experienceID: number;
  experienceDate?: string | null;
  experienceAlwaysApply?: string | null;
  experienceAlwaysApplyID?: number;
  experienceCriteria?: string | null;
  experienceCriteriaID?: number;
  experienceMinimumAmount?: string | null;
  experienceMinimumAmountType?: string | null;
  experienceMinimumAmountTypeID?: number;
  experienceStartType?: string | null;
  experienceStartTypeID?: number;
  experienceStartDate?: string | null;
  experienceStartYear?: string | null;
  experienceStartYearID?: number;
  experienceOther?: string | null;
  experiencePeriodAmount: string;
  experiencePeriodAmountType: string;
  experiencePeriodAmountTypeID: number;
  performanceMoment: string;
  performanceMomentID: number;
  performanceDate?: string | null;
  performanceIndividual?: string | null;
  performanceOther?: string | null;
  periodicals: string;
  appliesToAllFunctions: boolean;
  linkedFunctions: any[];
}

export interface TravelExpensesData {
  id: string;
  hasTravelExpenses: boolean;
  reimbursementAsWorkingHours: {
    value: string | null;
  };
  travelExpenses: TravelExpense[];
}

export interface TravelExpense {
  id: string;
  compensationType: {
    value: string;
  };
  components: TravelExpenseComponent[];
}

export interface TravelExpenseComponent {
  id: string;
  isDescribedAsAmount: boolean;
  amount: number;
  amountType: {
    value: string;
    otherValue: string | null;
  };
  intervalAmount: number;
  interval: string;
  description: string;
  conditions?: string | null;
  appliesToAllFunctions: boolean;
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

      const response = await fetch("./assets/renumerations-export.json");

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
   * Utility function to format amount with percentage symbol if needed
   * @param amount The amount value
   * @param amountType Object containing amount type information
   * @returns Formatted amount string
   */
  private formatAmountDisplay(amount: number | undefined, amountType: { value: string; otherValue: string | null; } | undefined): string {
    if (!amount) return "N/A";
    return `${amount}${amountType?.value === "Percentage" ? "%" : ""}`;
  }

  /**
   * Maps pensionable earning code values to component properties
   * @param code The code value from pensionable earning
   * @returns The corresponding component name
   */
  private mapCodeToComponent(code: string): string {
    switch (code) {
      case "CAOAllowances":
        return "allowances";
      case "CAOHolidayAllowance":
        return "holidayAllowance";
      case "CAOOneTimePayment":
        return "oneTimePayment";
      case "CAOFixedPayment":
        return "fixedPayment";
      case "CAOVariablePayment":
        return "variablePayment";
      case "CAOHoliday":
        return "holiday";
      case "CAOWageTable":
        return "wageTable";
      case "CAOWageIncrement":
        return "wageIncrement";
      case "CAOPeriodicIncrement":
        return "periodicIncrement";
      case "CAORetroactiveIncrement":
        return "retroactiveIncrement";
      case "CAOTravelExpenses":
        return "travelExpenses";
      case "CAOTravelTime":
        return "travelTime";
      case "Other":
        return "other";
      default:
        return "unknown";
    }
  }

  /**
   * Utility function to enhance pensionable earnings with component mapping
   * @param pensionableEarnings Array of pensionable earnings
   * @returns Array of pensionable earnings with component property added
   */
  private enhancePensionableEarningsWithComponents(pensionableEarnings: PensionableEarning[]): PensionableEarning[] {
    return pensionableEarnings.map(earning => ({
      ...earning,
      component: this.mapCodeToComponent(earning.code)
    }));
  }

  /**
   * Generic function to generate Excel data with consistent structure
   * @param partnerId Partner ID to filter data
   * @param noDataMessage Message to show when no data is found
   * @param headers Array of column headers
   * @param dataProcessor Function that processes partner data and returns rows
   */
  private generateDataWithStandardStructure(
    partnerId: string,
    noDataMessage: string,
    headers: string[],
    dataProcessor: (partnerData: RemunerationData[]) => any[][]
  ): any[][] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const noDataFound = Array(headers.length).fill("").map((_, index) => index === 0 ? noDataMessage : "");

    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Add header
    result.push(headers);

    // Process data using the provided processor function
    const dataRows = dataProcessor(partnerData);
    result.push(...dataRows);

    // If no data was processed, add info message
    if (result.length === 1) {
      result.push(noDataFound);
    }

    return result;
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
    const noDataFound = ["No holiday allowances data available", "", "", "", "", "", "", "", ""];
    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Check if holidayAllowance is pensionable for this partner
    const pensionableComponents = this.getPensionableComponents(partnerId);
    const isHolidayAllowancePensionable = pensionableComponents.some(earning => 
      earning.code === "CAOHolidayAllowance"
    );

    // Add header
    result.push([
      "Allowance Name",
      "Amount",
      "Amount Type",
      "Interval",
      "Based On",
      "Individual Choice Budget",
      "Pensionable",
      "Conditions",
      "Description",
    ]);

    // Process each remuneration record
    partnerData.forEach((record) => {
      if (record.holidayAllowances && record.holidayAllowances.length > 0) {
        record.holidayAllowances.forEach((holidayAllowance) => {
          const amountDisplay = this.formatAmountDisplay(holidayAllowance.amount, holidayAllowance.amountType);

          result.push([
            "Holiday Allowance",
            amountDisplay,
            holidayAllowance.amountType?.value || "N/A",
            holidayAllowance.interval || "N/A",
            holidayAllowance.basedOn || "N/A",
            (holidayAllowance as any).isPartOfIndividualChoiceBudget ? "true" : "",
            isHolidayAllowancePensionable ? "true" : "",
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
    const noDataFound = ["No allowances data available", "", "", "", "", "", "", "", ""];

    if (!partnerData || partnerData.length === 0) {
      return [noDataFound];
    }

    const result: any[][] = [];

    // Check if allowances are pensionable for this partner
    const pensionableComponents = this.getPensionableComponents(partnerId);
    const isAllowancesPensionable = pensionableComponents.some(earning => 
      earning.code === "CAOAllowances"
    );

    // Add main header with allowance name first, then component fields
    result.push([
      "Allowance Name",
      "Amount",
      "Amount Type",
      "Based On",
      "Interval",
      "Individual Choice Budget",
      "Pensionable",
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
              const amountDisplay = this.formatAmountDisplay(component.amount, component.amountType);

              result.push([
                allowanceName,
                amountDisplay,
                component.amountType?.value || "N/A",
                component.basedOn || "N/A",
                component.interval || "N/A",
                "", // Allowances are typically not part of ICB
                isAllowancesPensionable ? "true" : "",
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
              "",
              isAllowancesPensionable ? "true" : "",
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
   * Generate travel expenses data for Excel format
   * Processes travel expenses with their components similar to allowances
   */
  generateTravelExpensesData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No travel expenses data available",
      [
        "Component Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        // Check if travel expenses are pensionable for this partner
        const pensionableComponents = this.getPensionableComponents(partnerId);
        const isTravelExpensesPensionable = pensionableComponents.some(earning => 
          earning.code === "CAOTravelExpenses" || earning.code === "CAOTravelTime"
        );
        
        partnerData.forEach((record) => {
          if (record.travelExpensesData && record.travelExpensesData.hasTravelExpenses && record.travelExpensesData.travelExpenses) {
            // Process each travel expense
            record.travelExpensesData.travelExpenses.forEach((travelExpense) => {
              const compensationType = travelExpense.compensationType?.value || "Unknown";

              // Process travel expense components (similar to allowance components)
              if (travelExpense.components && travelExpense.components.length > 0) {
                travelExpense.components.forEach((component) => {
                  const amountDisplay = this.formatAmountDisplay(component.amount, component.amountType);

                  // Determine component name based on compensation type and amount type
                  let componentName = "Travel Expense";
                  if (component.amountType?.value === "Hours") {
                    componentName = "Travel Time";
                  } else if (compensationType === "Money") {
                    componentName = "Travel Allowance";
                  } else if (compensationType === "Both") {
                    componentName = component.amountType?.value === "Hours" ? "Travel Time" : "Travel Allowance";
                  }

                  rows.push([
                    componentName,
                    amountDisplay,
                    component.amountType?.value || "N/A",
                    compensationType || "N/A",
                    component.interval || "N/A",
                    "", // Travel expenses typically not part of ICB
                    isTravelExpensesPensionable ? "true" : "",
                    component.conditions || "N/A",
                    component.description || "N/A",
                  ]);
                });
              } else {
                // If no components, add a basic travel expense entry
                rows.push([
                  "Travel Expense",
                  "N/A",
                  "N/A",
                  compensationType || "N/A",
                  "N/A",
                  "", // Travel expenses typically not part of ICB
                  isTravelExpensesPensionable ? "true" : "",
                  "N/A",
                  `Travel expense with ${compensationType} compensation`,
                ]);
              }
            });
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate wage increments data for Excel format
   * Covers incrementData, retroactiveIncrementData, and periodicalData from wageIncrements
   */
  generateWageIncrementsData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No wage increments data available",
      [
        "Component Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        partnerData.forEach((record) => {
          if (record.wageIncrements && record.wageIncrements.hasWageIncrements) {
            // Process regular increment data
            if (record.wageIncrements.incrementData && record.wageIncrements.incrementData.length > 0) {
              record.wageIncrements.incrementData.forEach((increment) => {
                const amountDisplay = this.formatAmountDisplay(increment.percentage, { 
                  value: increment.incrementType === "%" ? "Percentage" : "Amount", 
                  otherValue: null 
                });

                rows.push([
                  "Regular Wage Increment",
                  amountDisplay,
                  increment.incrementType === "%" ? "Percentage" : "Amount",
                  "Salary",
                  `Effective ${increment.date}`,
                  "", // Wage increments typically not part of ICB
                  "", // Pensionable - dummy column for consistency
                  increment.appliesToAllFunctions ? "Applies to all functions" : "Function-specific",
                  increment.explanation || `Wage increment of ${increment.percentage}${increment.incrementType} effective ${increment.date}`,
                ]);
              });
            }

            // Process retroactive increment data
            if (record.wageIncrements.retroactiveIncrementData && record.wageIncrements.retroactiveIncrementData.length > 0) {
              record.wageIncrements.retroactiveIncrementData.forEach((increment) => {
                const amountDisplay = this.formatAmountDisplay(increment.percentage, { 
                  value: increment.incrementType === "%" ? "Percentage" : "Amount", 
                  otherValue: null 
                });

                rows.push([
                  "Retroactive Wage Increment",
                  amountDisplay,
                  increment.incrementType === "%" ? "Percentage" : "Amount",
                  "Salary",
                  `Effective ${increment.date}`,
                  "", // Wage increments typically not part of ICB
                  "", // Pensionable - dummy column for consistency
                  increment.appliesToAllFunctions ? "Applies to all functions" : "Function-specific",
                  increment.explanation || `Retroactive wage increment of ${increment.percentage}${increment.incrementType} effective ${increment.date}`,
                ]);
              });
            }

            // Process periodical data
            if (record.wageIncrements.periodicalData && record.wageIncrements.periodicalData.length > 0) {
              record.wageIncrements.periodicalData.forEach((periodical) => {
                rows.push([
                  "Periodical Increment",
                  "N/A",
                  "Schedule-based",
                  "Performance & Experience",
                  `${periodical.experiencePeriodAmount} ${periodical.experiencePeriodAmountType}`,
                  "", // Wage increments typically not part of ICB
                  "", // Pensionable - dummy column for consistency
                  `${periodical.eligibleType} - ${periodical.performanceMoment}`,
                  periodical.periodicals || "Performance-based salary progression",
                ]);
              });
            }
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Retrieve ALL components that are pensionable; this list is used to mark the components
   */
  getPensionableComponents(partnerId: string): PensionableEarning[] {
    const partnerData = this.getPartnerRemuneration(partnerId);
    const pensionData = partnerData.length > 0 ? partnerData[0].pensionData : null;

    //console.log('Pension Data:', pensionData);

    if (!pensionData || !pensionData.items || pensionData.items.length === 0) {
      return [];
    }

    const allPensionableEarnings: PensionableEarning[] = [];

    pensionData.items.forEach((item) => {
      if (item.pensionableEarnings && item.pensionableEarnings.length > 0) {
        allPensionableEarnings.push(...item.pensionableEarnings);
      }
    });

    // Enhance with component mapping
    const enhancedEarnings = this.enhancePensionableEarningsWithComponents(allPensionableEarnings);
    
    //console.log('Pensionable Earnings:', enhancedEarnings);
    return enhancedEarnings;
  }

  /**
   * Generate pension data for Excel format
   * Similar to holiday allowances - simple table format
   */
  generatePensionData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No pension data available",
      [
        "Component Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        partnerData.forEach((record) => {
          if (record.pensionData && record.pensionData.hasPension && record.pensionData.items) {
            record.pensionData.items.forEach((item) => {
              const amountDisplay = this.formatAmountDisplay(item.amount, item.amountType);

              // determine pensionable components
              const pensionableItems = item.pensionableEarnings && item.pensionableEarnings.length > 0
                ? item.pensionableEarnings.map(pe => pe.title).join(", ")
                : "N/A";

              rows.push([
                "Pension",
                amountDisplay,
                item.amountType?.value || "N/A",
                pensionableItems || "N/A",
                item.interval || "N/A",
                "", // Pension typically not part of ICB
                "", // Pensionable - dummy column for consistency
                "N/A",
                (item as any).description || "N/A",
              ]);
            });
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate leave data for Excel format
   * Each subcomponent treated as a new row like allowances
   */
  generateLeaveData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No leave data available",
      [
        "Leave Type",
        "Amount",
        "Amount Type",
        "Interval",
        "Based On",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        // Check if holiday/leave components are pensionable for this partner
        const pensionableComponents = this.getPensionableComponents(partnerId);
        const isLeavePensionable = pensionableComponents.some(earning => 
          earning.code === "CAOHoliday"
        );
        
        // Helper function to process leave components
        const processLeaveComponents = (components: LeaveComponent[], typeName: string) => {
          components.forEach((component) => {
            const amountDisplay = this.formatAmountDisplay(component.amount, component.amountType);

            rows.push([
              typeName,
              amountDisplay,
              component.amountType?.value || "N/A",
              component.interval || "N/A",
              "", // Based On - not in LeaveComponent interface
              (component as any).isPartOfIndividualChoiceBudget ? "true" : "",
              isLeavePensionable ? "true" : "",
              component.conditions || "N/A",
              component.description || "N/A",
            ]);
          });
        };

        partnerData.forEach((record) => {
          if (record.leave) {
            const leave = record.leave;

            // Process each leave type using the helper function
            if (leave.hasWorkingHoursReduction && leave.workingHoursReductions) {
              processLeaveComponents(leave.workingHoursReductions, "Working Hours Reduction");
            }

            if (leave.hasVacationDays && leave.vacationDays) {
              processLeaveComponents(leave.vacationDays, "Vacation Days");
            }

            if (leave.hasSpecialLeave && leave.specialLeave) {
              processLeaveComponents(leave.specialLeave, "Special Leave");
            }

            if (leave.hasAdditionalWAZO && leave.additionalWAZO) {
              processLeaveComponents(leave.additionalWAZO, "Additional WAZO");
            }

            if (leave.hasMandatoryLeave && leave.mandatoryLeave) {
              processLeaveComponents(leave.mandatoryLeave, "Mandatory Leave");
            }
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate sick pay data for Excel format
   * Each sick pay amount treated as a new row like allowances
   */
  generateSickPayData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No sick pay data available",
      [
        "Component Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        partnerData.forEach((record) => {
          if (
            record.sickPay &&
            record.sickPay.sickPayAmounts &&
            Array.isArray(record.sickPay.sickPayAmounts) &&
            record.sickPay.sickPayAmounts.length > 0
          ) {
            record.sickPay.sickPayAmounts.forEach((sickPayAmount) => {
              const amountDisplay = this.formatAmountDisplay(sickPayAmount.amount, sickPayAmount.amountType);

              rows.push([
                `Sick Pay (${sickPayAmount.sickPayYear?.value || "Unknown Year"})`,
                amountDisplay,
                sickPayAmount.amountType?.value || "N/A",
                sickPayAmount.basedOn || "N/A",
                sickPayAmount.interval || "N/A",
                "", // Sick pay typically not part of ICB
                "", // Pensionable - dummy column for consistency
                "N/A",
                "N/A",
              ]);
            });
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate individual choice budget data for Excel format
   * Includes items and ICB column
   */
  generateIndividualChoiceBudget(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No individual choice budget data available",
      [
        "Component Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        partnerData.forEach((record) => {
          if (
            record.individualChoiceBudget &&
            record.individualChoiceBudget.hasIndividualChoiceBudget
          ) {
            const budget = record.individualChoiceBudget;

            // Add main budget entry
            const amountDisplay = this.formatAmountDisplay(budget.amount, budget.amountType);

            rows.push([
              "Individual Choice Budget",
              amountDisplay,
              budget.amountType?.value || "N/A",
              budget.basedOn || "N/A",
              budget.interval || "N/A",
              "true", // Main budget is always part of ICB
              "", // Pensionable - dummy column for consistency
              "N/A",
              "Main budget allocation",
            ]);

            // Add items if they exist
            if (budget.items && budget.items.length > 0) {
              budget.items.forEach((item) => {
                const itemAmountDisplay = this.formatAmountDisplay(item.amount, item.amountType);

                const componentName =
                  item.name === "Other" && item.typeOther
                    ? item.typeOther
                    : item.name || "Unnamed Component";

                rows.push([
                  componentName,
                  itemAmountDisplay,
                  item.amountType?.value || "N/A",
                  item.basedOn || "N/A",
                  item.interval || "N/A",
                  item.isPartOfIndividualChoiceBudget ? "true" : "",
                  "", // Pensionable - dummy column for consistency
                  "N/A", // No conditions field in the data structure
                  item.description || "N/A",
                ]);
              });
            }
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate reimbursements data for Excel format
   * Each reimbursement treated as a new row like allowances
   */
  generateReimbursementsData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No reimbursements data available",
      [
        "Reimbursement Name",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        partnerData.forEach((record) => {
          if (
            record.reimbursementsData &&
            record.reimbursementsData.hasReimbursements &&
            record.reimbursementsData.reimbursements
          ) {
            record.reimbursementsData.reimbursements.forEach((reimbursement) => {
              const amountDisplay = this.formatAmountDisplay(reimbursement.amount, reimbursement.amountType);

              rows.push([
                reimbursement.name || "Unnamed Reimbursement",
                amountDisplay,
                reimbursement.amountType?.value || "N/A",
                reimbursement.basedOn || "N/A",
                reimbursement.interval || "N/A",
                "", // Reimbursements typically not part of ICB
                "", // Pensionable - dummy column for consistency
                reimbursement.conditions || "N/A",
                reimbursement.description || "N/A",
              ]);
            });
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Generate payments data for Excel format
   * Handles fixed payments, one-time payments, variable payments, and anniversary payments
   */
  generatePaymentsData(partnerId: string): any[][] {
    return this.generateDataWithStandardStructure(
      partnerId,
      "No payments data available",
      [
        "Payment Type",
        "Amount",
        "Amount Type",
        "Based On",
        "Interval",
        "Individual Choice Budget",
        "Pensionable",
        "Conditions",
        "Description",
      ],
      (partnerData) => {
        const rows: any[][] = [];
        
        // Check if payments are pensionable for this partner
        const pensionableComponents = this.getPensionableComponents(partnerId);
        const isFixedPaymentsPensionable = pensionableComponents.some(earning => 
          earning.code === "CAOFixedPayment"
        );
        const isOneTimePaymentsPensionable = pensionableComponents.some(earning => 
          earning.code === "CAOOneTimePayment"
        );
        const isVariablePaymentsPensionable = pensionableComponents.some(earning => 
          earning.code === "CAOVariablePayment"
        );
        
        partnerData.forEach((record) => {
          if (record.paymentsData) {
            // Process fixed payments
            if (record.paymentsData.hasFixedPayments && record.paymentsData.fixedPayments) {
              record.paymentsData.fixedPayments.forEach((payment) => {
                const amountDisplay = this.formatAmountDisplay(payment.amount, payment.amountType);
                const basedOnText = payment.basedOn === "Other" && payment.basedOnOther
                  ? payment.basedOnOther
                  : payment.basedOn || "N/A";

                rows.push([
                  "Fixed Payment",
                  amountDisplay,
                  payment.amountType?.value || "N/A",
                  basedOnText,
                  payment.interval || "N/A",
                  payment.isPartOfIndividualChoiceBudget ? "true" : "",
                  isFixedPaymentsPensionable ? "true" : "",
                  payment.date ? `Payment date: ${payment.date}` : "N/A",
                  payment.description || "N/A",
                ]);
              });
            }

            // Process one-time payments
            if (record.paymentsData.hasOneTimePayments && record.paymentsData.oneTimePayments) {
              record.paymentsData.oneTimePayments.forEach((payment) => {
                const amountDisplay = this.formatAmountDisplay(payment.amount, payment.amountType);
                const basedOnText = payment.basedOn === "Other" && payment.basedOnOther
                  ? payment.basedOnOther
                  : payment.basedOn || "N/A";

                rows.push([
                  "One-Time Payment",
                  amountDisplay,
                  payment.amountType?.value || "N/A",
                  basedOnText,
                  payment.interval || "N/A",
                  payment.isPartOfIndividualChoiceBudget ? "true" : "",
                  isOneTimePaymentsPensionable ? "true" : "",
                  payment.date ? `Payment date: ${payment.date}` : "N/A",
                  payment.description || "N/A",
                ]);
              });
            }

            // Process variable payments
            if (record.paymentsData.hasVariablePayments && record.paymentsData.variablePayments) {
              record.paymentsData.variablePayments.forEach((payment) => {
                let amountDisplay = this.formatAmountDisplay(payment.amount, payment.amountType);
                
                // Add min/max info if available
                if (payment.hasMinMaxAmount && (payment.minAmount !== null || payment.maxAmount !== null)) {
                  const minText = payment.minAmount !== null ? `Min: ${payment.minAmount}` : "";
                  const maxText = payment.maxAmount !== null ? `Max: ${payment.maxAmount}` : "";
                  const rangeText = [minText, maxText].filter(Boolean).join(", ");
                  if (rangeText) {
                    amountDisplay += ` (${rangeText})`;
                  }
                }

                const basedOnText = payment.basedOn === "Other" && payment.basedOnOther
                  ? payment.basedOnOther
                  : payment.basedOn || "N/A";

                const paymentName = payment.name ? `Variable Payment (${payment.name})` : "Variable Payment";

                rows.push([
                  paymentName,
                  amountDisplay,
                  payment.amountType?.value || "N/A",
                  basedOnText,
                  payment.interval || "N/A",
                  "", // Variable payments typically not part of ICB
                  isVariablePaymentsPensionable ? "true" : "",
                  payment.date ? `Payment date: ${payment.date}` : "N/A",
                  payment.description || "N/A",
                ]);
              });
            }

            // Process anniversary payments (existing logic)
            if (record.paymentsData.hasAnniversaryPayments && record.paymentsData.anniversaryPayments) {
              record.paymentsData.anniversaryPayments.forEach((payment) => {
                const amountDisplay = this.formatAmountDisplay(payment.amount, payment.amountType);

                const paymentName = `Anniversary Payment (${payment.amountOfYears} years)`;
                const basedOnText = payment.basedOn === "Other" && payment.basedOnOther
                  ? payment.basedOnOther
                  : payment.basedOn || "N/A";

                const conditionsText = payment.dateOrCondition === "Other" && payment.dateOrConditionOther
                  ? payment.dateOrConditionOther
                  : payment.dateOrCondition || "N/A";

                rows.push([
                  paymentName,
                  amountDisplay,
                  payment.amountType?.value || "N/A",
                  basedOnText,
                  payment.interval || "N/A",
                  "", // Anniversary payments typically not part of ICB
                  "", // Pensionable - dummy column for consistency
                  conditionsText,
                  payment.description || "N/A",
                ]);
              });
            }
          }
        });
        
        return rows;
      }
    );
  }

  /**
   * Extracts and concatenates all pensionable earnings from pension data items
   * @param pensionData The pension data object containing items with pensionable earnings
   * @returns Array of all pensionable earnings from all items with component mapping
   */
  extractAllPensionableEarnings(pensionData: PensionData): PensionableEarning[] {
    if (!pensionData || !pensionData.items || pensionData.items.length === 0) {
      return [];
    }

    const allPensionableEarnings: PensionableEarning[] = [];

    pensionData.items.forEach((item) => {
      if (item.pensionableEarnings && item.pensionableEarnings.length > 0) {
        allPensionableEarnings.push(...item.pensionableEarnings);
      }
    });

    // Enhance with component mapping
    return this.enhancePensionableEarningsWithComponents(allPensionableEarnings);
  }

  /**
   * Get all unique component types from pensionable earnings for a partner
   * @param partnerId Partner ID to filter data
   * @returns Array of unique component names
   */
  getPensionableComponentTypes(partnerId: string): string[] {
    const pensionableEarnings = this.getPensionableComponents(partnerId);
    const componentTypes = pensionableEarnings
      .map(earning => earning.component)
      .filter((component, index, array) => component && array.indexOf(component) === index);
    
    return componentTypes as string[];
  }

  /**
   * Check if a specific component type is pensionable for a partner
   * @param partnerId Partner ID to filter data
   * @param componentType The component type to check (e.g., "allowances", "holidayAllowance")
   * @returns True if the component type is pensionable
   */
  isComponentTypePensionable(partnerId: string, componentType: string): boolean {
    const componentTypes = this.getPensionableComponentTypes(partnerId);
    return componentTypes.includes(componentType);
  }
}

// Create a singleton instance
export const dataLoader = new DataLoaderService();
