interface Store {
    name: string;
    abbr: string;
    displayDate: string;
    employees: Employee[];
    regionalScore: number;
    topsalesAmount: number;
    accountNumbers: AccountNumbers;
}

interface AccountNumbers {
    retro: string;
    expense1: string;
    expense2: string;
    salesTax: string;
    salesBonus1: string;
    salesBonus2: string;
}

interface Employee {
    id: number;
    name: string;
    deals: Deal[];
    nps: Nps;
    spiffsTotal: number;
    priorDraw: number;
    retroPercent: number;
    bonus: Bonus;
    averageUnits: AverageUnits;
    units: Units;
}

interface Deal {
    id: string;
    customer: {
        id: number;
        name: string;
    };
    vehicle: Vehicle;
    unitCount: number;
    commission: Commission;
    retro: Retro;
}

interface Commission {
    fni: number;
    gross: number;
    amount: number;
}

interface Bonus {
    unit: number;
    topsales: number;
    nps: number;
}

interface LookupRow {
    min: number;
    max: number;
    val: number;
}

interface StoreVars {
    storeName: string;
    storeAbbr: string;
    date: string;
    regionalScore: number;
    topsalesmanBonusAmount: number;
    retroAcct: string;
    expenseAcct1: string;
    expenseAcct2: string;
    salesTaxAcct: string;
    salesBonusAcct1: string;
    salesBonusAcct2: string;
}

interface EmployeeUnits {
    employeeID: number;
    employeeName: string;
    count: number;
    average: number;
}

interface PriorDraw {
    employeeID: number;
    employeeName: string;
    commission: number;
    draw: number;
}

interface EmployeeSpiff {
    employeeID: number;
    amount: number;
}

interface DealRow {
    date: number;
    dealID: string;
    employeeID: number;
    employeeName: string;
    customerID: number;
    customerName: string;
    vehicleID: string;
    vehicleDescription: string;
    vehicleType: string;
    commissionFnI: number;
    commissionGross: number;
    dealUnitCount: number;
    commissionAmount: number;
}

interface EmployeeNps {
    employeeID: number;
    employeeName: string;
    surveyCount: number;
    currentPercent: number;
    averagePercent: number;
}