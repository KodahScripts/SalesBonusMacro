function main(workbook: ExcelScript.Workbook) {
    const initialSheets = workbook.getWorksheets();

    const lookup = {};
    const reports = {};

    initialSheets.map(sheet => {
        const sheetName = sheet.getName();
        const data = sheet.getUsedRange().getValues();
        if (sheetName.includes("-")) {
            const splitName = sheetName.split("-");
            if (splitName.length > 1) {
                const typeOfReport = splitName[0].trim().toLowerCase(); // In case we need more control
                const reportName = splitName[1].trim().toLowerCase();
                lookup[reportName] = new LookupReport(data);
            }
        } else {
            switch (sheetName) {
                case "INPUT": reports["globals"] = new InputReport(data);
                    break;
                case "0432": reports["deals"] = new DealsReport(data);
                    break;
                case "90": reports["averageunits"] = new AverageUnitReport(data);
                    break;
                case "3213": reports["priordraw"] = new PriorDrawReport(data);
                    break;
                case "SPIFFS": reports["spiffs"] = new SpiffReport(data);
                    break;
                case "NpsSheet":
                    const head = data[2];
                    const body = data.slice(3)
                    reports["nps"] = new NpsReport([head, ...body]);
                    break;
                default:
                    sheet.delete();
                    break;
            }
        }
    });

    const { storeName, storeAbbr, date, regionalScore, topsalesAmount, ...accounts}: Store = reports["globals"].vars;
    const { retroAcct, expenseAcct1, expenseAcct2, salesTaxAcct, salesBonusAcct1, salesBonusAcct2 }: AccountNumbers = accounts;
    const accountNumbers: AccountNumbers = { retro: retroAcct, expense1: expenseAcct1, expense2: expenseAcct2, salesTax: salesTaxAcct, salesBonus1: salesBonusAcct1, salesBonus2: salesBonusAcct2 }
    const store: Store = {
        name: storeName, abbr: storeAbbr, displayDate: date, employees: [],
        regionalScore: regionalScore, topsalesAmount, accountNumbers
    };

    const employeeList = getAllEmployees(reports);
    attachSpiffs(reports["spiffs"].list, employeeList);
    attachNps(reports["nps"].list, employeeList, store.regionalScore);
    attachPriorDraw(reports["priordraw"].list, employeeList);
    attachUnits(reports["averageunits"].list, employeeList);
    attachDeals(reports["deals"].list, employeeList);
    calculateRetro(lookup["retromini"], lookup["retropercentage"], employeeList);
    calculateUnitBonus(lookup["unit"], employeeList);
    // console.log(lookup)

    store.employees = employeeList;

    // for(const [k, v] of Object.entries(lookup)) {
    //     console.log(k, v.rows)
    // }

    console.log(store)
}


function getAllEmployees(reports: InputReport | DealsReport | AverageUnitReport | PriorDrawReport | SpiffReport | NpsReport) {
    const allEmps: Employee[] = [];
    for (const [key, val] of Object.entries(reports)) {
        switch (key) {
            case "globals":
            case "spiffs":
                break;
            default:
                val.list.map(v => {
                    const employee = allEmps.filter(emp => emp.id === v.employeeID)[0];
                    const nps: Nps = { surveyCount: 0, currentPercent: 0, averagePercent: 0, outcome: "B" };
                    const units: Units = { average: 0, count: 0, threeMonthCount: 0, total: { new: 0, used: 0 } };
                    if (!employee) allEmps.push({ id: v.employeeID, name: v.employeeName, deals: [], nps, spiffsTotal: 0, priorDraw: 0, retroPercent: 0, bonus: { unit: 0, topsales: 0 }, units });
                });
                break;
        }
    }
    return allEmps;
}


function attachSpiffs(spiffs: EmployeeSpiff[], employeeList: Employee[]) {
    spiffs.forEach(spiff => {
        employeeList.filter(emp => emp.id === spiff.employeeID)[0].spiffsTotal = spiff.amount;
    });
}


function attachNps(npsList: EmployeeNps[], employeeList: Employee[], regionalScore: number) {
    const buffed = regionalScore + (regionalScore * 0.03);
    npsList.forEach(nps => {
        const employee = employeeList.filter(emp => emp.id === nps.employeeID)[0];
        const score = nps.currentPercent > nps.averagePercent ? nps.currentPercent : nps.averagePercent;
        employee.nps.surveyCount = nps.surveyCount;
        employee.nps.averagePercent = nps.averagePercent;
        employee.nps.currentPercent = nps.currentPercent;

        if (score * 100 >= buffed) {
            employee.nps.outcome = "3P"
        } else if (score === regionalScore) {
            employee.nps.outcome = "A"
        } else {
            employee.nps.outcome = "B"
        }
    });
}


function attachPriorDraw(pd: PriorDraw[], employeeList: Employee[]) {
    pd.forEach(d => {
        employeeList.filter(emp => emp.id === d.employeeID)[0].priorDraw = d.draw;
    });
}


function attachUnits(units: EmployeeUnits[], employeeList: Employee[]) {
    units.forEach(unit => {
        const employee = employeeList.filter(emp => emp.id === unit.employeeID)[0];
        employee.units.average = unit.average;
        employee.units.threeMonthCount = unit.count;
        employee.units.count = 0;
        employee.units.total = { new: 0, used: 0 };
    });
}


function attachDeals(deals: DealRow[], employeeList: Employee[]) {
    employeeList.forEach(employee => {
        employee.deals = deals.filter(emp => emp.employeeID === employee.id).map(deal => {
            employee.units.total[deal.vehicleType.toLowerCase()] += deal.dealUnitCount;
            employee.units.count += deal.dealUnitCount;
            return {
                id: deal.dealID,
                customer: { id: deal.customerID, name: deal.customerName },
                vehicle: new Vehicle(deal.vehicleID, deal.vehicleDescription, deal.vehicleType),
                unitCount: deal.dealUnitCount,
                commission: { fni: deal.commissionFnI, gross: deal.commissionGross, amount: deal.commissionAmount },
                retro: { mini: 0, owed: 0, payout: 0, total: 0 }
            }
        });
    });
}


function calculateRetro(retroMiniLookup: LookupReport, retroPercentLookup: LookupReport, employees: Employee[]) {
    employees.forEach(employee => {
        employee.retroPercent = retroPercentLookup.getValue(employee.units.count);
        employee.deals.forEach(deal => {
            const mini = deal.commission.amount <= 251 ? retroMiniLookup.getValue(employee.units.average) * deal.unitCount : 0;
            const owed = mini > 0 ? mini - deal.commission.amount : 0;
            const payout = mini === 0 ? deal.commission.gross * employee.retroPercent : 0;
            const total = owed + payout;
            deal.retro = { mini, owed, payout, total };
        });
    });
}


function calculateUnitBonus(unitBonusLookup: LookupReport, employees: Employee[]) {
    employees.forEach(employee => {
        employee.bonus.unit = unitBonusLookup.getValue(employee.units.count)
    });
}


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
    units: Units;
}

interface Nps {
    surveyCount: number;
    currentPercent: number;
    averagePercent: number;
    outcome: string;
}

interface Units {
    average: number;
    threeMonthCount: number;
    count: number;
    total: {
        new: number;
        used: number;
    }
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

interface Retro {
    mini: number;
    owed: number;
    payout: number;
    total: number;
}

interface Bonus {
    unit: number;
    topsales: number;
}

class Vehicle {
    public year: number;
    public make: string;
    public model: string;
    public desc: string;
    constructor(public id: string, protected description: string, public saleType: string) {
        this.id = id;
        this.saleType = saleType;

        const data = description.split(',');
        this.year = Number(data[0]);
        this.make = String(data[1]);
        this.model = String(data[2]);
        this.desc = String(data[3]);
    }
}


interface LookupRow {
    min: number;
    max: number;
    val: number;
}
class LookupReport {
    public rows: LookupRow[];
    constructor(private data: Array<string | number | boolean>[]) {
        this.rows = [];
        data.map((d, i) => {
            const next = data[i + 1];
            const max = next ? Number(data[i + 1][0]) : 100000000;
            this.rows.push({ min: Number(d[0]), max, val: Number(d[1]) });
        });
    }
    getValue(query: number) {
        if (this.rows.some(row => query >= row.min && query < row.max)) return this.rows.filter(row => query >= row.min && query < row.max)[0].val;
        return 0;
        // console.log(this.rows, query)
    }
}


class Report {
    private header: Array<string | number | boolean>;
    constructor(protected data: Array<string | number | boolean>[]) {
        this.header = data.shift();
        this.data = data;
    }
    getColumn(columnLabel: string): number {
        return this.header.indexOf(columnLabel);
    }
    getValue(columnLabel: string, rowIndex: number = 0): string | number | boolean {
        return this.data[rowIndex][this.getColumn(columnLabel)];
    }
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
class InputReport extends Report {
    public vars: StoreVars;
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.vars = {
            storeName: String(this.getValue("Store Name")),
            storeAbbr: String(this.getValue("Store Abbr")),
            date: String(this.getValue("Date")),
            regionalScore: Number(this.getValue("Regional Score")),
            topsalesmanBonusAmount: Number(this.getValue("Topsalesman Bonus")),
            retroAcct: String(this.getValue("Retro Acct")),
            expenseAcct1: String(this.getValue("Expense 1 Acct")),
            expenseAcct2: String(this.getValue("Expense 2 Acct")),
            salesTaxAcct: String(this.getValue("Sales Tax Acct")),
            salesBonusAcct1: String(this.getValue("Sales Bonus Tax 1")),
            salesBonusAcct2: String(this.getValue("Sales Bonus Tax 2")),
        }
    }
}


interface EmployeeUnits {
    employeeID: number;
    employeeName: string;
    count: number;
    average: number;
}
class AverageUnitReport extends Report {
    public list: EmployeeUnits[];
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Salesperson#", i));
            const employeeName = String(this.getValue("Salesperson Name", i));
            const count = Number(this.getValue("units", i));
            const average = Number((count / 3).toFixed(1));
            this.list.push({ employeeID, employeeName, count, average });
        });
    }
}


interface PriorDraw {
    employeeID: number;
    employeeName: string;
    commission: number;
    draw: number;
}
class PriorDrawReport extends Report {
    public list: PriorDraw[];
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Control#", i));
            const employeeName = String(this.getValue("Description", i));
            const commission = Number(this.getValue("8321C", i));
            const draw = Number(this.getValue("8321D", i));
            this.list.push({ employeeID, employeeName, commission, draw });
        });
    }
}


interface EmployeeSpiff {
    employeeID: number;
    amount: number;
}
class SpiffReport extends Report {
    public list: EmployeeSpiff[];
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Employee #", i));
            const amount = Number(this.getValue("Total", i));
            if (employeeID > 0) this.list.push({ employeeID, amount });
        });
    }
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
class DealsReport extends Report {
    public list: DealRow[];
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const date = Number(this.getValue("Date", i));
            const dealID = String(this.getValue("Reference#", i));
            const employeeID = Number(this.getValue("Salesperson#", i));
            const employeeName = String(this.getValue("Salesperson Name", i));
            const customerID = Number(this.getValue("Customer#", i));
            const customerName = String(this.getValue("Customer Name", i));
            const vehicleID = String(this.getValue("Stock#", i));
            const vehicleDescription = String(this.getValue("Description", i));
            const vehicleType = String(this.getValue("Sale Type", i));
            const commissionFnI = Number(this.getValue("COMMBL F&I", i));
            const commissionGross = Number(this.getValue("COMMBL FRONT", i));
            const dealUnitCount = Number(this.getValue("Units", i));
            const commissionAmount = Number(this.getValue("Commission Amount", i));
            if (dealUnitCount > 0) this.list.push({ date, dealID, employeeID, employeeName, customerID, customerName, vehicleID, vehicleDescription, vehicleType, commissionFnI, commissionGross, dealUnitCount, commissionAmount });
        });
    }
}


interface EmployeeNps {
    employeeID: number;
    employeeName: string;
    surveyCount: number;
    currentPercent: number;
    averagePercent: number;
}
class NpsReport extends Report {
    public list: EmployeeNps[];
    constructor(protected data: Array<string | number | boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Employee #", i));
            const employeeName = String(this.getValue("BMWNC TEAM", i));
            const surveyCount = Number(this.getValue("PROMOTER", i));
            const currentPercent = Number(this.getValue("NPS%", i));
            const averagePercentCol = 23; // Not Ideal, need a better way to get this number
            const averagePercent = Number(d[averagePercentCol]);
            if (employeeID > 0) this.list.push({ employeeID, employeeName, surveyCount, currentPercent, averagePercent });
        });
    }
}