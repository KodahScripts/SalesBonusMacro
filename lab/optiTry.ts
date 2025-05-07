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
            switch(sheetName) {
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
    // console.log("Globals", reports["globals"].vars);
    // console.log("Average Units", reports["averageunits"].list);
    // console.log("Prior Draw", reports["priordraw"].list);
    // console.log("Spiffs", reports["spiffs"].list);
    // console.log("Deals", reports["deals"].list);
    // console.log("NPS", reports["nps"].list);
    // console.log("Lookup", lookup);
    // console.log("Reports", reports);

    // for(const [k, v] of Object.entries(lookup)) {
    //     console.log(k, v.rows)
    // }

    // for(const [k, v] of Object.entries(reports)) {
    //     switch(k) {
    //         case "globals": console.log(k, v.vars);
    //             break;
    //         default: console.log(k, v.list);
    //             break;
    //     }
    // }

    const [storeName, storeAbbr, displayDate, regionalScore, ...accountNumbers] = reports["globals"].vars;
    const [retro, expense1, expense2, salesTax, salesBonus1, salesBonus2] = accountNumbers;
    const accounts: AccountNumbers = { retro, expense1, expense2, salesTax, salesBonus1, salesBonus2 };
    const store: Store = { name: storeName, abbr: storeAbbr, displayDate, regionalScore, accountNumbers: accounts };
    console.log(store);
}

interface Store {
    name: string;
    abbr: string;
    displayDate: string;
    regionalScore: number;
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


interface LookupRow {
    min: number;
    max: number;
    val: number;
}
class LookupReport {
    public rows: LookupRow[];
    constructor(private data: Array<string|number|boolean>[]) {
        this.rows = [];
        data.map((d, i) => {
            const next = data[i + 1];
            const max = next ? Number(data[i + 1][0]) : 0;
            this.rows.push({ min: Number(d[0]), max, val: Number(d[1]) });
        });
    }
    getValue(query: number): number {
        this.rows.map(r => {
            if (r.max === 0 && query >= r.min || query >= r.min && query < r.max) {
                return r.val;
            }
        });
        return 0;
    }
}


class Report {
    private header: Array<string|number|boolean>;
    constructor(protected data: Array<string|number|boolean>[]) {
        this.header = data.shift();
        this.data = data;
    }
    getColumn(columnLabel: string): number {
        return this.header.indexOf(columnLabel);
    }
    getValue(columnLabel: string, rowIndex:number = 0): string | number | boolean {
        return this.data[rowIndex][this.getColumn(columnLabel)];
    }
}


interface StoreVars {
    storeName: string;
    storeAbbr: string;
    date: string;
    regionalScore: number;
    retroAcct: string;
    expenseAcct1: string;
    expenseAcct2: string;
    salesTaxAcct: string;
    salesBonusAcct1: string;
    salesBonusAcct2: string;
}
class InputReport extends Report {
    public vars: StoreVars;
    constructor(protected data: Array<string|number|boolean>[]) {
        super(data);
        this.vars = {
            storeName: String(this.getValue("Store Name")),
            storeAbbr: String(this.getValue("Store Abbr")),
            date: String(this.getValue("Date")),
            regionalScore: Number(this.getValue("Regional Score")),
            retroAcct: String(this.getValue("Retro Acct")),
            expenseAcct1: String(this.getValue("Expense 1 Acct")),
            expenseAcct2: String(this.getValue("Expense 2 Acct")),
            salesTaxAcct: String(this.getValue("Sales Tax Acct")),
            salesBonusAcct1: String(this.getValue("Sales Bonus Tax 1")),
            salesBonusAcct2: String(this.getValue("Sales Bonus Tax 2"))
        }
    }
}


interface EmployeeUnits {
    employeeID: number;
    count: number;
    average: number;
}
class AverageUnitReport extends Report {
    public list: EmployeeUnits[];
    constructor(protected data: Array<string|number|boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Salesperson#", i));
            const count = Number(this.getValue("units", i));
            const average = Number((count / 3).toFixed(1));
            this.list.push({ employeeID, count, average });
        });
    }
}


interface PriorDraw {
    employeeID: number;
    commission: number;
    draw: number;
}
class PriorDrawReport extends Report {
    public list: PriorDraw[];
    constructor(protected data: Array<string|number|boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Control#", i));
            const commission = Number(this.getValue("8321C", i));
            const draw = Number(this.getValue("8321D", i));
            this.list.push({ employeeID, commission, draw });
        });
    }
}


interface EmployeeSpiff {
    employeeID: number;
    amount: number;
}
class SpiffReport extends Report {
    public list: EmployeeSpiff[];
    constructor(protected data: Array<string|number|boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Employee #", i));
            const amount = Number(this.getValue("Total", i));
            if(employeeID > 0) this.list.push({ employeeID, amount });
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
    constructor(protected data: Array<string|number|boolean>[]) {
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
            if(dealUnitCount > 0) this.list.push({ date, dealID, employeeID, employeeName, customerID, customerName, vehicleID, vehicleDescription, vehicleType, commissionFnI, commissionGross, dealUnitCount, commissionAmount });
        });
    }
}


interface EmployeeNps {
    employeeID: number;
    surveyCount: number;
    currentPercent: number;
    averagePercent: number;
}
class NpsReport extends Report{
    public list: EmployeeNps[];
    constructor(protected data: Array<string|number|boolean>[]) {
        super(data);
        this.list = [];
        this.data.map((d, i) => {
            const employeeID = Number(this.getValue("Employee #", i));
            const surveyCount = Number(this.getValue("PROMOTER", i));
            const currentPercent = Number(this.getValue("NPS%", i));
            const averagePercentCol = 23; // Not Ideal, need a better way to get this number
            const averagePercent = Number(d[averagePercentCol]);
            if (employeeID > 0) this.list.push({ employeeID, surveyCount, currentPercent, averagePercent });
        });
    }
}


class Store {
    constructor(public name: string, public abbr: string) {
        this.name = name;
        this.abbr = abbr;
    }
}


class Employee {
    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;
    }
}


class Deal {
    constructor(public id: number) {
        this.id = id;
    }
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