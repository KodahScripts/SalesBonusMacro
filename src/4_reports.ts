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