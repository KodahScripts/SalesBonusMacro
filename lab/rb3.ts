function main(workbook: ExcelScript.Workbook) {
    const initialSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    const store: Store = {
        name: "BMW of South Miami",
        abbr: "BOSM",
        salesTotals: {
            new: 0,
            used: 0
        },
        employees: []
    };
    
    

    initialSheets.forEach(sheet => {
        const reportData = sheet.getUsedRange().getValues();
        const sheetName = sheet.getName();
        switch(sheetName) {
            case '0432': 
                const header = reportData.shift();
                reportData.forEach(row => {
                    const empID = row[header.indexOf("Salesperson#")];
                    const customer: Person = { id: row[header.indexOf("Customer#")], name: row[header.indexOf("Customer Name")] };
                    const dealId = row[header.indexOf("Reference#")];
                    const date = row[header.indexOf("Date")];
                    const vehicle: Vehicle = new Vehicle(row[header.indexOf("Stock#")], row[header.indexOf("Description")], row[header.indexOf("Sale Type")]);
                    const unitCount = Number(row[header.indexOf("Units")]);
                    const retro: Retro = { mini: 0, owed: 0, payout: 0 };
                    const commission: Commission = { fni: row[header.indexOf("COMMBL F&I")], gross: row[header.indexOf("COMMBL FRONT")], amount: row[header.indexOf("Commission Amount")], retro };
                    const deal: Deal = { id: dealId, date, customer, vehicle, unitCount, commission }
                    if(!store.employees.some(emp => emp.id === empID)) store.employees.push(new Employee(empID, row[header.indexOf("Salesperson Name")]));
                    let employee: Employee = store.employees.find(emp => emp.id === empID);
                    if(deal.unitCount > 0) employee.addDeal(deal);
                });
                break;
            case '90':
                break;
            case '3213':
                break;
            case 'SPIFFS':
                break;
            case 'NPS Sheet':
                break;
            case 'Look Up Table':
                break;
            case 'INPUT':
                break;
            default: sheet.delete();
                break;
        }
        store.employees.forEach(emp => console.log(emp.getTotalUnits()))
    });
}

interface Store {
    name: string;
    abbr: string;
    salesTotals: SalesTotals;
    topSalesman: TopSalesman;
    employees: Employee[];
}

interface Deal {
    id: string;
    date: number;
    customer: Person;
    vehicle: Vehicle;
    unitCount: number;
    commission: Commission;
}

interface UnitAverage {
    units: number;
    average: number;
    rounded: number;
}

interface NPS {
    employeeId: {
        surveys: number;
        current: number;
        average: number;
    }
}

interface Person {
    id: number;
    name: string;
}

interface Commission {
    fni: number;
    gross: number;
    amount: number;
    retro: Retro;
}

interface Retro {
    mini: number;
    owed: number;
    payout: number;
}

interface PriorDraw {
    id: number;
    amount: number;
}

interface Spiff {
    id: number;
    amount: number;
}

// interface Employee {
//     id: number;
//     name: string;
//     averageSoldUnits: UnitAverage;
//     commissionTotals: Commission;
//     unitCount: number;
//     priorDraw: number;
//     spiffs: number;
//     nps: NPS;
//     retroPercentage: number;
//     retroTotal: number;
//     fniTotal: FnI;
//     bonus: Bonus;
// }

interface SalesTotals {
    new: number;
    used: number;
}

interface TopSalesman {
    id: number;
    count: number;
}

interface FnI {
    reserve: number;
    gross: number;
    payout: number;
}

interface Bonus {
    unit: number;
    topsales: number;
    total: number;
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

class Employee {
    public deals: Deal[];
    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;
        this.deals = [];
    }

    addDeal(deal: Deal) {
        this.deals.push(deal);
    }

    getTotalUnits() {
        this.deals.reduce((acc, curr) => {
            return acc + curr.unitCount;
        }, 0);
    }
}