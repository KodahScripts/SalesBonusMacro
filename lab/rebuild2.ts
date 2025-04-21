function main(workbook: ExcelScript.Workbook) {
    const initialSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    let allEmployees: Person[] = [];
    let allDeals: Deal[] = [];
    let allUnitAverages: UnitAverage[] = [];
    let priorDraws: PriorDraw[] = [];
    let all_spiffs: Spiff[] = [];
    let nps_averages: NPS[] = [];
    const store: Store = {
        name: "BMW of South Miami",
        abbr: "BOSM",
        salesTotals: {
            new: 0,
            used: 0
        },
        topSalesman: {
            id: 0,
            count: 0
        },
        employees: []
    };

    const allowedSheets = ['0432', '90', '3213', 'SPIFFS', 'NPS Sheet', 'Look Up Table', 'Input'];
    initialSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (!allowedSheets.includes(sheetName)) {
            sheet.delete();
        } else {
            const reportData = sheet.getUsedRange().getValues();
            switch (sheetName) {
                case '0432': new Report0432(allEmployees, allDeals, reportData);
                    break;
                case '90': new Report90(allUnitAverages, reportData);
                    break;
                default: 
                    // console.log(sheetName, reportData);
                    break;
            }
        }
    });
    console.log("Employees", allEmployees);
    console.log("Deals", allDeals);
    console.log("Unit Averages", allUnitAverages);
}


class Report0432 {
    constructor(private allEmployees: Person[], private allDeals: Deal[], private data: Array<string | number | boolean>[]) {
        const header: Array<string | number | boolean> = data.shift();
        const dealIdCol: number = header.indexOf("Reference#");
        const empIdCol: number = header.indexOf("Salesperson#");
        const empNameCol: number = header.indexOf("Salesperson Name");
        const dateCol: number = header.indexOf("Date");
        const custIdCol: number = header.indexOf("Customer#");
        const custNameCol: number = header.indexOf("Customer Name");
        const vehIdCol: number = header.indexOf("Stock#");
        const vehDescCol: number = header.indexOf("Description");
        const saleTypeCol: number = header.indexOf("Sale Type");
        const commFniCol: number = header.indexOf("COMMBL F&I");
        const commGrossCol: number = header.indexOf("COMMBL FRONT");
        const unitsCol: number = header.indexOf("Units");
        const commAmountCol: number = header.indexOf("Commission Amount");

        data.forEach(row => {
            const employee: Person = { id: Number(row[empIdCol]), name: String(row[empNameCol]) };
            if (!allEmployees.find(emp => emp.id === employee.id)) allEmployees.push(employee);
            if(Number(row[unitsCol]) > 0) {
                const unitCount = Number(row[unitsCol]);
                if(unitCount > 0) {
                    const employeeId = employee.id;
                    const id = String(row[dealIdCol]);
                    const date = Number(row[dateCol]);
                    const customer: Person = { id: Number(row[custIdCol]), name: String(row[custNameCol]) };
                    const vehicle = new Vehicle(String(row[vehIdCol]), String(row[vehDescCol]), String(row[saleTypeCol]));
                    const retro: Retro = { mini: 0, owed: 0, payout: 0 };
                    const commission: Commission = { fni: Number(row[commFniCol]), gross: Number(row[commGrossCol]), amount: Number(row[commAmountCol]), retro };
                    allDeals.push({ [employeeId]: { id, date, customer, vehicle, unitCount, commission } });
                }
            }
        });
    }
}


class Report90 {
    constructor(private allUnitAverages: UnitAverage[], private data: Array<string | number | boolean>[]) {
        const header: Array<string | number | boolean> = data.shift();
        const empIdCol: number = header.indexOf("Salesperson#");
        const unitCol: number = header.indexOf("Units");
        
        data.forEach(row => {
            const employeeId = Number(row[empIdCol]);
            const allUnits = Number(row[unitCol]);
            const average = allUnits / 3;

            const ua: UnitAverage = {
                [employeeId]: {
                    units: allUnits,
                    average,
                    rounded: Math.round(average)
                }
            };
            allUnitAverages.push(ua);
        });
    }
}


interface Store {
    name: string;
    abbr: string;
    salesTotals: SalesTotals;
    topSalesman: TopSalesman;
    employees: Employee[];
}

interface Deal {
    [employeeId: number]: {
        id: string;
        date: number;
        customer: Person;
        vehicle: Vehicle;
        unitCount: number;
        commission: Commission;
    }
}

interface UnitAverage {
    [employeeId: number]: {
        units: number;
        average: number;
        rounded: number;
    }
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

interface Employee {
    id: number;
    name: string;
    averageSoldUnits: UnitAverage;
    commissionTotals: Commission;
    unitCount: number;
    priorDraw: number;
    spiffs: number;
    nps: NPS;
    retroPercentage: number;
    retroTotal: number;
    fniTotal: FnI;
    bonus: Bonus;
    deals: Deal[];
}

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
    constructor(public id: string, private description: string, public saleType: string) {
        this.id = id;
        this.saleType = saleType;

        const data = description.split(',');
        this.year = Number(data[0]);
        this.make = String(data[1]);
        this.model = String(data[2]);
        this.desc = String(data[3]);
    }
}