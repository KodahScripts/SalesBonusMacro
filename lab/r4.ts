class Store {
    public employees: Employee[];
    constructor(public name: string, public abbr: string) {
        this.name = name;
        this.abbr = abbr;
        this.employees = [];
    }

    employeeExists(employeeId: number): boolean {
        return this.employees.some(emp => emp.id === employeeId);
    }

    getSaleTypeTotals(): SalesTotals {
        const sales: SalesTotals = { new: 0, used: 0 };
        this.employees.forEach(employee => {
            employee.deals.forEach(deal => {
                if(deal.vehicle.saleType === "New") {
                    sales.new++;
                } else {
                    sales.used++;
                }
            });
        });
        return sales;
    }
}

class Employee {
    public deals: Deal[];
    public spiff: number;
    public averageUnits: number;
    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;
        this.spiff = 0;
        this.deals = [];
    }

    getTotalUnits(): number {
        return this.deals.reduce((acc: number, curr: Deal) => {
            return acc + curr.unitCount;
        }, 0);
    }

    getRetroPercentage(): number {
        const unitCount = this.getTotalUnits();
        if (unitCount >= 16) return 0.07;
        if (unitCount >= 12 && unitCount < 16) return 0.04;
        return 0;
    }
}

class Deal {
    constructor(public id: string, public date: number, public customer: Person, public vehicle: Vehicle, public unitCount: number, public commission: Commission) {
        this.id = id;
        this.date = date;
        this.customer = customer;
        this.vehicle = vehicle;
        this.unitCount = unitCount;
        this.commission = commission;
    }

    getRetro(retroPercentage: number, averageUnits: number): Retro {
        const mini = calculateRetroMini(this.commission.amount, averageUnits, this.unitCount);
        const owed = mini > 0 ? mini - this.commission.amount : 0;
        const payout = mini === 0 ? this.commission.gross * retroPercentage : 0;
        return { mini, owed, payout }
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

class UnitAverage {
    public average: number;
    public rounded: number;
    constructor(public averageUnits: number) {
        this.units = averageUnits;
        this.average = averageUnits / 3;
        this.rounded = Math.round(this.average);
    }
}

function calculateRetroMini(commissionAmount: number, unitAvg: number, dealUnitCount: number): number {
    if (commissionAmount <= 251) {
        if (unitAvg >= 24) return 400 * dealUnitCount;
        if (unitAvg >= 20 && unitAvg < 24) return 350 * dealUnitCount;
        if (unitAvg >= 16 && unitAvg < 20) return 300 * dealUnitCount;
        if (unitAvg >= 12 && unitAvg < 16) return 250 * dealUnitCount;
        if (unitAvg >= 1 && unitAvg < 12) return 200 * dealUnitCount;
        return 0;
    }
    return 0;
}

interface Person {
    id: number;
    name: string;
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

interface Commission {
    fni: number;
    gross: number;
    amount: number;
}

interface Retro {
    mini: number;
    owed: number;
    payout: number;
}
