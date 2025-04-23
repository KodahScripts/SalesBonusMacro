class Store {
    public employees: Employee[];
    public commission: Commission;
    public retro: Retro;
    constructor(public name: string, public abbr: string) {
        this.name = name;
        this.abbr = abbr;
        this.commission = { fni: 0, gross: 0, amount: 0 };
        this.retro = { mini: 0, owed: 0, payout: 0 };
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

    getCommission() {
        this.employees.forEach((employee) => {
            employee.getCommission();
            this.commission.fni += employee.commission.fni;
            this.commission.gross += employee.commission.gross;
            this.commission.amount += employee.commission.amount;
        });
    }

    getRetro() {
        this.employees.forEach((employee) => {
            employee.getRetro();
            this.retro.mini += employee.retro.mini;
            this.retro.owed += employee.retro.owed;
            this.retro.payout += employee.retro.payout;
        });
    }
}

class Employee {
    public deals: Deal[];
    public spiff: number;
    public priorDraw: number;
    public averageUnits: number;
    public nps: NPS;
    public commission: Commission;
    public retro: Retro;
    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;
        this.averageUnits = 0;
        this.priorDraw = 0;
        this.spiff = 0;
        this.commission = { fni: 0, gross: 0, amount: 0 };
        this.retro = { mini: 0, owed: 0, payout: 0 };
        this.deals = [];
    }

    getTotalUnits(): number {
        return this.deals.reduce((acc: number, curr: Deal) => {
            return acc + curr.unitCount;
        }, 0);
    }

    setAverageUnits(ninetyDayUnitCount: number) {
        const average = ninetyDayUnitCount / 3;
        this.averageUnits = Math.round(average);
    }

    getRetroPercentage(): number {
        const unitCount = this.getTotalUnits();
        if (unitCount >= 16) return 0.07;
        if (unitCount >= 12 && unitCount < 16) return 0.04;
        return 0;
    }

    getCommission() {
        this.deals.forEach((deal) => {
            this.commission.fni += deal.commission.fni;
            this.commission.gross += deal.commission.gross;
            this.commission.amount += deal.commission.amount;
        });
    }

    getRetro() {
        this.deals.forEach(deal => {
            const retro = deal.setRetro(this.getRetroPercentage(), this.averageUnits);
            this.retro.mini += retro.mini;
            this.retro.owed += retro.owed;
            this.retro.payout += retro.payout;
        });
    }
}

class Deal {
    public retro: Retro;
    constructor(public id: string, public date: number, public customer: Person, public vehicle: Vehicle, public unitCount: number, public commission: Commission) {
        this.id = id;
        this.date = date;
        this.customer = customer;
        this.vehicle = vehicle;
        this.unitCount = unitCount;
        this.commission = commission;
        this.retro = { mini: 0, owed: 0, payout: 0 };
    }

    setRetro(retroPercentage: number, averageUnits: number): Retro {
        const mini = calculateRetroMini(this.commission.amount, averageUnits, this.unitCount);
        const owed = mini > 0 ? mini - this.commission.amount : 0;
        const payout = mini === 0 ? this.commission.gross * retroPercentage : 0;
        this.retro = { mini: mini, owed: owed, payout: payout };
        return this.retro;
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