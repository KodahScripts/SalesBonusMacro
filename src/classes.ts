class Store {
    public employees: Employee[];
    public regionalScore: number;
    public commission: Commission;
    public retro: Retro;
    public fni: FnI;
    public bonus: Bonus;
    public topSalesman: TopSalesman;
    public totalCommission: number;
    public ytdBucket: number;
    public accounts: Account;
    public units: Units;
    constructor(public name: string, public abbr: string) {
        this.name = name;
        this.abbr = abbr;
        this.regionalScore = 0;
        this.totalCommission = 0;
        this.ytdBucket = 0;
        this.units = { new: 0, used: 0, total: 0 };
        this.accounts = { retro: '', expense1: '', expense2: '', salesTax: '', salesBonusTax1: '', salesBonusTax2: '' }
        this.fni = { reserve: 0, gross: 0, payout: 0 }
        this.commission = { fni: 0, gross: 0, amount: 0 };
        this.retro = { mini: 0, owed: 0, payout: 0, total: 0 };
        this.bonus = { unit: 0, topsales: 0, csi: 0, eom: 0, total: 0 };
        this.topSalesman = { id: 0, count: 0 };
        this.employees = [];
    }

    employeeExists(employeeId: number): boolean {
        return this.employees.some(emp => emp.id === employeeId);
    }

    calculateAll() {
        this.getTotalUnits();
        this.getTopSalesman();
        this.getCommission();
        this.getRetro();
        this.getFni();
        this.getBonus();
    }

    getTotalUnits() {
        this.employees.forEach((employee) => {
            employee.getTotalUnits();
            this.units.new += employee.units.new;
            this.units.used += employee.units.used;
            this.units.total += employee.units.total;
        });
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
            this.retro.total += employee.retro.total;
        });
    }

    getFni() {
        this.employees.forEach((employee) => {
            employee.getFni();
            this.fni.reserve += employee.fni.reserve;
            this.fni.gross += employee.fni.gross;
            this.fni.payout += employee.fni.payout;
        });
    }

    getBonus() {
        this.employees.forEach((employee) => {
            employee.getBonus();
            this.bonus.unit += employee.bonus.unit;
            this.bonus.topsales += employee.bonus.topsales;
            this.bonus.csi += employee.bonus.csi;
            this.bonus.eom += employee.bonus.eom;
            this.bonus.total += employee.bonus.total;
            this.totalCommission += employee.totalCommission;
            this.ytdBucket += employee.ytdBucket;
        });
    }

    getTopSalesman() {
        this.employees.forEach((employee) => {
            if(employee.units.total > this.topSalesman.count) this.topSalesman = { id: employee.id, count: employee.units.total };
        });
        this.employees.filter(emp => emp.id === this.topSalesman.id)[0].bonus.topsales = 500;
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
    public fni: FnI;
    public bonus: Bonus;
    public totalCommission: number;
    public commissionBalance: number;
    public ytdBucket: number;
    public drawAmount: number;
    public units: Units;
    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;
        this.averageUnits = 0;
        this.priorDraw = 0;
        this.drawAmount = 0;
        this.spiff = 0;
        this.totalCommission = 0;
        this.ytdBucket = 0;
        this.commissionBalance = 0;
        this.units = { new: 0, used: 0, total: 0 };
        this.nps = { surveys: 0, current: 0, average: 0, outcome: "B" };
        this.commission = { fni: 0, gross: 0, amount: 0 };
        this.retro = { mini: 0, owed: 0, payout: 0, total: 0 };
        this.fni = { reserve: 0, gross: 0, payout: 0 };
        this.bonus = { unit: 0, topsales: 0, csi: 0, eom: 0, total: 0 };
        this.deals = [];
    }

    getTotalUnits() {
        this.deals.forEach(deal => {
            if(deal.vehicle.saleType === "New") {
                this.units.new += deal.unitCount;
            } else {
                this.units.used += deal.unitCount;
            }
            this.units.total += deal.unitCount;
        });
    }

    setAverageUnits(ninetyDayUnitCount: number) {
        const average = ninetyDayUnitCount / 3;
        this.averageUnits = Math.round(average);
    }

    getRetroPercentage(): number {
        if (this.units.total >= 16) return 0.07;
        if (this.units.total >= 12 && this.units.total < 16) return 0.04;
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
            this.retro.total += retro.total;
        });
    }

    getFni() {
        const reserve = this.commission.fni * 0.25;
        const gross = this.commission.fni - reserve;
        const payout = gross * 0.05;
        this.fni = { reserve, gross, payout };
    }

    getBonus() {
        const unitBonus = calculateUnitBonus(this.units.total);
        const csiBonus = caclulateCsiBonus(this.nps.surveys, this.nps.outcome, this.units.total);
        this.bonus.unit = unitBonus;
        this.bonus.csi = csiBonus;
        this.bonus.total = unitBonus + csiBonus + this.bonus.topsales;
        this.bonus.eom = this.retro.total + this.fni.payout + this.bonus.total + this.spiff;
        this.totalCommission = this.bonus.eom + this.commission.amount;
        this.ytdBucket = calculateYtdBucket(this.totalCommission, this.priorDraw, this.spiff);
        this.drawAmount = this.commission.amount + this.bonus.total >= this.priorDraw ? this.priorDraw : this.commission.amount + this.bonus.total;
    }

    getOwed(): number {
        return this.totalCommission - this.priorDraw;
    }

    getDepositGross(): number {
        return this.totalCommission - this.priorDraw - this.ytdBucket;
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
        this.retro = { mini: 0, owed: 0, payout: 0, total: 0 };
    }

    setRetro(retroPercentage: number, averageUnits: number): Retro {
        const mini = calculateRetroMini(this.commission.amount, averageUnits, this.unitCount);
        const owed = mini > 0 ? mini - this.commission.amount : 0;
        const payout = mini === 0 ? this.commission.gross * retroPercentage : 0;
        const total = payout + owed;
        this.retro = { mini: mini, owed: owed, payout: payout, total: total };
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