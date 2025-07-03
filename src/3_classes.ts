class LookupReport {
    public rows: LookupRow[];
    constructor(private data: Array<string | number | boolean>[]) {
        this.rows = [];
        data.map((d, i) => {
            const next = data[i + 1];
            const max = next ? Number(data[i + 1][0]) : 100000000000;
            this.rows.push({ min: Number(d[0]), max, val: Number(d[1]) });
        });
    }
    getValue(query: number) {
        if (this.rows.some(row => query >= row.min && query < row.max)) return this.rows.filter(row => query >= row.min && query < row.max)[0].val;
        return 0;
    }
}

class Units {
    public new: number;
    public used: number;
    constructor() {
        this.new = 0;
        this.used = 0;
    }
    calculatePercent(): object {
        const total = this.getTotal();
        return {
            new: this.new / total,
            used: this.used / total
        };
    }
    getTotal(): number {
        return this.new + this.used;
    }
    getBonus(unitBonusLookup: LookupReport) {
        const total = this.getTotal();
        return unitBonusLookup.getValue(total);
    }
}

class AverageUnits {
    constructor(public threeMonthCount: number) { }
    getAverage(): number {
        return Math.round(this.threeMonthCount / 3);
    }
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

class Retro {
    public mini: number;
    public owed: number;
    public payout: number;
    constructor() {
        this.mini = 0;
        this.owed = 0;
        this.payout = 0;
    }
    calculateMini(retroMiniLookup: LookupReport, dealCommissionAmount: number, employeeAverageUnits: number, dealUnitCount: number) {
        this.mini = dealCommissionAmount <= 251 ? retroMiniLookup.getValue(employeeAverageUnits) * dealUnitCount : 0;
    }
    calculateOwed(dealCommissionAmount: number) {
        this.owed = this.mini > 0 ? this.mini - dealCommissionAmount : 0;
    }
    calculatePayout(dealCommissionGross: number, employeeRetroPercent) {
        this.payout = this.mini === 0 ? dealCommissionGross * employeeRetroPercent : 0;
    }
    getPercent(retroPercentLookup: LookupReport, employeeUnitCount: number): number {
        return retroPercentLookup.getValue(employeeUnitCount);
    }
    getTotal(): number {
        return this.owed + this.payout;
    }
}


class Nps {
    public surveyCount: number;
    public currentPercent: number;
    public averagePercent: number;
    constructor() {
        this.surveyCount = 0;
        this.currentPercent = 0;
        this.averagePercent = 0;
    }
    getBonus(dealUnitCount: number, regionalScore: number): number {
        if(this.surveyCount > 3){
            const outcome = this.getOutcome(regionalScore);
            if(outcome === '3P') return dealUnitCount * 50;
            if(outcome === 'B') return dealUnitCount * -50;
            return 0;
        }
        return dealUnitCount * -50;
    }
    getOutcome(regionalScore: number) {
        const score = (this.currentPercent > this.averagePercent ? this.currentPercent : this.averagePercent) * 100;
        const buffed = regionalScore + 0.03;
        if(score >= buffed) return '3P';
        if(score === regionalScore) return 'A';
        return 'B';
    }
}
