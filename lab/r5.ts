class Commission {
    public fni: number;
    public gross: number;
    public amount: number;
    public taxes: number;
    constructor() {
        this.fni = 0;
        this.gross = 0;
        this.amount = 0;
        this.taxes = 0;
    }
}

class NPS {
    public surveys: number;
    public current: number;
    public average: number;
    public outcome: string;
    constructor() {
        this.surveys = 0;
        this.current = 0;
        this.average = 0;
        this.outcome = "";
    }
}

class Bonus {
    public unit: number;
    public topsales: number;
    public csi: number;
    public eom: number;
    public total: number;
    constructor() {
        this.unit = 0;
        this.topsales = 0;
        this.csi = 0;
        this.eom = 0;
        this.total = 0;
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