interface Person {
    id: number;
    name: string;
}

interface Units {
    new: number;
    used: number;
    total: number;
}

interface UnitPercent {
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
    csi: number;
    eom: number;
    total: number;
}

interface Commission {
    fni: number;
    gross: number;
    amount: number;
    taxes: number;
}

interface Retro {
    mini: number;
    owed: number;
    payout: number;
    total: number;
}

interface NPS {
    surveys: number;
    current: number;
    average: number;
    outcome: string;
}

interface Account {
    retro: string;
    expense: {
        one: string,
        two: string
    };
    salesTax: string;
    salesBonusTax1: string;
    salesBonusTax2: string;
}

interface Expense {
    one: number,
    two: number
}