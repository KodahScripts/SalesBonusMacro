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

interface NPS {
    surveys: number;
    current: number;
    average: number;
}