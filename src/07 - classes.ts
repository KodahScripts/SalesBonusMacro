interface UnitAverage {
    id: number;
    units: number;
    average: number;
    rounded: number;
}

interface PriorDraw {
    id: number;
    amount: number;
}

interface Spiff {
    id: number;
    amount: number;
}

interface NPS {
    id: number;
    surveys: number;
    curr_percent: number;
    avg_percent: number;
}

interface Commission {
    fni: number;
    front: number;
    amount: number;
    retroMini: number;
    retroOwed: number;
    retroPayout: number;
}

interface Vehicle {
    id: string;
    year: number;
    make: string;
    model: string;
    desc: string;
    saleType: string;
}

interface Deal {
    empID: number;
    id: string;
    date: number;
    customer: Person;
    vehicle: Vehicle;
    unitCount: number;
    commission: Commission;
}

interface Person {
    id: number;
    name: string;
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

interface Store {
    name: string;
    abbr: string;
    salesTotals: SalesTotals;
    topSalesman: TopSalesman;
    employees: Employee[];
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
}