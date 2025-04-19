function main(workbook: ExcelScript.Workbook) {
    const initialSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    let all_employees: Person[] = [];
    let unitAverages: UnitAverage[] = [];
    let priorDraws: PriorDraw[] = [];
    let all_spiffs: Spiff[] = [];
    let nps_averages: NPS[] = [];
    let all_deals: Deal[] = [];
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
        switch (sheetName) {
            case '0432':
                let id_0432: number, emp_id_col: number, emp_name_col: number, date_col: number, cust_id_col: number, cust_name_col: number, veh_id_col: number, veh_desc_col: number, sale_type_col: number, comm_fni_col: number, comm_front_col: number, units_0432: number, comm_amount_col: number;
                reportData.forEach((row, index) => {
                    if (index == 0) {
                        id_0432 = row.indexOf("Reference#");
                        emp_id_col = row.indexOf("Salesperson#");
                        emp_name_col = row.indexOf("Salesperson Name");
                        date_col = row.indexOf("Date");
                        cust_id_col = row.indexOf("Customer#");
                        cust_name_col = row.indexOf("Customer Name");
                        veh_id_col = row.indexOf("Stock#");
                        veh_desc_col = row.indexOf("Description");
                        sale_type_col = row.indexOf("Sale Type");
                        comm_fni_col = row.indexOf("COMMBL F&I");
                        comm_front_col = row.indexOf("COMMBL FRONT");
                        units_0432 = row.indexOf("Units");
                        comm_amount_col = row.indexOf("Commission Amount");
                    } else {
                        if (!all_employees.find(emp => emp.id === Number(row[emp_id_col])) && Number(row[emp_id_col]) != 0) all_employees.push({ id: Number(row[emp_id_col]), name: String(row[emp_name_col]) });
                        const cust: Person = {
                            id: Number(row[cust_id_col]),
                            name: String(row[cust_name_col])
                        };
                        const [year, make, model, desc] = String(row[veh_desc_col]).split(',');
                        const veh: Vehicle = {
                            id: String(row[veh_id_col]),
                            year: Number(year),
                            make,
                            model,
                            desc,
                            saleType: String(row[sale_type_col])
                        };
                        const comm: Commission = {
                            fni: Number(row[comm_fni_col]),
                            front: Number(row[comm_front_col]),
                            amount: Number(row[comm_amount_col]),
                            retroMini: 0,
                            retroPayout: 0
                        };
                        const deal: Deal = {
                            empID: Number(row[emp_id_col]),
                            id: String(row[id_0432]),
                            date: Number(row[date_col]),
                            customer: cust,
                            vehicle: veh,
                            unitCount: Number(row[units_0432]),
                            commission: comm
                        };
                        if (deal.unitCount > 0) all_deals.push(deal);
                        deal.vehicle.saleType === 'New' ? store.salesTotals.new += deal.unitCount : store.salesTotals.used += deal.unitCount;
                    }
                });
                break;
            case '90':
                let id_90: number, unitsCol: number;
                reportData.forEach((row, index) => {
                    if (index == 0) {
                        id_90 = row.indexOf("Salesperson#");
                        unitsCol = row.indexOf("units");
                    } else {
                        const employee: UnitAverage = {
                            id: Number(row[id_90]),
                            units: Number(row[unitsCol]),
                            average: Number(row[unitsCol]) / 3,
                            rounded: Math.round(Number(row[unitsCol]) / 3)
                        }
                        unitAverages.push(employee);
                    }
                });
                break;
            case '3213':
                let id_3213: number, value_3213: number;
                reportData.forEach((row, index) => {
                    if (index == 0) {
                        id_3213 = row.indexOf("Control#");
                        value_3213 = row.indexOf("8321D");
                    } else {
                        const employee: PriorDraw = {
                            id: Number(row[id_3213]),
                            amount: Number(row[value_3213]),
                        }
                        priorDraws.push(employee);
                    }
                });
                break;
            case 'SPIFFS':
                let id_spiff: number, value_spiff: number;
                reportData.forEach((row, index) => {
                    if (index == 0) {
                        id_spiff = row.indexOf("Employee #");
                        value_spiff = row.indexOf("Total");
                    } else {
                        const employee: Spiff = {
                            id: Number(row[id_spiff]),
                            amount: Number(row[value_spiff]),
                        }
                        all_spiffs.push(employee);
                    }
                });
                break;
            case 'NPS Sheet':
                let id_nps: number, survey_value: number, survey_percent: number, average_percent: number;
                reportData.forEach((row, index) => {
                    if (index > 1) {
                        if (index == 2) {
                            id_nps = row.indexOf("Employee #");
                            survey_value = row.indexOf("PROMOTER");
                            survey_percent = row.indexOf("NPS%");
                            average_percent = 23;
                        } else {
                            const current = Number(row[survey_percent]);
                            const employee: NPS = {
                                id: Number(row[id_nps]),
                                surveys: Number(row[survey_value]),
                                curr_percent: !current ? 0 : current,
                                avg_percent: Number(row[average_percent])
                            }
                            nps_averages.push(employee);
                        }
                    }
                });
                break;
            case 'Look Up Table':
                // console.log("Look Up Table - ", reportData);
                break;
            default: sheet.delete();
                break;
        }
    });

    all_employees.forEach(emp => {
        const person = emp;
        let spiffs: number, priorDraw: number, averageSoldUnits: UnitAverage;

        try {
            spiffs = all_spiffs.find(sp => sp.id === emp.id).amount;
        } catch {
            spiffs = 0;
        }

        try {
            priorDraw = priorDraws.find(draw => draw.id === emp.id).amount;
        } catch {
            priorDraw = 0;
        }

        try {
            averageSoldUnits = unitAverages.find(unit => unit.id === emp.id);
        } catch {
            averageSoldUnits = {id: person.id, units: 0, average: 0, rounded: 0, }
        }

        const nps = nps_averages.find(n => n.id === emp.id);
        const deals = all_deals.filter(deal => deal.empID === emp.id);
        let unitCount = 0
        deals.forEach(deal => { unitCount += deal.unitCount });

        const commissionTotals: CommissionTotals = { fni: 0, front: 0, amount: 0 };

        deals.forEach(deal => {
            commissionTotals.fni += deal.commission.fni;
            commissionTotals.front += deal.commission.front;
            commissionTotals.amount += deal.commission.amount;
        });

        const employee: Employee = {
            id: person.id,
            name: person.name,
            averageSoldUnits,
            unitCount,
            commissionTotals,
            priorDraw,
            spiffs,
            nps,
            retroPercentage: getRetroPercentage(unitCount),
            deals
        }
        store.employees.push(employee);
    });

    store.employees.forEach(employee => {
        let unitAvg = 0;

        try {
            unitAvg = employee.averageSoldUnits.rounded;
        } catch {
            unitAvg = 0;
        }

        employee.deals.forEach(deal => {
            const mini = calculateRetroMini(deal.commission.amount, unitAvg, deal.unitCount);
            if(mini == 0) {
                deal.commission.retroPayout = calculateRetroPayout(deal.commission.front, employee.retroPercentage);
            }
            deal.commission.retroMini = mini;
        });
    });

    // console.log("STORE", store)
}

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
    commissionTotals: CommissionTotals;
    unitCount: number;
    priorDraw: number;
    spiffs: number;
    nps: NPS;
    retroPercentage: number;
    deals: Deal[];
}

interface Store {
    name: string;
    abbr: string;
    salesTotals: SalesTotals;
    employees: Employee[];
}

interface SalesTotals {
    new: number;
    used: number;
}

interface CommissionTotals {
    fni: number;
    front: number;
    amount: number;
}

function calculateRetroMini(commissionAmount: number, unitAvg: number, dealUnitCount: number): number {
    if (commissionAmount <= 251) {
        if (unitAvg >= 24) return 400 * dealUnitCount;
        if (unitAvg >= 20 && unitAvg < 24) return 350 * dealUnitCount;
        if (unitAvg >= 16 && unitAvg < 20) return 300 * dealUnitCount;
        if (unitAvg >= 12 && unitAvg < 16) return 250 * dealUnitCount;
        if (unitAvg >= 1 && unitAvg < 12) return 200 * dealUnitCount;
        return 0;
    } else {
        return 0;
    }
}

function getRetroPercentage(unitCount: number): number {
    if (unitCount >= 16) return 0.07;
    if (unitCount >= 12 && unitCount < 16) return 0.04;
    return 0;
}

function calculateRetroPayout(commissionFront: number, retroPercentage: number): number {
    return commissionFront * retroPercentage;
}

function calculateRetroTotal(retroPayout: number, retroOwed: number): number {
    return retroPayout + retroOwed;
}