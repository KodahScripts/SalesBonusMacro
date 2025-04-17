
function main(workbook: ExcelScript.Workbook) {
    const initialSheets    : ExcelScript.Worksheet [] = workbook.getWorksheets();
    let employeesWithDeals : Person                [] = [];
    let unitAverages       : UnitAverage           [] = [];
    let priorDraws         : PriorDraw             [] = [];
    let all_spiffs         : Spiff                 [] = [];
    let nps_averages       : NPS                   [] = [];

    initialSheets.forEach(sheet => {
        const reportData = sheet.getUsedRange().getValues();
        const sheetName  = sheet.getName();
        switch (sheetName) {
            case '0432': 
                // console.log("0432 - ", reportData);
                break;
            case '90': 
                let id_90    : number;
                let unitsCol : number;
                reportData.forEach((row, index) => {
                    if(index == 0) {
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
                let id_3213: number;
                let value_3213: number;
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
                let id_spiff: number;
                let value_spiff: number;
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
                let id_nps: number;
                let survey_value: number;
                let survey_percent: number;
                let average_percent: number;
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
    // console.log(nps_averages)
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
    id: string;
    date: number;
    customer: Person;
    vehicle: Vehicle;
    unitCount: number;
    commission: Commission;
}

class Person {
    public deals: Deal[];

    constructor(public id: number, public name: string) {
        this.id = id;
        this.name = name;

        this.deals = [];
    }

    addDeal(deal: Deal) {
        this.deals.push(deal);
    }

    getDealCount() {
        return this.deals.length;
    }
}