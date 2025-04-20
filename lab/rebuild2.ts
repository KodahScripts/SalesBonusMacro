function main(workbook: ExcelScript.Workbook) {
    const initialSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    let all_employees: Person[] = [];
    let unitAverages: UnitAverage[] = [];
    let priorDraws: PriorDraw[] = [];
    let all_spiffs: Spiff[] = [];
    let nps_averages: NPS[] = [];
    let all_deals: Deal[] = [];
    const store: Store = {
        name: STORE_NAME,
        abbr: STORE_ABBR,
        salesTotals: {
            new: 0,
            used: 0
        },
        topSalesman: {
            id: 0,
            count: 0
        },
        employees: []
    };

    const allowedSheets = ['0432', '90', '3213', 'SPIFFS', 'NPS Sheet', 'Look Up Table', 'Input'];
    initialSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if(!allowedSheets.includes(sheetName)) {
            sheet.delete();
        } else {
            const reportData = sheet.getUsedRange().getValues();
            switch(sheetName) {

            }
        }

    });
}



interface Store {
    name: string;
    abbr: string;
    salesTotals: SalesTotals;
    topSalesman: TopSalesman;
    employees: Employee[];
}

interface Report {
    store: Store;
    columnNames: String[];
}

interface Deal {
    employeeId: {
        id: string;
        date: number;
        customer: Customer;
        vehicle: Vehicle;
        unitCount: number;
        commission: Commission;
    }
}

interface UnitAverage {
    employeeId: {
        units: number;
        average: number;
        rounded: number;
    }
}

interface NPS {
    employeeId: {
        surveys: number;
        current: number;
        average: number;
    }
}

interface Person {
    id: number;
    name: string;
}

interface Commission {
    fni: number;
    front: number;
    amount: number;
    retro: Retro;
}

interface Retro {
    mini: number;
    owed: number;
    payout: number;
}